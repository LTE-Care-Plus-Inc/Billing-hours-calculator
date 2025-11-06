#!/usr/bin/env python3
import io
import os
import re
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode
from st_aggrid.shared import JsCode

from core import process_file, per_staff_per_day, DATE_COL, START_COL, END_COL, _fmt_time, load_input


APP_TITLE = "Billable Hours Aggregator"


def _apply_theme(theme: str):
    theme = (theme or "light").lower()
    if theme not in ("light", "dark"):
        theme = "light"
    # Basic CSS override to emulate light/dark switching
    if theme == "dark":
        bg = "#1b1d20"; text = "#ffffff"; sbg = "#2b2d31"
    else:
        bg = "#ffffff"; text = "#000000"; sbg = "#f5f7fb"
    st.markdown(
        f"""
        <style>
        html, body, [data-testid="stAppViewContainer"] {{
            background-color: {bg} !important;
            color: {text} !important;
        }}
        [data-testid="stSidebar"] {{
            background-color: {sbg} !important;
        }}
        /* tweak dataframes background */
        .stDataFrame [role="grid"] {{
            background-color: {bg} !important;
            color: {text} !important;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def _process_uploaded_file(uploaded_file) -> tuple[pd.DataFrame, pd.DataFrame]:
    suffix = Path(uploaded_file.name).suffix.lower()
    # Reuse existing process_file by writing to a temp file for robustness
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded_file.getbuffer())
        tmp_path = Path(tmp.name)
    try:
        summary, details = process_file(tmp_path)
    finally:
        try:
            os.unlink(tmp_path)
        except Exception:
            pass
    return summary, details


def _bytes_csv(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


def _bytes_excel(sheets: dict[str, pd.DataFrame]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xl:
        for name, df in sheets.items():
            # Excel sheet names are limited to 31 chars
            xl_name = name[:31]
            df.to_excel(xl, index=False, sheet_name=xl_name)
    buf.seek(0)
    return buf.read()


def _norm_name(s: str) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    raw = str(s).strip()
    # Handle "Last, First" → "first last"
    if "," in raw:
        parts = [p.strip() for p in raw.split(",") if p is not None]
        if len(parts) >= 2:
            first = parts[1].split()[0] if parts[1] else ""
            last = parts[0]
            raw = f"{first} {last}".strip()
    # Remove punctuation that commonly differs between sources
    raw = re.sub(r"[\.,\-_/\\]+", " ", raw)
    # Collapse whitespace and lowercase
    return " ".join(raw.split()).lower()


def _coerce_rate(x):
    if pd.isna(x):
        return None
    try:
        s = str(x).strip()
        # Keep digits, decimal point, and optional leading minus
        s = re.sub(r"[^0-9.\-]", "", s)
        return float(s) if s else None
    except Exception:
        return None

def _fmt_usd(x):
    try:
        if x is None or pd.isna(x):
            return "missing pay rate"
    except Exception:
        pass
    try:
        return f"${float(x):,.2f}"
    except Exception:
        return "missing pay rate"

def _escape_dollar(text: str) -> str:
    # Replace dollar signs with HTML entity to avoid MathJax/Markdown interpretation
    try:
        return text.replace("$", "&#36;") if isinstance(text, str) else text
    except Exception:
        return text

def _style_pay_df(df: pd.DataFrame):
    def red_missing(v):
        try:
            if v is None or pd.isna(v):
                return 'color: red'
        except Exception:
            pass
        if isinstance(v, str) and v.strip().lower() == 'missing pay rate':
            return 'color: red'
        return ''
    sty = df.style.format({"Rate": _fmt_usd, "Total Pay": _fmt_usd})
    try:
        sty = sty.applymap(red_missing, subset=["Rate", "Total Pay"])  # pandas <2.1
    except Exception:
        sty = sty.map(red_missing, subset=["Rate", "Total Pay"])      # pandas >=2.1
    return sty


def _process_payroll_file(uploaded_file) -> pd.DataFrame:
    suffix = Path(uploaded_file.name).suffix.lower()
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded_file.getbuffer())
        tmp_path = Path(tmp.name)
    try:
        df = load_input(tmp_path)
    finally:
        try:
            os.unlink(tmp_path)
        except Exception:
            pass
    # Normalize column names for robust access
    cols = {c.strip().lower(): c for c in df.columns}
    # Required columns
    staff_col = cols.get("staff name".lower()) or cols.get("staff name") or "Staff Name"
    rate_col = cols.get("rate".lower()) or cols.get("rate") or "Rate"
    job_col = cols.get("job title".lower()) or cols.get("job title") or "Job Title"
    if staff_col not in df.columns or rate_col not in df.columns:
        raise ValueError("Payroll report must include 'Staff Name' and 'Rate' columns")

    # Filter out BCBA job titles where applicable
    if job_col in df.columns:
        mask = ~df[job_col].astype(str).str.contains("bcba", case=False, na=False)
        df = df[mask].copy()
    else:
        df = df.copy()

    # Build normalized key and clean rate
    df["__name_key__"] = df[staff_col].apply(_norm_name)
    df["__rate__"] = df[rate_col].apply(_coerce_rate)
    # Deduplicate by name key, keeping the first non-null rate occurrence
    df = df.sort_index()
    df = df.dropna(subset=["__name_key__"])  # drop empty names
    # Prefer rows where rate is present
    df_nonnull = df.dropna(subset=["__rate__"]).drop_duplicates(subset=["__name_key__"], keep="first")
    df_null = df[df["__rate__"].isna()].drop_duplicates(subset=["__name_key__"], keep="first")
    df_merged = pd.concat([df_nonnull, df_null[~df_null["__name_key__"].isin(df_nonnull["__name_key__"])]] , ignore_index=True)
    return df_merged[["__name_key__", "__rate__"]]


def _filter_and_sort(df: pd.DataFrame, query: str, sort_by_minutes: bool) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    out = df.copy()
    q = (query or "").strip().lower()
    if q:
        out = out[out["Staff Name"].str.lower().str.contains(q, na=False)]
    if sort_by_minutes:
        out = out.sort_values(["Total_Minutes", "Staff Name"], ascending=[False, True])
    else:
        out = out.sort_values(["Staff Name"], ascending=[True])
    return out


def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)

    with st.sidebar:
        st.header("Input")
        uploaded = st.file_uploader(
            "Billable Hours Report",
            type=["csv", "xlsx", "xlsm", "xltx", "xltm"],
            accept_multiple_files=False,
        )

        payroll_uploaded = st.file_uploader(
            "Payroll Report",
            type=["csv", "xlsx", "xlsm", "xltx", "xltm"],
            accept_multiple_files=False,
            key="payroll_report_uploader",
            help="CSV/Excel with 'Staff Name', 'Rate', and optional 'Job Title'. 'BCBA' titles are ignored.",
        )

        use_sample = st.button("Use bundled sample data")

        st.header("Theme")
        if "theme" not in st.session_state:
            st.session_state.theme = "light"
        theme_choice = st.radio("Mode", ["Light", "Dark"], index=(1 if st.session_state.theme=="dark" else 0))
        st.session_state.theme = theme_choice.lower()

    _apply_theme(st.session_state.theme)
    # Additional UI polish and higher-contrast dark mode (no functional changes)
    _t = (st.session_state.get("theme") or "light").lower()
    if _t == "dark":
        _bg = "#0f1419"; _text = "#e6edf3"; _sbg = "#161b22"; _input = "#1b222c"; _border = "#2d333b"; _link = "#58a6ff"; _primary = "#2C7BE5"
    else:
        _bg = "#ffffff"; _text = "#0b1220"; _sbg = "#f5f7fb"; _input = "#ffffff"; _border = "#d8dee4"; _link = "#1f6feb"; _primary = "#2C7BE5"
    st.markdown(
        f"""
        <style>
        /* Normalize fonts app-wide */
        :root, html, body, [data-testid="stAppViewContainer"], [data-testid="stSidebar"],
        .stMarkdown, .stButton > button, [data-baseweb="input"], [data-baseweb="select"], textarea,
        details[data-testid="stExpander"] > summary, .stDataFrame, .stDataFrame * {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, "Noto Sans", "Liberation Sans", "DejaVu Sans", Ubuntu, Cantarell, "Fira Sans", "Droid Sans", sans-serif !important;
            font-style: normal !important;
            font-weight: 500;
            font-variant-ligatures: none !important;
            letter-spacing: 0 !important;
            white-space: normal;
        }}
        /* Ensure no italics appear in expander headers */
        details[data-testid="stExpander"] > summary, 
        details[data-testid="stExpander"] > summary *,
        details[data-testid="stExpander"] > summary em,
        details[data-testid="stExpander"] > summary i {{
            font-style: normal !important;
            font-weight: 500 !important;
            font-variant-ligatures: none !important;
            letter-spacing: 0 !important;
            white-space: normal !important;
        }}
        /* Soften section headers */
        section.main > div > div h2, section.main > div > div h3, section.main > div > div h4 {{
            letter-spacing: .2px;
        }}
        /* Inputs and selects */
        [data-baseweb="input"], [data-baseweb="select"], textarea {{
            background-color: {_input} !important;
            color: {_text} !important;
            border-color: {_border} !important;
        }}
        [data-baseweb="input"] input, textarea {{ color: {_text} !important; }}
        [data-baseweb="select"] * {{ color: {_text} !important; }}
        /* Placeholders */
        input::placeholder, textarea::placeholder, [data-baseweb="input"] input::placeholder {{
            color: rgba(230,237,243,0.60) !important;
        }}
        [data-baseweb="input"]:focus-within, [data-baseweb="select"]:focus-within, textarea:focus {{
            box-shadow: 0 0 0 1px {_primary} inset !important;
            border-color: {_primary} !important;
        }}
        /* Buttons */
        .stButton > button {{
            background-color: {_primary} !important;
            color: #ffffff !important;
            border: 1px solid {_primary} !important;
        }}
        .stButton > button:hover {{ filter: brightness(1.06); }}
        /* Expanders */
        details[data-testid="stExpander"] {{
            background: {_sbg};
            border: 1px solid {_border};
            border-radius: 10px;
            overflow: hidden;
            margin-bottom: 10px;
        }}
        details[data-testid="stExpander"] > summary {{
            padding-top: 8px; padding-bottom: 8px; color: {_text};
        }}
        /* Captions spacing + color (make darker for readability) */
        p, .stCaption, [data-testid="stCaptionContainer"] {{ margin-top: 0.25rem; }}
        [data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] * , .stCaption {{
            color: {_text} !important;
            opacity: 0.92 !important;
            font-weight: 500;
        }}
        [data-testid="stCaptionContainer"] * {{ color: rgba(230,237,243,0.75) !important; }}
        /* Reduce extra vertical gaps around controls */
        div.row-widget.stButton {{ margin-top: .25rem; }}
        /* Dataframes */
        .stDataFrame [role="grid"] {{ background-color: {_bg} !important; color: {_text} !important; }}
        .stDataFrame [role="columnheader"] {{ background-color: {_sbg} !important; }}
        /* AgGrid shells */
        .ag-theme-balham, .ag-theme-balham-dark {{ border-radius: 8px; background-color: {_sbg}; border: 1px solid {_border}; }}
        .ag-theme-balham-dark .ag-header, .ag-theme-balham .ag-header {{ background-color: {_sbg}; border-bottom: 1px solid {_border}; }}
        /* AgGrid rows and cells for dark mode */
        .ag-theme-balham-dark .ag-root-wrapper, .ag-theme-balham-dark .ag-center-cols-viewport {{ background-color: {_bg} !important; }}
        .ag-theme-balham-dark .ag-row {{ background-color: {_bg} !important; color: {_text} !important; }}
        .ag-theme-balham-dark .ag-row-odd {{ background-color: #121821 !important; }}
        .ag-theme-balham-dark .ag-row-hover {{ background-color: #18212b !important; }}
        .ag-theme-balham-dark .ag-row-selected {{ background-color: #203044 !important; }}
        .ag-theme-balham-dark .ag-cell {{ border-color: {_border} !important; color: {_text} !important; }}
        .ag-theme-balham-dark .ag-checkbox-input-wrapper input:checked + span {{ background-color: {_primary} !important; border-color: {_primary} !important; }}
        /* Tabs */
        [role="tablist"] button[role="tab"] {{ color: {_text}; }}
        [role="tablist"] button[role="tab"][aria-selected="true"] {{ border-bottom: 2px solid {_primary}; }}
        /* Links */
        a {{ color: {_link} !important; }}
        </style>
        """,
        unsafe_allow_html=True,
    )

    summary_df = None
    details_df = None
    source_label = None

    if uploaded is not None:
        try:
            summary_df, details_df = _process_uploaded_file(uploaded)
            source_label = f"Uploaded: {uploaded.name}"
        except Exception as e:
            st.error(f"Failed to process uploaded file: {e}")

    elif use_sample:
        sample_path = Path(__file__).parent / "sample_data.xlsx"
        if not sample_path.exists():
            st.warning("Sample file not found next to the app.")
        else:
            try:
                summary_df, details_df = process_file(sample_path)
                source_label = f"Sample: {sample_path.name}"
            except Exception as e:
                st.error(f"Failed to process sample file: {e}")

    if summary_df is None or details_df is None:
        st.info("Upload a file or click 'Use bundled sample data' in the sidebar.")
        st.stop()

    # Controls
    st.subheader("Summary")
    col1, col2, col3 = st.columns([2, 2, 3])
    with col1:
        query = st.text_input("Search staff", placeholder="Type a name…")
    with col2:
        sort_choice = st.selectbox("Sort by", ["Staff Name", "Minutes (desc)"])
    with col3:
        if source_label:
            st.caption(source_label)

    sort_by_minutes = sort_choice.lower().startswith("minutes")
    filtered_df = _filter_and_sort(summary_df, query, sort_by_minutes)

    # Totals footer
    if filtered_df is not None and not filtered_df.empty:
        mins = int(filtered_df["Total_Minutes"].sum())
        hours = round(mins / 60.0, 2)
        units = int(round(mins / 15.0))
        st.caption(f"Totals (filtered): {mins} minutes | {hours:.2f} hours | {units} units")

    # Optional: load payroll mapping (after we have summary/details)
    payroll_map = None
    payroll_err = None
    if payroll_uploaded is not None:
        try:
            payroll_df = _process_payroll_file(payroll_uploaded)
            # Build mapping from normalized name to rate
            payroll_map = {r["__name_key__"]: r["__rate__"] for _, r in payroll_df.iterrows()}
        except Exception as e:
            payroll_err = str(e)

    # Summary view: interactive grid (row-click) and inline expanders
    if filtered_df is None:
        st.dataframe(pd.DataFrame(), use_container_width=True, hide_index=True)
    elif filtered_df.empty:
        st.dataframe(filtered_df, use_container_width=True, hide_index=True)
    else:
        tab_interactive, tab_inline = st.tabs(["Interactive Table", "Inline Breakdown"])
        with tab_interactive:
            # Build a copy for the grid and hide any internal id column if present
            df_grid = filtered_df.copy()
            if "::auto_unique_id::" in df_grid.columns:
                df_grid = df_grid.drop(columns=["::auto_unique_id::"])

            # If payroll loaded, add Rate and Total Pay columns (interactive table only)
            if payroll_err:
                st.warning(f"Payroll report error: {payroll_err}")
            if payroll_map is not None:
                name_keys = df_grid["Staff Name"].apply(_norm_name)
                rates = name_keys.map(lambda k: payroll_map.get(k))
                rates = pd.to_numeric(rates, errors='coerce')
                df_grid["Rate"] = rates
                # Compute Total Pay based on Total_Hours * Rate (coerce to numeric first)
                if "Total_Hours" in df_grid.columns:
                    th = pd.to_numeric(df_grid["Total_Hours"], errors='coerce')
                    df_grid["Total Pay"] = (th * rates).round(2)
                else:
                    df_grid["Total Pay"] = pd.NA

                # Warn if any visible staff are missing a rate
                missing = sorted(set(df_grid.loc[df_grid["Rate"].isna(), "Staff Name"].astype(str)))
                if missing:
                    miss_list = ", ".join(missing[:10]) + (" ..." if len(missing) > 10 else "")
                    st.warning(f"Missing rates for: {miss_list}")
                # Show match coverage to help verify mapping
                try:
                    matched = int(rates.notna().sum())
                    total = int(len(rates))
                    st.caption(f"Payroll matches: {matched}/{total}")
                except Exception:
                    pass

            gob = GridOptionsBuilder.from_dataframe(df_grid)
            gob.configure_selection(selection_mode="multiple", use_checkbox=True)
            gob.configure_grid_options(rowSelection="multiple", suppressRowClickSelection=False)
            gob.configure_default_column(sortable=True, filter=True, resizable=True)
            # Column formatting
            if "Total_Minutes" in df_grid.columns:
                gob.configure_column("Total_Minutes", header_name="Total Minutes", type=["numericColumn","numberColumnFilter","customNumericFormat"], precision=0, minWidth=140)
            if "Total_Hours" in df_grid.columns:
                gob.configure_column("Total_Hours", header_name="Total Hours", type=["numericColumn","numberColumnFilter","customNumericFormat"], precision=2, minWidth=120)
            if "Total_Units" in df_grid.columns:
                gob.configure_column("Total_Units", header_name="Total Units", type=["numericColumn","numberColumnFilter","customNumericFormat"], precision=0, minWidth=120)

            # Currency formatting for Rate and Total Pay when present
            usd_fmt = JsCode("""
                function(params) {
                    if (params.value === null || params.value === undefined || isNaN(params.value)) { return 'missing pay rate'; }
                    try { return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(params.value); }
                    catch(e) { return '$' + Number(params.value).toFixed(2); }
                }
            """)
            # Style missing values in red via cellStyle
            missing_red = JsCode("""
                function(params) {
                    var v = params.value;
                    if (v === null || v === undefined || (typeof v === 'number' && isNaN(v))) {
                        return {color: 'red'};
                    }
                    if (typeof v === 'string' && v.toLowerCase() === 'missing pay rate') {
                        return {color: 'red'};
                    }
                    return {};
                }
            """)
            if "Rate" in df_grid.columns:
                gob.configure_column("Rate", header_name="Rate", type=["numericColumn","numberColumnFilter","customNumericFormat"], minWidth=120, valueFormatter=usd_fmt, cellStyle=missing_red)
            if "Total Pay" in df_grid.columns:
                gob.configure_column("Total Pay", header_name="Total Pay", type=["numericColumn","numberColumnFilter","customNumericFormat"], minWidth=140, valueFormatter=usd_fmt, cellStyle=missing_red)
            gob.configure_pagination(enabled=True, paginationAutoPageSize=False, paginationPageSize=25)
            grid_options = gob.build()

            ag_theme = "balham-dark" if st.session_state.get("theme") == "dark" else "balham"
            grid_response = AgGrid(
                df_grid,
                gridOptions=grid_options,
                height=400,
                data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
                update_mode=GridUpdateMode.SELECTION_CHANGED,
                allow_unsafe_jscode=True,
                fit_columns_on_grid_load=True,
                theme=ag_theme,
            )

            selected_rows = grid_response.get("selected_rows", [])
            # Normalize to a list of dicts across AgGrid versions
            if isinstance(selected_rows, pd.DataFrame):
                selected_rows = selected_rows.to_dict("records")
            elif not isinstance(selected_rows, list):
                selected_rows = []

            if len(selected_rows) > 0:
                names = []
                for r in selected_rows:
                    nm = str(r.get("Staff Name", "")).strip()
                    if nm:
                        names.append(nm)
                names = list(dict.fromkeys(names))  # de-duplicate, preserve order
                st.caption(f"Selected rows: {len(names)}")

                # Expand/collapse controls for interactive section
                if "interactive_expand_all" not in st.session_state:
                    st.session_state.interactive_expand_all = False
                ec1, ec2 = st.columns([1,1])
                with ec1:
                    if st.button("Expand all", key="btn_expand_interactive"):
                        st.session_state.interactive_expand_all = True
                with ec2:
                    if st.button("Collapse all", key="btn_collapse_interactive"):
                        st.session_state.interactive_expand_all = False
                for name in names:
                    # Build expander title with totals and optional total pay
                    _mins = None
                    if filtered_df is not None and not filtered_df.empty:
                        try:
                            _row = filtered_df[filtered_df["Staff Name"] == name].iloc[0]
                            _mins = int(_row.get("Total_Minutes", 0) or 0)
                        except Exception:
                            _mins = None
                    if _mins is None:
                        try:
                            _mins = int(details_df.loc[details_df["Staff Name"] == name, "Rounded Minutes"].sum())
                        except Exception:
                            _mins = 0
                    _hours = _mins / 60.0
                    _units = int(round(_mins / 15.0))
                    _title_pay = ""
                    try:
                        _rate = payroll_map.get(_norm_name(name)) if payroll_map is not None else None
                    except Exception:
                        _rate = None
                    if _rate is not None:
                        try:
                            _title_pay_val = round(float(_hours) * float(_rate), 2)
                            _title_pay = f" | {_fmt_usd(_title_pay_val)}"
                        except Exception:
                            _title_pay = ""
                    elif payroll_map is not None:
                        _title_pay = " | missing pay rate"
                    exp_title = f"{name} | {_mins} min | {_hours:.2f} h | {_units} units{_title_pay}"
                    with st.expander(exp_title, expanded=bool(st.session_state.get("interactive_expand_all", False))):
                        try:
                            daily = per_staff_per_day(details_df, name)
                        except Exception as e:
                            daily = pd.DataFrame()
                            st.error(f"Failed to build per-day breakdown for {name}: {e}")
                        if payroll_map is not None:
                            rate = payroll_map.get(_norm_name(name))
                            if rate is not None:
                                daily["Rate"] = rate
                                daily["Total Pay"] = (daily.get("Total_Hours", 0) * rate).round(2)
                            else:
                                daily["Rate"] = pd.NA
                                daily["Total Pay"] = pd.NA
                            st.dataframe(_style_pay_df(daily), use_container_width=True, hide_index=True)
                        else:
                            st.dataframe(daily, use_container_width=True, hide_index=True)

                        c1, c2 = st.columns(2)
                        with c1:
                            st.download_button(
                                label=f"Download {name} per-day CSV",
                                data=_bytes_csv(daily),
                                file_name=f"{name}_by_day.csv",
                                mime="text/csv",
                                disabled=daily.empty,
                                key=f"dl_perday_interactive_{name}",
                            )
                        with c2:
                            staff_rows = details_df[details_df["Staff Name"] == name]
                            st.download_button(
                                label=f"Download {name} details CSV",
                                data=_bytes_csv(staff_rows),
                                file_name=f"{name}_details.csv",
                                mime="text/csv",
                                disabled=staff_rows.empty,
                                key=f"dl_details_interactive_{name}",
                            )

        with tab_inline:
            records = filtered_df.to_dict("records")
            # Expand/collapse controls for inline breakdown
            if "inline_expand_all" not in st.session_state:
                st.session_state.inline_expand_all = False
            ic1, ic2 = st.columns([1,1])
            with ic1:
                if st.button("Expand all", key="btn_expand_inline"):
                    st.session_state.inline_expand_all = True
            with ic2:
                if st.button("Collapse all", key="btn_collapse_inline"):
                    st.session_state.inline_expand_all = False
            for rec in records:
                name = str(rec.get("Staff Name", ""))
                mins = int(rec.get("Total_Minutes", 0) or 0)
                hours = mins / 60.0
                units = int(round(mins / 15.0))
                title = f"{name} — {mins} min | {hours:.2f} h | {units} units"
                # Build title with Total Pay and Completed pay buckets
                __rate = payroll_map.get(_norm_name(name)) if payroll_map is not None else None
                if __rate is not None:
                    try:
                        _daily_title = per_staff_per_day(details_df, name)
                    except Exception:
                        _daily_title = pd.DataFrame(columns=["Total_Hours", "Completed"])
                    yes_hours = pd.to_numeric(_daily_title.loc[_daily_title.get("Completed", pd.Series()).eq("Yes"), "Total_Hours"], errors='coerce').fillna(0).sum()
                    no_hours  = pd.to_numeric(_daily_title.loc[_daily_title.get("Completed", pd.Series()).eq("No"),  "Total_Hours"], errors='coerce').fillna(0).sum()
                    tot_hours = pd.to_numeric(_daily_title.get("Total_Hours", pd.Series()), errors='coerce').fillna(0).sum()
                    total_pay = _escape_dollar(_fmt_usd(round(float(tot_hours) * float(__rate), 2)))
                    yes_pay   = _escape_dollar(_fmt_usd(round(float(yes_hours) * float(__rate), 2)))
                    no_pay    = _escape_dollar(_fmt_usd(round(float(no_hours) * float(__rate), 2)))
                    __suffix = f" | Total Pay:\u00A0{total_pay} | Completed Pay:\u00A0{yes_pay} | Completed No Pay:\u00A0{no_pay}"
                elif payroll_map is not None:
                    __suffix = " | missing pay rate"
                else:
                    __suffix = ""
                title = f"{name} | {mins} min | {hours:.2f} h | {units} units{__suffix}"
                title = _escape_dollar(title)
                with st.expander(title, expanded=bool(st.session_state.get("inline_expand_all", False))):
                    try:
                        daily = per_staff_per_day(details_df, name)
                    except Exception as e:
                        daily = pd.DataFrame()
                        st.error(f"Failed to build per-day breakdown for {name}: {e}")
                    if payroll_map is not None:
                        rate = payroll_map.get(_norm_name(name))
                        if rate is not None:
                            daily["Rate"] = rate
                            daily["Total Pay"] = (daily.get("Total_Hours", 0) * rate).round(2)
                        else:
                            daily["Rate"] = pd.NA
                            daily["Total Pay"] = pd.NA
                        st.dataframe(_style_pay_df(daily), use_container_width=True, hide_index=True)
                    else:
                        st.dataframe(daily, use_container_width=True, hide_index=True)

                    c1, c2 = st.columns(2)
                    with c1:
                        st.download_button(
                            label=f"Download {name} per-day CSV",
                            data=_bytes_csv(daily),
                            file_name=f"{name}_by_day.csv",
                            mime="text/csv",
                            disabled=daily.empty,
                            key=f"dl_perday_inline_{name}",
                        )
                    with c2:
                        staff_rows = details_df[details_df["Staff Name"] == name]
                        st.download_button(
                            label=f"Download {name} details CSV",
                            data=_bytes_csv(staff_rows),
                            file_name=f"{name}_details.csv",
                            mime="text/csv",
                            disabled=staff_rows.empty,
                            key=f"dl_details_inline_{name}",
                        )

    # Selection for per-staff breakdown
    st.subheader("Per-Staff Per-Day Breakdown")

    # Build full staff list from the filtered summary (fallback to details if needed)
    if filtered_df is not None and not filtered_df.empty:
        all_staff = sorted(set(filtered_df["Staff Name"].astype(str)))
    else:
        all_staff = sorted(set(details_df.get("Staff Name", pd.Series(dtype=str)).astype(str)))

    # Add a dedicated search box to filter staff options
    staff_search = st.text_input("Search staff for breakdown", placeholder="Type a name…")
    if staff_search:
        q = staff_search.strip().lower()
        shown_staff = [s for s in all_staff if q in s.lower()]
    else:
        shown_staff = all_staff

    if not shown_staff:
        st.info("No staff match the search. Clear the search to see all.")

    selected_list = st.multiselect("Staff (multi-select)", options=shown_staff, default=[])

    if len(selected_list) == 1:
        selected = selected_list[0]
        try:
            daily = per_staff_per_day(details_df, selected)
        except Exception as e:
            daily = pd.DataFrame()
            st.error(f"Failed to build per-day breakdown: {e}")
        if payroll_map is not None:
            rate = payroll_map.get(_norm_name(selected))
            if rate is not None:
                daily["Rate"] = rate
                daily["Total Pay"] = (daily.get("Total_Hours", 0) * rate).round(2)
            else:
                daily["Rate"] = pd.NA
                daily["Total Pay"] = pd.NA
            st.dataframe(_style_pay_df(daily), use_container_width=True, hide_index=True)
        else:
            st.dataframe(daily, use_container_width=True, hide_index=True)

        exp1, exp2 = st.columns(2)
        with exp1:
            st.download_button(
                label="Download per-day CSV",
                data=_bytes_csv(daily),
                file_name=f"{selected}_by_day.csv",
                mime="text/csv",
                disabled=daily.empty,
                key=f"dl_single_perday_{selected}",
            )
        with exp2:
            staff_rows = details_df[details_df["Staff Name"] == selected]
            st.download_button(
                label="Download staff details CSV",
                data=_bytes_csv(staff_rows),
                file_name=f"{selected}_details.csv",
                mime="text/csv",
                disabled=staff_rows.empty,
                key=f"dl_single_details_{selected}",
            )
    elif len(selected_list) > 1:
        st.caption(f"Selected staff: {len(selected_list)}")
        # Expand/collapse controls for multi-select breakdown
        if "multi_expand_all" not in st.session_state:
            st.session_state.multi_expand_all = False
        mc1, mc2 = st.columns([1,1])
        with mc1:
            if st.button("Expand all", key="btn_expand_multi"):
                st.session_state.multi_expand_all = True
        with mc2:
            if st.button("Collapse all", key="btn_collapse_multi"):
                st.session_state.multi_expand_all = False
        for staff in selected_list:
            # Build expander title with totals and optional total pay
            _mins = None
            _hours = None
            _units = None
            if filtered_df is not None and not filtered_df.empty:
                try:
                    _row = filtered_df[filtered_df["Staff Name"] == staff].iloc[0]
                    _mins = int(_row.get("Total_Minutes", 0) or 0)
                except Exception:
                    _mins = None
            if _mins is None:
                try:
                    _mins = int(details_df.loc[details_df["Staff Name"] == staff, "Rounded Minutes"].sum())
                except Exception:
                    _mins = 0
            _hours = _mins / 60.0
            _units = int(round(_mins / 15.0))
            # Build suffix with Total Pay and Completed pay buckets
            _suffix = ""
            try:
                _rate = payroll_map.get(_norm_name(staff)) if payroll_map is not None else None
            except Exception:
                _rate = None
            if _rate is not None:
                try:
                    _daily_title = per_staff_per_day(details_df, staff)
                except Exception:
                    _daily_title = pd.DataFrame(columns=["Total_Hours", "Completed"])
                _yes_hours = pd.to_numeric(_daily_title.loc[_daily_title.get("Completed", pd.Series()).eq("Yes"), "Total_Hours"], errors='coerce').fillna(0).sum()
                _no_hours  = pd.to_numeric(_daily_title.loc[_daily_title.get("Completed", pd.Series()).eq("No"),  "Total_Hours"], errors='coerce').fillna(0).sum()
                _tot_hours = pd.to_numeric(_daily_title.get("Total_Hours", pd.Series()), errors='coerce').fillna(0).sum()
                _total_pay = _escape_dollar(_fmt_usd(round(float(_tot_hours) * float(_rate), 2)))
                _yes_pay   = _escape_dollar(_fmt_usd(round(float(_yes_hours) * float(_rate), 2)))
                _no_pay    = _escape_dollar(_fmt_usd(round(float(_no_hours) * float(_rate), 2)))
                _suffix = f" | Total Pay:\u00A0{_total_pay} | Completed Pay:\u00A0{_yes_pay} | Completed No Pay:\u00A0{_no_pay}"
                exp_title = f"{staff} | {_mins} min | {_hours:.2f} h | {_units} units{_suffix}"
                exp_title = _escape_dollar(exp_title)
            elif payroll_map is not None:
                _suffix = " | missing pay rate"
                exp_title = f"{staff} | {_mins} min | {_hours:.2f} h | {_units} units{_suffix}"
            with st.expander(exp_title, expanded=bool(st.session_state.get("multi_expand_all", False))):
                try:
                    daily = per_staff_per_day(details_df, staff)
                except Exception as e:
                    daily = pd.DataFrame()
                    st.error(f"Failed to build per-day breakdown for {staff}: {e}")
                if payroll_map is not None:
                    rate = payroll_map.get(_norm_name(staff))
                    if rate is not None:
                        daily["Rate"] = rate
                        daily["Total Pay"] = (daily.get("Total_Hours", 0) * rate).round(2)
                    else:
                        daily["Rate"] = pd.NA
                        daily["Total Pay"] = pd.NA
                    st.dataframe(_style_pay_df(daily), use_container_width=True, hide_index=True)
                else:
                    st.dataframe(daily, use_container_width=True, hide_index=True)

                exp1, exp2 = st.columns(2)
                with exp1:
                    st.download_button(
                        label=f"Download {staff} per-day CSV",
                        data=_bytes_csv(daily),
                        file_name=f"{staff}_by_day.csv",
                        mime="text/csv",
                        disabled=daily.empty,
                        key=f"dl_multi_perday_{staff}",
                    )
                with exp2:
                    staff_rows = details_df[details_df["Staff Name"] == staff]
                    st.download_button(
                        label=f"Download {staff} details CSV",
                        data=_bytes_csv(staff_rows),
                        file_name=f"{staff}_details.csv",
                        mime="text/csv",
                        disabled=staff_rows.empty,
                        key=f"dl_multi_details_{staff}",
                    )

    st.subheader("Export All")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.download_button(
            label="Download summary.csv",
            data=_bytes_csv(filtered_df if filtered_df is not None else pd.DataFrame()),
            file_name="summary.csv",
            mime="text/csv",
            disabled=filtered_df is None or filtered_df.empty,
            key="dl_all_summary",
        )
    with c2:
        st.download_button(
            label="Download details.csv",
            data=_bytes_csv(details_df if details_df is not None else pd.DataFrame()),
            file_name="details.csv",
            mime="text/csv",
            disabled=details_df is None or details_df.empty,
            key="dl_all_details",
        )
    with c3:
        st.download_button(
            label="Download results.xlsx",
            data=_bytes_excel({"Summary": filtered_df if filtered_df is not None else pd.DataFrame(),
                              "Details": details_df if details_df is not None else pd.DataFrame()}),
            file_name="results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            disabled=(filtered_df is None or filtered_df.empty) and (details_df is None or details_df.empty),
            key="dl_all_results_xlsx",
        )


if __name__ == "__main__":
    main()
