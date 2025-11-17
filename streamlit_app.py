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


@st.cache_data(show_spinner=False)
def _process_billable_bytes(content: bytes, suffix: str):
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(content)
        tmp_path = Path(tmp.name)
    try:
        return process_file(tmp_path)
    finally:
        try:
            os.unlink(tmp_path)
        except Exception:
            pass


def _process_uploaded_file(uploaded_file) -> tuple[pd.DataFrame, pd.DataFrame]:
    suffix = Path(uploaded_file.name).suffix.lower()
    content = uploaded_file.getbuffer().tobytes()
    summary, details = _process_billable_bytes(content, suffix)
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


def _style_pay_prep_df(df: pd.DataFrame):
    def red_missing(v):
        try:
            if v is None or pd.isna(v):
                return 'color: red'
        except Exception:
            pass
        if isinstance(v, str) and v.strip().lower() == 'missing pay rate':
            return 'color: red'
        return ''
    cols = {}
    for c in ["Rate", "Amount", "Hours x Rate"]:
        if c in df.columns:
            cols[c] = _fmt_usd
    sty = df.style.format(cols)
    try:
        sty = sty.applymap(red_missing, subset=["Rate", "Amount", "Hours x Rate"])  # pandas <2.1
    except Exception:
        sty = sty.map(red_missing, subset=["Rate", "Amount", "Hours x Rate"])      # pandas >=2.1
    return sty


@st.cache_data(show_spinner=False)
def _per_day_cached(details_df: pd.DataFrame, staff_name: str) -> pd.DataFrame:
    return per_staff_per_day(details_df, staff_name)


# Try to detect a client column from details
_CLIENT_CANDIDATES = [
    "Client Name",
    "Client",
    "Patient Name",
    "Member Name",
    "Student Name",
]


def _detect_client_col(df: pd.DataFrame) -> str | None:
    for c in _CLIENT_CANDIDATES:
        if c in df.columns:
            return c
    # Fallback: any column containing 'client' (case-insensitive)
    for c in df.columns:
        try:
            if "client" in str(c).lower():
                return c
        except Exception:
            continue
    return None


def _sessions_for_staff(details: pd.DataFrame, staff_name: str) -> pd.DataFrame:
    if details is None or details.empty:
        return pd.DataFrame()
    try:
        df = details[details.get("Staff Name", pd.Series(dtype=str)).astype(str) == str(staff_name)].copy()
    except Exception:
        return pd.DataFrame()
    if df.empty:
        return pd.DataFrame()

    # Ensure required columns exist
    for col in [DATE_COL, START_COL, END_COL]:
        if col not in df.columns:
            df[col] = pd.NaT

    client_col = _detect_client_col(df)
    if client_col is None:
        df["Client Name"] = ""
        client_col = "Client Name"

    # Sort by date then start time
    try:
        _sort_date = pd.to_datetime(df[DATE_COL], errors="coerce")
    except Exception:
        _sort_date = pd.to_datetime(pd.Series([], dtype=str))
    try:
        _sort_start = pd.to_datetime(df[START_COL], errors="coerce")
    except Exception:
        _sort_start = pd.to_datetime(pd.Series([], dtype=str))
    df = df.assign(__sort_date=_sort_date, __sort_start=_sort_start).sort_values(["__sort_date", "__sort_start", DATE_COL, START_COL])

    # Build display columns
    disp = pd.DataFrame({
        DATE_COL: df[DATE_COL],
        "Appt Start": df[START_COL].apply(_fmt_time),
        "Appt End": df[END_COL].apply(_fmt_time),
        "Client Name": df[client_col].astype(str),
    })
    # Optional columns
    if "Service Name" in df.columns:
        disp["Service Name"] = df["Service Name"].astype(str)
    if "Rounded Minutes" in df.columns:
        disp["Rounded Minutes"] = pd.to_numeric(df["Rounded Minutes"], errors="coerce").fillna(0).astype(int)
        disp["Rounded Hours"] = (disp["Rounded Minutes"] / 60.0).round(2)
    if "Completed" in df.columns:
        disp["Completed"] = df["Completed"]
    if "Completed Date" in df.columns:
        disp["Completed Date"] = df["Completed Date"]

    try:
        disp = disp.reset_index(drop=True)
    except Exception:
        pass
    return disp


@st.cache_data(show_spinner=False)
def _sessions_cached(details_df: pd.DataFrame, staff_name: str) -> pd.DataFrame:
    return _sessions_for_staff(details_df, staff_name)

# Build a payroll-prep summary per staff and earning code
def _build_payroll_prep_df(
    details_df: pd.DataFrame,
    payroll_rate_map: dict | None,
    payroll_hours_hourly: dict | None,
    payroll_hours_ot: dict | None,
    payroll_hours_cast: dict | None,
    payroll_period_label: str | None = None,
):
    # Final column label to match example: "Payroll Report for Payroll Period {period}"
    _last_col = (
        f"Payroll Report for Payroll Period {str(payroll_period_label).strip()}"
        if payroll_period_label and str(payroll_period_label).strip()
        else "Payroll Report"
    )
    if details_df is None or details_df.empty:
        return pd.DataFrame(columns=[
            "Staff Name","Earning Code","Rounded Hours","Miles","Rate","Amount","Hours x Rate","Notes","IF function for original amount compared to hours x rate",_last_col
        ])

    # Hours from details (billing): sum of rounded minutes → hours
    try:
        det_hours = (
            pd.to_numeric(details_df.get("Rounded Minutes", pd.Series(dtype=float)), errors='coerce')
              .fillna(0).groupby(details_df["Staff Name"].astype(str)).sum() / 60.0
        )
    except Exception:
        det_hours = pd.Series(dtype=float)

    rows = []
    staff_names = sorted(set(details_df["Staff Name"].astype(str)))

    def add_row(staff, code, hours, used_from):
        h = None if hours is None else float(hours)
        if h is not None:
            h = _round_quarter_hours(h)
        norm = _norm_name(staff)
        rate = None
        if payroll_rate_map is not None:
            try:
                rate = payroll_rate_map.get(norm)
                rate = float(rate) if rate is not None else None
            except Exception:
                rate = None
        amount = round(h * rate, 2) if (h is not None and rate is not None) else None
        hours_x_rate = amount

        notes = []
        if used_from:
            notes.append(f"used {used_from} hours")
        if rate is None:
            notes.append("missing pay rate")
        note_str = " | ".join(notes)

        mismatch = ""
        if code.lower().startswith("hourly") and 'overtime' not in code.lower() and 'cast' not in code.lower():
            # Compare payroll vs details hours when payroll hours exists
            d_h = float(det_hours.get(staff, 0.0) or 0.0)
            p_h = None
            if payroll_hours_hourly is not None:
                try:
                    p_h = payroll_hours_hourly.get(norm)
                except Exception:
                    p_h = None
            if p_h is not None:
                p_h = _round_quarter_hours(p_h)
                mismatch = "TRUE" if abs((d_h or 0.0) - (p_h or 0.0)) > 0.01 else "FALSE"
            else:
                mismatch = "FALSE"

        rows.append({
            "Staff Name": staff,
            "Earning Code": code,
            "Rounded Hours": h,
            "Miles": "-",
            "Rate": rate,
            "Amount": amount,
            "Hours x Rate": hours_x_rate,
            "Notes": note_str,
            "IF function for original amount compared to hours x rate": mismatch,
            _last_col: "",
        })

    for staff in staff_names:
        norm = _norm_name(staff)
        # Hourly
        p_hourly = payroll_hours_hourly.get(norm) if payroll_hours_hourly else None
        d_hourly = float(det_hours.get(staff, 0.0) or 0.0)
        if p_hourly is not None:
            add_row(staff, "Hourly", p_hourly, "payroll")
        else:
            add_row(staff, "Hourly", d_hourly, "billing")
        # Overtime
        p_ot = payroll_hours_ot.get(norm) if payroll_hours_ot else None
        if p_ot is not None and p_ot > 0:
            add_row(staff, "Hourly - Overtime", p_ot, "payroll")
        # Cast Back
        p_cb = payroll_hours_cast.get(norm) if payroll_hours_cast else None
        if p_cb is not None and p_cb > 0:
            add_row(staff, "Hourly - Cast Back", p_cb, "payroll")

    out = pd.DataFrame(rows)
    if not out.empty:
        try:
            out = out.sort_values(["Staff Name", "Earning Code"]).reset_index(drop=True)
        except Exception:
            pass
    # Ensure exact column order
    cols = [
        "Staff Name","Earning Code","Rounded Hours","Miles","Rate","Amount","Hours x Rate","Notes","IF function for original amount compared to hours x rate",_last_col
    ]
    for c in cols:
        if c not in out.columns:
            out[c] = ""
    out = out[cols]
    return out


def _format_payroll_prep_for_export(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    out = df.copy()
    # Currency formatting as strings for CSV to resemble Excel appearance
    for c in ["Rate", "Amount", "Hours x Rate"]:
        if c in out.columns:
            out[c] = out[c].apply(_fmt_usd)
    # Rounded Hours to 2 decimals
    if "Rounded Hours" in out.columns:
        out["Rounded Hours"] = pd.to_numeric(out["Rounded Hours"], errors='coerce').round(2)
    return out
# Removed per-day sessions helper per user request (keep only sessions-by-start-time in Inline view)

def _round_quarter_hours(x):
    try:
        v = float(x)
        return round(v * 4.0) / 4.0
    except Exception:
        return None


@st.cache_data(show_spinner=False)
def _process_payroll_bytes(content: bytes, suffix: str) -> pd.DataFrame:
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(content)
        tmp_path = Path(tmp.name)
    try:
        df = load_input(tmp_path)
    finally:
        try:
            os.unlink(tmp_path)
        except Exception:
            pass
    return df


def _process_payroll_file(uploaded_file) -> pd.DataFrame:
    suffix = Path(uploaded_file.name).suffix.lower()
    content = uploaded_file.getbuffer().tobytes()
    df = _process_payroll_bytes(content, suffix)
    # Normalize column names for robust access
    cols = {c.strip().lower(): c for c in df.columns}
    # Required columns
    staff_col = cols.get("staff name".lower()) or cols.get("staff name") or "Staff Name"
    rate_col = cols.get("rate".lower()) or cols.get("rate") or "Rate"
    job_col = cols.get("job title".lower()) or cols.get("job title") or "Job Title"
    earning_col = cols.get("earning code".lower()) or cols.get("earning code") or "Earning Code"
    hours_col = cols.get("hours".lower()) or cols.get("hours") or "Hours"
    if staff_col not in df.columns or rate_col not in df.columns:
        raise ValueError("Payroll report must include 'Staff Name' and 'Rate' columns")

    # Include all employees (no job title exclusions)
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
    # Compute hours per earning code (Hourly, Hourly - Overtime, Hourly - Cast Back)
    hours_hourly_map, hours_ot_map, hours_cast_map = {}, {}, {}
    try:
        if earning_col in df.columns and hours_col in df.columns:
            tmp = df.copy()
            ec = tmp[earning_col].astype(str).str.strip().str.lower()
            ec = ec.str.replace("–", "-", regex=False).str.replace("—", "-", regex=False)
            ec = ec.str.replace("  ", " ", regex=False)
            tmp["__ec__"] = ec
            tmp["__hours__"] = pd.to_numeric(tmp[hours_col], errors="coerce")
            grp = tmp.groupby(["__name_key__", "__ec__"])['__hours__'].sum(min_count=1)
            for (k, code), v in grp.items():
                if v is None:
                    continue
                q = _round_quarter_hours(v)
                code_l = str(code or '').lower()
                if 'overtime' in code_l:
                    hours_ot_map[k] = (hours_ot_map.get(k, 0.0) or 0.0) + (q or 0.0)
                elif 'cast' in code_l and 'back' in code_l:
                    hours_cast_map[k] = (hours_cast_map.get(k, 0.0) or 0.0) + (q or 0.0)
                elif 'hourly' in code_l:
                    hours_hourly_map[k] = (hours_hourly_map.get(k, 0.0) or 0.0) + (q or 0.0)
    except Exception:
        pass
    df_merged["__hours_hourly__"] = df_merged["__name_key__"].map(lambda k: hours_hourly_map.get(k))
    df_merged["__hours_ot__"] = df_merged["__name_key__"].map(lambda k: hours_ot_map.get(k))
    df_merged["__hours_castback__"] = df_merged["__name_key__"].map(lambda k: hours_cast_map.get(k))
    return df_merged[["__name_key__", "__rate__", "__hours_hourly__", "__hours_ot__", "__hours_castback__"]]


def _filter_and_sort(df: pd.DataFrame, query: str, sort_by_minutes: bool, details_df: pd.DataFrame | None = None) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    out = df.copy()
    q = (query or "").strip().lower()
    if q:
        mask_summary = out["Staff Name"].astype(str).str.lower().str.contains(q, na=False)
        mask_staff_from_details = pd.Series([False] * len(out), index=out.index)
        if details_df is not None and not details_df.empty and "Staff Name" in details_df.columns:
            try:
                dm = None
                for col in details_df.columns:
                    try:
                        s = details_df[col].astype(str).str.contains(q, case=False, na=False)
                    except Exception:
                        continue
                    dm = s if dm is None else (dm | s)
                staff_matched = set(details_df.loc[dm.fillna(False), "Staff Name"].astype(str)) if dm is not None else set()
                if staff_matched:
                    mask_staff_from_details = out["Staff Name"].astype(str).isin(staff_matched)
            except Exception:
                pass
        out = out[mask_summary | mask_staff_from_details]
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
            help=("CSV/Excel with 'Staff Name', 'Rate', optional 'Job Title',"
                  " 'Earning Code', and 'Hours'. Overtime totals read from"
                  " rows where Earning Code = 'Hourly- Overtime'."),
        )

        # Optional payroll period label to include in the final column header
        payroll_period_label = st.text_input(
            "Payroll Period",
            placeholder="e.g. 10/13/25-10/26/25",
            help=("Used to name the last column as \"Payroll Report for Payroll Period {value}\"."
                  " Leave blank to use \"Payroll Report\"."),
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
        _bg = "#0f1419"; _text = "#e6edf3"; _sbg = "#161b22"; _input = "#1b222c"; _border = "#2d333b"; _link = "#58a6ff"; _primary = "#2C7BE5"; _ph = "rgba(230,237,243,0.85)"
    else:
        _bg = "#ffffff"; _text = "#0b1220"; _sbg = "#f5f7fb"; _input = "#ffffff"; _border = "#d8dee4"; _link = "#1f6feb"; _primary = "#2C7BE5"; _ph = "#4b5563"
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
            color: {_ph} !important;
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
        /* Captions spacing + strong color (better readability) */
        p, .stCaption, [data-testid="stCaptionContainer"] {{ margin-top: 0.25rem; }}
        [data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] * , .stCaption {{
            color: {_text} !important;
            opacity: 1.0 !important;
            font-weight: 600;
        }}
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

    # In dark mode, force any white non-text backgrounds/borders to gray variants
    if _t == "dark":
        st.markdown(
            f"""
            <style>
            /* Force white backgrounds to gray in dark mode */
            [data-testid="stAppViewContainer"] .block-container,
            .stDataFrame [role="grid"],
            .stDataFrame [role="columnheader"],
            .stTextInput, .stSelectbox, .stMultiSelect, .stNumberInput,
            .stFileUploader, .stDateInput, .stTimeInput, .stTextArea,
            .stDownloadButton, .stRadio, .stCheckbox, .stColorPicker,
            [data-baseweb="input"], [data-baseweb="select"],
            .stAlert, .stMetric, .stTabs, [role="tablist"] button[role="tab"],
            .stProgress, .stSlider, .stSlider > div, .row-widget {{
                background-color: {_sbg} !important;
            }}

            /* Catch inline white backgrounds/borders and make them gray */
            [style*="background-color: #fff"],
            [style*="background-color: white"],
            [style*="background-color: rgb(255, 255, 255)"] {{
                background-color: {_sbg} !important;
            }}
            [style*="border-color: #fff"],
            [style*="border-color: white"],
            [style*="border-color: rgb(255, 255, 255)"] {{
                border-color: {_border} !important;
            }}

            /* Ensure AG Grid dark sections don't revert to white */
            .ag-theme-balham-dark, .ag-theme-balham-dark .ag-root-wrapper,
            .ag-theme-balham-dark .ag-header, .ag-theme-balham-dark .ag-row,
            .ag-theme-balham-dark .ag-cell {{
                background-color: {_bg} !important;
            }}
            .ag-theme-balham-dark .ag-row {{ color: {_text} !important; }}
            .ag-theme-balham-dark .ag-row-odd {{ background-color: #121821 !important; }}
            .ag-theme-balham-dark .ag-row-hover {{ background-color: #18212b !important; }}
            .ag-theme-balham-dark .ag-row-selected {{ background-color: #203044 !important; }}
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
        query = st.text_input("Search", placeholder="Type any term (e.g., BT, BCBA, service)")
    with col2:
        sort_choice = st.selectbox("Sort by", ["Staff Name", "Minutes (desc)"])
    with col3:
        if source_label:
            st.caption(source_label)

    sort_by_minutes = sort_choice.lower().startswith("minutes")
    filtered_df = _filter_and_sort(summary_df, query, sort_by_minutes, details_df)

    # Totals footer
    if filtered_df is not None and not filtered_df.empty:
        mins = int(filtered_df["Total_Minutes"].sum())
        hours = round(mins / 60.0, 2)
        units = int(round(mins / 15.0))
        st.caption(f"Totals (filtered): {mins} minutes | {hours:.2f} hours | {units} units")

    # Optional: load payroll mapping (after we have summary/details)
    payroll_map = None
    payroll_hours_hourly = None
    payroll_hours_ot = None
    payroll_hours_cast = None
    payroll_err = None
    if payroll_uploaded is not None:
        try:
            payroll_df = _process_payroll_file(payroll_uploaded)
            # Build mapping from normalized name to rate
            payroll_map = {r["__name_key__"]: r["__rate__"] for _, r in payroll_df.iterrows()}
            if "__hours_hourly__" in payroll_df.columns:
                payroll_hours_hourly = {r["__name_key__"]: r["__hours_hourly__"] for _, r in payroll_df.iterrows() if pd.notna(r.get("__hours_hourly__"))}
            if "__hours_ot__" in payroll_df.columns:
                payroll_hours_ot = {r["__name_key__"]: r["__hours_ot__"] for _, r in payroll_df.iterrows() if pd.notna(r.get("__hours_ot__"))}
            if "__hours_castback__" in payroll_df.columns:
                payroll_hours_cast = {r["__name_key__"]: r["__hours_castback__"] for _, r in payroll_df.iterrows() if pd.notna(r.get("__hours_castback__"))}
        except Exception as e:
            payroll_err = str(e)

    # Enrich details with Rate, Rounded Hours (if missing), and Total Amount
    try:
        if details_df is not None and not details_df.empty:
            if "Rounded Hours" not in details_df.columns and "Rounded Minutes" in details_df.columns:
                details_df["Rounded Hours"] = (pd.to_numeric(details_df["Rounded Minutes"], errors='coerce') / 60.0).round(2)
            if payroll_map is not None:
                name_keys = details_df.get("Staff Name", pd.Series(dtype=str)).astype(str).apply(_norm_name)
                rates = name_keys.map(lambda k: payroll_map.get(k))
                details_df["Rate"] = pd.to_numeric(rates, errors='coerce')
                if "Rounded Hours" in details_df.columns:
                    details_df["Total Amount"] = (pd.to_numeric(details_df["Rounded Hours"], errors='coerce') * details_df["Rate"]).round(2)
    except Exception:
        pass

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
                    # Display as: Last, First | Last2, First2 | ...
                    shown = missing[:10]
                    miss_list = " | ".join(shown) + (" | ..." if len(missing) > 10 else " |")
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
                        # Show overtime hours header when available
                        try:
                            if payroll_hours_ot is not None:
                                _ot = payroll_hours_ot.get(_norm_name(name))
                                if _ot is not None and _ot > 0:
                                    st.markdown(f"### Overtime Hours (payroll): {_ot:.2f}")
                        except Exception:
                            pass
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

                        # Removed sessions sections from Interactive view per request

        with tab_inline:
            records = filtered_df.to_dict("records")
            # Pagination controls for inline breakdown
            if "inline_page_size" not in st.session_state:
                st.session_state.inline_page_size = 20
            if "inline_page" not in st.session_state:
                st.session_state.inline_page = 1

            topc1, topc2, topc3, topc4 = st.columns([1,1,2,2])
            with topc1:
                page_size = st.selectbox("Rows/page", [10, 20, 50, 100], index=[10,20,50,100].index(st.session_state.inline_page_size), key="inline_page_size")
            total = len(records)
            pages = max(1, (total + page_size - 1) // page_size)
            # Clamp page within bounds
            st.session_state.inline_page = min(max(1, st.session_state.inline_page), pages)
            with topc2:
                # Prev/Next
                prev, next_ = st.columns(2)
                with prev:
                    if st.button("< Prev", disabled=(st.session_state.inline_page <= 1), key="inline_prev"):
                        st.session_state.inline_page = max(1, st.session_state.inline_page - 1)
                with next_:
                    if st.button("Next >", disabled=(st.session_state.inline_page >= pages), key="inline_next"):
                        st.session_state.inline_page = min(pages, st.session_state.inline_page + 1)
            with topc3:
                st.caption(f"Showing page {st.session_state.inline_page} of {pages}")
            with topc4:
                st.caption(f"Total staff: {total}")

            # Expand/collapse controls
            if "inline_expand_all" not in st.session_state:
                st.session_state.inline_expand_all = False
            ic1, ic2 = st.columns([1,1])
            with ic1:
                if st.button("Expand all", key="btn_expand_inline"):
                    st.session_state.inline_expand_all = True
            with ic2:
                if st.button("Collapse all", key="btn_collapse_inline"):
                    st.session_state.inline_expand_all = False

            start = (st.session_state.inline_page - 1) * page_size
            end = min(total, start + page_size)
            page_records = records[start:end]

            for rec in page_records:
                name = str(rec.get("Staff Name", ""))
                mins = int(rec.get("Total_Minutes", 0) or 0)
                hours = mins / 60.0
                units = int(round(mins / 15.0))
                title = f"{name} | {mins} min | {hours:.2f} h | {units} units"
                # Build title with Total Pay and Completed pay buckets
                __rate = payroll_map.get(_norm_name(name)) if payroll_map is not None else None
                if __rate is not None:
                    # Lightweight title: compute total pay without per-day grouping
                    total_pay = _escape_dollar(_fmt_usd(round(float(hours) * float(__rate), 2)))
                    __suffix = f" | Total Pay:\u00A0{total_pay}"
                elif payroll_map is not None:
                    __suffix = " | missing pay rate"
                else:
                    __suffix = ""
                title = f"{name} | {mins} min | {hours:.2f} h | {units} units{__suffix}"
                title = _escape_dollar(title)
                with st.expander(title, expanded=bool(st.session_state.get("inline_expand_all", False))):
                    # Show overtime hours header when available
                    try:
                        if payroll_hours_ot is not None:
                            _ot = payroll_hours_ot.get(_norm_name(name))
                            if _ot is not None and _ot > 0:
                                st.markdown(f"### Overtime Hours (payroll): {_ot:.2f}")
                    except Exception:
                        pass
                    # Lazy render: only compute when toggled
                    show_daily = st.checkbox("Show per-day breakdown", value=False, key=f"show_daily_inline_{name}")
                    if show_daily:
                        try:
                            daily = _per_day_cached(details_df, name)
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
                        if show_daily:
                            st.download_button(
                                label=f"Download {name} per-day CSV",
                                data=_bytes_csv(daily),
                                file_name=f"{name}_by_day.csv",
                                mime="text/csv",
                                disabled=daily.empty,
                                key=f"dl_perday_inline_{name}",
                            )
                        else:
                            st.caption("Enable per-day to download")
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

                    # Sessions breakdown for this staff (sorted by start time) — keep only in Inline view
                    # Sessions on-demand to save time
                    show_sessions = st.checkbox("Show sessions (by start time)", value=False, key=f"show_sessions_inline_{name}")
                    if show_sessions:
                        try:
                            sessions = _sessions_cached(details_df, name)
                        except Exception as e:
                            sessions = pd.DataFrame()
                            st.error(f"Failed to build sessions: {e}")
                        st.dataframe(sessions, use_container_width=True, hide_index=True)
                        st.download_button(
                            label=f"Download {name} sessions CSV",
                            data=_bytes_csv(sessions),
                            file_name=f"{name}_sessions.csv",
                            mime="text/csv",
                            disabled=sessions.empty,
                            key=f"dl_sessions_inline_{name}",
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
            daily = _per_day_cached(details_df, selected)
        except Exception as e:
            daily = pd.DataFrame()
            st.error(f"Failed to build per-day breakdown: {e}")
        # Show overtime header for single selected staff
        try:
            if payroll_hours_ot is not None:
                _ot = payroll_hours_ot.get(_norm_name(selected))
                if _ot is not None and _ot > 0:
                    st.markdown(f"### Overtime Hours (payroll): {_ot:.2f}")
        except Exception:
            pass

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

        # Removed sessions sections from single-staff view per request
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
        # Removed sessions download in single-staff view
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
                _total_pay = _escape_dollar(_fmt_usd(round(float(_hours) * float(_rate), 2)))
                _suffix = f" | Total Pay:\u00A0{_total_pay}"
                exp_title = f"{staff} | {_mins} min | {_hours:.2f} h | {_units} units{_suffix}"
                exp_title = _escape_dollar(exp_title)
            elif payroll_map is not None:
                _suffix = " | missing pay rate"
                exp_title = f"{staff} | {_mins} min | {_hours:.2f} h | {_units} units{_suffix}"
            with st.expander(exp_title, expanded=bool(st.session_state.get("multi_expand_all", False))):
                # Show overtime hours header when available
                try:
                    if payroll_hours_ot is not None:
                        _ot = payroll_hours_ot.get(_norm_name(staff))
                        if _ot is not None and _ot > 0:
                            st.markdown(f"### Overtime Hours (payroll): {_ot:.2f}")
                except Exception:
                    pass
                try:
                    daily = _per_day_cached(details_df, staff)
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
                # Removed sessions sections from multi-staff view per request

    # Payroll Prep (aggregated)
    st.subheader("Payroll Prep")
    try:
        prep_df = _build_payroll_prep_df(
            details_df,
            payroll_map,
            payroll_hours_hourly,
            payroll_hours_ot,
            payroll_hours_cast,
            payroll_period_label,
        )
    except Exception as e:
        prep_df = pd.DataFrame()
        st.error(f"Failed to build Payroll Prep: {e}")
    if not prep_df.empty:
        try:
            st.dataframe(_style_pay_prep_df(prep_df), use_container_width=True, hide_index=True)
        except Exception:
            st.dataframe(prep_df, use_container_width=True, hide_index=True)
    else:
        st.dataframe(prep_df, use_container_width=True, hide_index=True)
    st.download_button(
        label="Download payroll_prep.csv",
        data=_bytes_csv(_format_payroll_prep_for_export(prep_df)),
        file_name="payroll_prep.csv",
        mime="text/csv",
        disabled=prep_df.empty,
        key="dl_payroll_prep",
    )

    st.subheader("Export All")
    c1, c2, c3 = st.columns(3)
    with c1:
        # Removed summary download per requirements
        pass
    with c2:
        # Export details as payroll-prep formatted CSV (as requested)
        _export_df = _format_payroll_prep_for_export(prep_df)
        st.download_button(
            label="Download details.csv",
            data=_bytes_csv(_export_df),
            file_name="details.csv",
            mime="text/csv",
            disabled=_export_df is None or _export_df.empty,
            key="dl_all_details",
        )
    with c3:
        # Removed results.xlsx download per requirements
        pass


if __name__ == "__main__":
    main()

