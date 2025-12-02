"""
core.py — Billable Hours Aggregator v5.1
- Flexible header aliases for Appt. Date / Start / End
- Per-day aggregation includes earliest Start and latest End
"""
import pandas as pd
from pathlib import Path

# Required columns for filtering/aggregation
REQUIRED = {
    "service_name": "Service Name",
    "completed": "Completed",
    "staff_name": "Staff Name",
    "billing_minutes": "Billing Minutes",
    "billing_hours": "Billing Hours",
    "units": "Units",
}

# Canonical labels we will use in the DataFrame
DATE_COL  = "Appt. Date"
START_COL = "Appt. Start"
END_COL   = "Appt. End"

# Common header variants → canonical
ALIASES = {
    # Date
    "appt. date": DATE_COL, "appt date": DATE_COL,
    "appointment date": DATE_COL, "date": DATE_COL,

    # Start
    "appt. start": START_COL, "appt start": START_COL,
    "appt. start time": START_COL, "appt start time": START_COL,
    "appointment start": START_COL, "appointment start time": START_COL,
    "start": START_COL, "start time": START_COL,

    # End
    "appt. end": END_COL, "appt end": END_COL,
    "appt. end time": END_COL, "appt end time": END_COL,
    "appointment end": END_COL, "appointment end time": END_COL,
    "end": END_COL, "end time": END_COL,
}

def _lower(s): return s.strip().lower() if isinstance(s, str) else s

def normalize_colnames(df: pd.DataFrame) -> pd.DataFrame:
    # Lowercase everything
    df2 = df.rename(columns=lambda c: c.strip().lower())

    # First: map required columns back to canonical labels
    rev = {}
    for _, human in REQUIRED.items():
        key = human.lower()
        if key in df2.columns:
            rev[key] = human
    # Second: map known aliases for date/start/end
    for col in list(df2.columns):
        if col in ALIASES:
            rev[col] = ALIASES[col]
    # Explicitly normalize Completed Date if present
    if "completed date" in df2.columns:
        rev["completed date"] = "Completed Date"
    if rev:
        df2 = df2.rename(columns=rev)

    # Ensure required columns exist
    missing = [v for v in REQUIRED.values() if v not in df2.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}. Found: {list(df2.columns)}")

    return df2

def coerce_number(x):
    if pd.isna(x): return None
    if isinstance(x, (int, float)):
        return float(x)
    try:
        s = str(x).strip().replace(",", "")
        return float(s) if s else None
    except Exception:
        return None

def effective_minutes(row):
    m = coerce_number(row.get("Billing Minutes"))
    if m is not None: return m, "Minutes"
    h = coerce_number(row.get("Billing Hours"))
    if h is not None: return h * 60.0, "Hours"
    u = coerce_number(row.get("Units"))
    if u is not None: return u * 15.0, "Units"
    return None, None

def round_15_with_8_rule(m):
    if m is None or pd.isna(m): return None
    m = float(m)
    r = m % 15
    return m - r + 15 if r >= 7 else m - r

def discrepancy_flag(row):
    bm = coerce_number(row.get("Billing Minutes"))
    if bm is None: return False
    h = coerce_number(row.get("Billing Hours"))
    u = coerce_number(row.get("Units"))
    alt = h * 60.0 if h is not None else (u * 15.0 if u is not None else None)
    return False if alt is None else abs(bm - alt) > 7

def session_duration_minutes(row):
    """Compute duration from Appt. Start to Appt. End in minutes, if both are valid."""
    try:
        date_val = row.get(DATE_COL)
        start_raw = row.get(START_COL)
        end_raw = row.get(END_COL)
        if pd.isna(start_raw) or pd.isna(end_raw):
            return None

        # Try parsing with date first for better accuracy, else fall back to time-only
        if pd.isna(date_val):
            start = pd.to_datetime(start_raw, errors="coerce")
            end = pd.to_datetime(end_raw, errors="coerce")
        else:
            start = pd.to_datetime(f"{date_val} {start_raw}", errors="coerce")
            end = pd.to_datetime(f"{date_val} {end_raw}", errors="coerce")

        if pd.isna(start) or pd.isna(end):
            # Final fallback: time-only parsing
            start = pd.to_datetime(start_raw, errors="coerce")
            end = pd.to_datetime(end_raw, errors="coerce")

        if pd.isna(start) or pd.isna(end):
            return None

        delta = end - start
        mins = delta.total_seconds() / 60.0
        return mins if mins >= 0 else None
    except Exception:
        return None

def load_input(path: Path) -> pd.DataFrame:
    suf = path.suffix.lower()
    if suf in [".xlsx", ".xlsm", ".xltx", ".xltm"]:
        return pd.read_excel(path)
    if suf == ".csv":
        return pd.read_csv(path)
    raise ValueError("Unsupported file type. Provide .csv or .xlsx")

def parse_date_time_cols(df: pd.DataFrame) -> pd.DataFrame:
    # Date
    if DATE_COL in df.columns:
        try: df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce").dt.date
        except Exception: pass
    else:
        df[DATE_COL] = pd.NaT

    # Start/End times → time objects when possible
    for col in [START_COL, END_COL]:
        if col in df.columns:
            try:
                parsed = pd.to_datetime(df[col], errors="coerce")
                df[col] = parsed.dt.time
            except Exception:
                # leave as-is (string or NaN)
                pass
        else:
            df[col] = pd.NaT
    return df

def filter_rows(df: pd.DataFrame) -> pd.DataFrame:
    # Only include sessions that are completed for billing
    out = df.copy()
    try:
        comp_series = out.get("Completed")
        if comp_series is None:
            # If no Completed column (should not happen due to REQUIRED), keep empty
            out = out.iloc[0:0]
        else:
            # Normalize completion values
            comp_norm = comp_series.apply(lambda v: "" if pd.isna(v) else str(v).strip().lower())
            complete_markers = {
                "yes", "y", "true", "complete", "completed", "done", "signed", "finished", "finalized"
            }
            mask_yes = comp_norm.isin(complete_markers)

            # Treat presence of a Completed Date as completed
            if "Completed Date" in out.columns:
                mask_date = out["Completed Date"].apply(
                    lambda v: False if (pd.isna(v) or str(v).strip() == "") else True
                )
            else:
                mask_date = False

            mask = mask_yes | mask_date
            out = out[mask].copy()
    except Exception:
        # On any unexpected error, fall back to original rows
        out = df.copy()
    return parse_date_time_cols(out)

def aggregate(df_details: pd.DataFrame) -> pd.DataFrame:
    df_details["Staff Name"] = df_details["Staff Name"].apply(lambda s: "" if pd.isna(s) else " ".join(str(s).split()))
    grp = df_details.groupby("Staff Name", dropna=False, as_index=False).agg(
        Total_Minutes=("Rounded Minutes", "sum")
    )
    grp["Total_Hours"] = (grp["Total_Minutes"] / 60.0).round(2)
    grp["Total_Units"] = (grp["Total_Minutes"] / 15.0).round(0).astype("Int64")
    return grp

def _fmt_time(t):
    if pd.isna(t) or t is None: return ""
    try: return t.strftime("%H:%M")
    except Exception:
        s = str(t).strip()
        try: return pd.to_datetime(s).strftime("%H:%M")
        except Exception: return s

def per_staff_per_day(details: pd.DataFrame, staff_name: str) -> pd.DataFrame:
    df = details[details["Staff Name"] == staff_name].copy()
    if df.empty:
        return pd.DataFrame(columns=[DATE_COL, "Appt Start", "Appt End", "Total_Minutes", "Total_Hours", "Total_Units", "Completed"])

    # Ensure columns exist
    for col in [DATE_COL, START_COL, END_COL]:
        if col not in df.columns:
            df[col] = pd.NaT

    df["_date_key"] = df[DATE_COL].astype("string").fillna("(no date)")
    def _summarize_completed(series):
        # Normalize values, track originals for display
        normalized = []
        originals = []
        for x in series:
            if pd.isna(x):
                nx = ""
                ox = ""
            else:
                ox = str(x).strip()
                nx = ox.lower()
            normalized.append(nx)
            originals.append(ox)

        norm_set = set(normalized)
        # If any 'yes' is present, the day is Yes
        if "yes" in norm_set:
            return "Yes"

        # Gather distinct non-empty, non-yes labels
        other_pairs = [(n, o) for n, o in zip(normalized, originals) if n not in ("", "yes")]
        other_norms = list(dict.fromkeys([n for n, _ in other_pairs]))  # preserve order, unique by normalized

        if len(other_norms) == 0:
            # No other label present; treat as No by default
            return "No"
        if len(other_norms) == 1:
            # Display the first original occurrence of that label
            for n, o in other_pairs:
                if n == other_norms[0]:
                    return o or "No"
        # Multiple distinct labels found
        return "Mixed"

    # Summarize Completed Date values: single distinct value -> that value; multiple -> 'Mixed'
    def _summarize_completed_date(series):
        try:
            vals = []
            for x in series:
                if pd.isna(x) or x is None:
                    continue
                vals.append(str(x).strip())
            uniq = list(dict.fromkeys([v for v in vals if v]))
            if len(uniq) == 0:
                return ""
            if len(uniq) == 1:
                return uniq[0]
            return "Mixed"
        except Exception:
            return ""

    agg_dict = {
        "Total_Minutes": ("Rounded Minutes", "sum"),
        "StartMin": (START_COL, "min"),
        "EndMax": (END_COL, "max"),
        "Completed": ("Completed", _summarize_completed),
    }
    if "Completed Date" in df.columns:
        agg_dict["Completed_Date"] = ("Completed Date", _summarize_completed_date)
    g = df.groupby("_date_key", as_index=False).agg(**agg_dict)
    g[DATE_COL] = g["_date_key"]
    g["Appt Start"] = g["StartMin"].apply(_fmt_time)
    g["Appt End"]   = g["EndMax"].apply(_fmt_time)
    for c in ["_date_key", "StartMin", "EndMax"]:
        if c in g.columns:
            try:
                g.drop(columns=[c], inplace=True)
            except Exception:
                pass
    g["Total_Hours"] = (g["Total_Minutes"] / 60.0).round(2)
    g["Total_Units"] = (g["Total_Minutes"] / 15.0).round(0).astype("Int64")
    try:
        g["_sort"] = pd.to_datetime(g[DATE_COL], errors="coerce")
        g = g.sort_values(by=["_sort", DATE_COL]).drop(columns=["_sort"])
    except Exception:
        pass
    # Add Service Name(s) per day if available
    if "Service Name" in df.columns:
        svc_series = (
            df.groupby("_date_key")["Service Name"]
              .apply(lambda s: ", ".join(list(dict.fromkeys([str(x).strip() for x in s if not pd.isna(x) and str(x).strip()]))))
        )
        svc_df = svc_series.reset_index().rename(columns={"Service Name": "Service Name(s)"})
        # Merge back on date key
        svc_df = svc_df.rename(columns={"_date_key": DATE_COL})
        g = g.merge(svc_df, on=DATE_COL, how="left")

    # Rename Completed_Date to display label
    if "Completed_Date" in g.columns:
        g = g.rename(columns={"Completed_Date": "Completed Date"})

    # Order columns with Completed Date next to Completed
    cols = [DATE_COL, "Appt Start", "Appt End"]
    if "Service Name(s)" in g.columns:
        cols.append("Service Name(s)")
    cols += ["Total_Minutes", "Total_Hours", "Total_Units", "Completed"]
    if "Completed Date" in g.columns:
        cols.append("Completed Date")
    return g[[c for c in cols if c in g.columns]]

def process_file(path: Path):
    df = load_input(path)
    df = normalize_colnames(df)

    # Compute per-staff count of sessions excluded due to being incomplete
    incomplete_by_staff = {}
    try:
        comp_series = df.get("Completed")
        if comp_series is not None:
            comp_norm = comp_series.apply(lambda v: "" if pd.isna(v) else str(v).strip().lower())
            complete_markers = {"yes", "y", "true", "complete", "completed", "done", "signed", "finished", "finalized"}
            mask_yes = comp_norm.isin(complete_markers)
            if "Completed Date" in df.columns:
                mask_date = df["Completed Date"].apply(lambda v: False if (pd.isna(v) or str(v).strip() == "") else True)
            else:
                mask_date = False
            if isinstance(mask_date, bool):
                mask_completed = mask_yes
            else:
                mask_completed = mask_yes | mask_date
            mask_incomplete = ~mask_completed
            try:
                tmp = df.loc[mask_incomplete].copy()
                tmp["Staff Name"] = tmp["Staff Name"].astype(str)
                incomplete_by_staff = tmp.groupby("Staff Name", dropna=False).size().to_dict()
            except Exception:
                incomplete_by_staff = {}
        else:
            incomplete_by_staff = {}
    except Exception:
        incomplete_by_staff = {}

    df_f = filter_rows(df)

    eff_minutes, sources, rounded, discrepancies, invalid_idx = [], [], [], [], []
    for idx, row in df_f.iterrows():
        m, src = effective_minutes(row)
        duration_m = session_duration_minutes(row)
        if m is None or (isinstance(m, (int, float)) and (m < 0 or m > 24*60)):
            invalid_idx.append(idx)
            eff_minutes.append(None); sources.append(src or ""); rounded.append(None); discrepancies.append(False)
            continue
        r = round_15_with_8_rule(m)
        # Cap rounded minutes to the scheduled duration (floor to 15-min blocks) when start/end are present
        if duration_m is not None and r is not None:
            try:
                dur_cap = int(float(duration_m) // 15) * 15  # floor to nearest 15 to avoid slight overages
                r = min(r, dur_cap)
            except Exception:
                pass
        eff_minutes.append(m); sources.append(src or ""); rounded.append(r); discrepancies.append(discrepancy_flag(row))

    details = df_f.copy()
    details["Effective Minutes"] = eff_minutes
    details["Rounded Minutes"]   = rounded
    # Provide hours in 0.25 increments derived from Rounded Minutes
    try:
        details["Rounded Hours"] = (pd.to_numeric(details["Rounded Minutes"], errors="coerce") / 60.0).round(2)
    except Exception:
        details["Rounded Hours"] = pd.NA
    details["Source"]            = sources
    details["Discrepancy"]       = discrepancies
    if invalid_idx:
        details = details.drop(index=invalid_idx)

    # Attach per-staff incomplete counts to details rows
    try:
        details["Incomplete_Excluded_Count"] = details["Staff Name"].astype(str).map(lambda s: int(incomplete_by_staff.get(s, 0))).astype("Int64")
        details["Has_Incomplete_Excluded"] = details["Incomplete_Excluded_Count"].apply(lambda x: False if pd.isna(x) else bool(int(x) > 0))
    except Exception:
        pass

    summary = aggregate(details).sort_values(["Staff Name"], ascending=[True])
    # Attach per-staff incomplete counts to summary too (for UI convenience)
    try:
        summary["Incomplete_Excluded_Count"] = summary["Staff Name"].astype(str).map(lambda s: int(incomplete_by_staff.get(s, 0))).astype("Int64")
        summary["Has_Incomplete_Excluded"] = summary["Incomplete_Excluded_Count"].apply(lambda x: False if pd.isna(x) else bool(int(x) > 0))
    except Exception:
        pass
    return summary, details
