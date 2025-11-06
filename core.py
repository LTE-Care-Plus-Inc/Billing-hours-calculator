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
    def norm(x): return "" if pd.isna(x) else str(x).strip().lower()
    # Keep service filter, but do NOT filter by Completed anymore
    svc_ok  = df["Service Name"].apply(lambda x: norm(x) == "direct service bt")
    out = df[svc_ok].copy()
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

    g = df.groupby("_date_key", as_index=False).agg(
        Total_Minutes=("Rounded Minutes", "sum"),
        StartMin=(START_COL, "min"),
        EndMax=(END_COL, "max"),
        Completed=("Completed", _summarize_completed),
    )
    g[DATE_COL] = g["_date_key"]
    g["Appt Start"] = g["StartMin"].apply(_fmt_time)
    g["Appt End"]   = g["EndMax"].apply(_fmt_time)
    g.drop(columns=["_date_key", "StartMin", "EndMax"], inplace=True)
    g["Total_Hours"] = (g["Total_Minutes"] / 60.0).round(2)
    g["Total_Units"] = (g["Total_Minutes"] / 15.0).round(0).astype("Int64")
    try:
        g["_sort"] = pd.to_datetime(g[DATE_COL], errors="coerce")
        g = g.sort_values(by=["_sort", DATE_COL]).drop(columns=["_sort"])
    except Exception:
        pass
    return g[[DATE_COL, "Appt Start", "Appt End", "Total_Minutes", "Total_Hours", "Total_Units", "Completed"]]

def process_file(path: Path):
    df = load_input(path)
    df = normalize_colnames(df)
    df_f = filter_rows(df)

    eff_minutes, sources, rounded, discrepancies, invalid_idx = [], [], [], [], []
    for idx, row in df_f.iterrows():
        m, src = effective_minutes(row)
        if m is None or (isinstance(m, (int, float)) and (m < 0 or m > 24*60)):
            invalid_idx.append(idx)
            eff_minutes.append(None); sources.append(src or ""); rounded.append(None); discrepancies.append(False)
            continue
        r = round_15_with_8_rule(m)
        eff_minutes.append(m); sources.append(src or ""); rounded.append(r); discrepancies.append(discrepancy_flag(row))

    details = df_f.copy()
    details["Effective Minutes"] = eff_minutes
    details["Rounded Minutes"]   = rounded
    details["Source"]            = sources
    details["Discrepancy"]       = discrepancies
    if invalid_idx:
        details = details.drop(index=invalid_idx)

    summary = aggregate(details).sort_values(["Staff Name"], ascending=[True])
    return summary, details
