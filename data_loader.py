"""
data_loader.py
==============
Reads and cleans the Excel files uploaded by the user (Phase 4).

TWO DATA SOURCES handled here:
  1. Backend Data   — one row per batching vehicle with Created by, Quantity, Date.
  2. Maintenance Cost — one row per plant with Plant Code and YTD Maintenance Cost.

WHAT "CLEANING" MEANS:
  Raw Excel files from the field often have:
    - Extra spaces in column names or cell values
    - Quantity stored as text ("1,200" instead of 1200)
    - Dates in various formats ("12-Jun-25", "12/06/2025", etc.)
    - Duplicate rows or blank rows
  This file fixes all of that before anything else touches the data.

PUBLIC FUNCTIONS:
  load_backend_data(uploaded_file)          -> (DataFrame, warnings_list)
  save_backend_data(df, filename, replace)  -> rows_inserted
  load_maintenance_cost(uploaded_file)      -> (DataFrame, warnings_list)
  save_maintenance_cost(df)                 -> rows_inserted
"""

from datetime import datetime

import pandas as pd

import config
import database


def _now() -> str:
    """Return current time as a short, sortable ISO string."""
    return datetime.now().isoformat(timespec="seconds")


# Date formats tried in order — most specific first so ambiguous values
# (like "01/06/25") are read as DD/MM/YY (Indian style), not MM/DD/YY.
_DATE_FORMATS = [
    # ── datetime (with time component) ─────────────────────────────────────
    "%Y-%m-%d %H:%M:%S",   # Excel export:   2026-05-01 00:00:00
    "%Y-%m-%dT%H:%M:%S",   # ISO-8601:       2026-05-01T00:00:00
    "%d/%m/%Y %H:%M:%S",   # Indian dt:      01/05/2026 17:48:25
    "%d-%m-%Y %H:%M:%S",   # Indian dt:      01-05-2026 17:48:25
    # ── date only ───────────────────────────────────────────────────────────
    "%Y-%m-%d",            # ISO:            2025-06-01
    "%d/%m/%Y",            # Indian long:    01/06/2025
    "%d-%m-%Y",            # Indian dashes:  01-06-2025
    "%d/%m/%y",            # Indian short:   01/06/25
    "%d-%m-%y",            # Indian short:   01-06-25
    "%d %b %Y",            # Text month:     01 Jun 2025
    "%d %B %Y",            # Full month:     01 June 2025
    "%b %d, %Y",           # US text:        Jun 01, 2025
]


def _parse_date(val) -> "pd.Timestamp":
    """
    Parse a single date/datetime value. Tries explicit formats first so that
    Indian DD/MM/YYYY is never confused with US MM/DD/YYYY.
    Falls back to pandas auto-detection as a last resort.
    Returns pd.NaT if nothing works — the caller will skip that row.
    """
    if pd.isna(val) or str(val).strip() in ("", "nan"):
        return pd.NaT
    s = str(val).strip()
    # Remove trailing time-zone info (e.g. "+05:30") if present
    s = s.split("+")[0].strip()
    for fmt in _DATE_FORMATS:
        try:
            return pd.to_datetime(s, format=fmt)
        except (ValueError, TypeError):
            pass
    # Final safety net — let pandas figure out any remaining formats
    try:
        return pd.to_datetime(s, dayfirst=True)
    except Exception:
        return pd.NaT


def _match_columns(df: pd.DataFrame, required_cols: list) -> tuple[dict, list]:
    """
    Two-pass column matching — handles different names and column orders.

    Pass 1 — exact match (case-insensitive):
        "Created by"  matches  "Created By"        ✓
        "Quantity"    matches  "quantity"           ✓

    Pass 2 — substring match (if exact fails):
        "Date"        matches  "Production date"   ✓  (required word inside sheet column)
        "Created by"  matches  "Created By (Emp)"  ✓
        "Quantity"    matches  "Batch Quantity"     ✓

    Returns:
        rename_map  – {actual_column_name: standard_name}  for df.rename()
        missing     – required columns that could not be matched at all
    """
    # Build lookup: lowercase → original name
    col_lower = {str(c).strip().lower(): str(c).strip() for c in df.columns}
    already_mapped = set()   # track sheet columns already claimed
    rename_map = {}
    missing = []

    for req in required_cols:
        req_l = req.lower()

        # ── Pass 1: exact case-insensitive match ────────────────────────────
        if req_l in col_lower and col_lower[req_l] not in already_mapped:
            actual = col_lower[req_l]
            rename_map[actual] = req
            already_mapped.add(actual)
            continue

        # ── Pass 2: required word(s) appear inside a sheet column name ──────
        # e.g.  req="Date"  matches sheet col "Production date"
        #        req="Created by"  matches "Created By Code"
        candidates = [
            orig for lower, orig in col_lower.items()
            if req_l in lower and orig not in already_mapped
        ]
        if candidates:
            actual = candidates[0]   # take the first (closest) match
            rename_map[actual] = req
            already_mapped.add(actual)
            continue

        missing.append(req)

    return rename_map, missing


# ---------------------------------------------------------------------------
# 1. BACKEND DATA
# ---------------------------------------------------------------------------

def load_backend_data(uploaded_file) -> tuple:
    """
    Read an uploaded Backend Data Excel (.xlsx) file and return a clean DataFrame.

    The file must contain (at minimum):
        Created by   – employee code of the person who batched
        Quantity     – how much was produced (numeric)
        Date         – the date of production

    Any extra columns in the file are ignored — we keep only the three above.

    Returns:
        df        – cleaned pandas DataFrame with columns: Created by, Quantity, Date
        warnings  – list of warning strings (e.g. "5 rows skipped: blank quantity")

    Raises:
        ValueError  – if the file can't be read or required columns are missing.
    """
    # Read the file. dtype=str keeps codes and dates exactly as typed.
    try:
        df = pd.read_excel(uploaded_file, dtype=str)
    except Exception as exc:
        raise ValueError(f"Could not read the Excel file: {exc}")

    if df.empty:
        raise ValueError("The uploaded file is empty.")

    # Clean up column names.
    df.columns = [str(c).strip() for c in df.columns]

    # Match required columns (case-insensitive).
    rename_map, missing = _match_columns(df, config.BACKEND_REQUIRED_COLUMNS)
    if missing:
        raise ValueError(
            f"These required columns are missing: {missing}\n"
            f"Columns found in the file: {list(df.columns)}\n\n"
            f"The file must have these column headers: {config.BACKEND_REQUIRED_COLUMNS}"
        )

    df = df.rename(columns=rename_map)[config.BACKEND_REQUIRED_COLUMNS].copy()
    warnings = []

    # --- Clean "Created by" (employee code) ---
    # fillna("") first because empty Excel cells arrive as float NaN even with dtype=str
    df["Created by"] = df["Created by"].fillna("").astype(str).str.strip()
    blank_codes = (df["Created by"].str.len() == 0) | (df["Created by"] == "nan")
    if blank_codes.sum() > 0:
        warnings.append(f"{blank_codes.sum()} rows skipped: blank 'Created by'.")
    df = df[~blank_codes].reset_index(drop=True)

    # --- Clean "Quantity" (must be a positive number) ---
    # Remove commas ("1,200" → "1200") then convert to float.
    df["Quantity"] = df["Quantity"].astype(str).str.replace(",", "", regex=False)
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    bad_qty = df["Quantity"].isna() | (df["Quantity"] <= 0)
    if bad_qty.sum() > 0:
        warnings.append(f"{bad_qty.sum()} rows skipped: missing or zero Quantity.")
    df = df[~bad_qty].reset_index(drop=True)

    # --- Clean "Date" (parse many formats, store as YYYY-MM-DD) ---
    # We try common formats in order so that Indian DD/MM/YYYY and ISO YYYY-MM-DD
    # both work correctly. pd.to_datetime with dayfirst=True alone fails on ISO
    # dates in pandas 2.x, so we use explicit formats first.
    df["Date"] = df["Date"].apply(_parse_date)
    bad_date = df["Date"].isna()
    if bad_date.sum() > 0:
        warnings.append(f"{bad_date.sum()} rows skipped: unrecognisable Date.")
    df = df[~bad_date].reset_index(drop=True)
    df["Date"] = df["Date"].dt.strftime("%Y-%m-%d")

    return df, warnings


def save_backend_data(df: pd.DataFrame, source_file: str,
                      replace: bool = True) -> int:
    """
    Save a cleaned Backend Data DataFrame to the SQLite backend_data table.

    Args:
        df          – cleaned DataFrame from load_backend_data()
        source_file – original filename (stored for traceability)
        replace     – if True, clears existing data first (default);
                      if False, appends to existing rows

    Returns the number of rows saved.
    """
    now = _now()
    rows = [
        {
            "created_by":  str(row["Created by"]),
            "quantity":    float(row["Quantity"]),
            "date":        str(row["Date"]),
            "source_file": source_file,
            "uploaded_at": now,
        }
        for _, row in df.iterrows()
    ]

    if replace:
        return database.replace_table_rows("backend_data", rows)
    return database.insert_rows("backend_data", rows)


# ---------------------------------------------------------------------------
# 2. MAINTENANCE COST DATA
# ---------------------------------------------------------------------------

def load_maintenance_cost(uploaded_file) -> tuple:
    """
    Read an uploaded Maintenance Cost Excel (.xlsx) file and return a clean DataFrame.

    The file must contain:
        Plant Code           – identifies the plant
        YTD Maintenance Cost – year-to-date cost per cubic metre (numeric)

    Returns:
        df        – cleaned DataFrame with columns: Plant Code, YTD Maintenance Cost
        warnings  – list of warning strings

    Raises:
        ValueError  – if the file can't be read or required columns are missing.
    """
    try:
        df = pd.read_excel(uploaded_file, dtype=str)
    except Exception as exc:
        raise ValueError(f"Could not read the Excel file: {exc}")

    if df.empty:
        raise ValueError("The uploaded file is empty.")

    df.columns = [str(c).strip() for c in df.columns]

    rename_map, missing = _match_columns(df, config.MAINTENANCE_REQUIRED_COLUMNS)
    if missing:
        raise ValueError(
            f"These required columns are missing: {missing}\n"
            f"Columns found in the file: {list(df.columns)}\n\n"
            f"The file must have these column headers: {config.MAINTENANCE_REQUIRED_COLUMNS}"
        )

    df = df.rename(columns=rename_map)[config.MAINTENANCE_REQUIRED_COLUMNS].copy()
    warnings = []

    # --- Clean "Plant Code" ---
    df["Plant Code"] = df["Plant Code"].fillna("").astype(str).str.strip()
    blank_plant = (df["Plant Code"].str.len() == 0) | (df["Plant Code"] == "nan")
    if blank_plant.sum() > 0:
        warnings.append(f"{blank_plant.sum()} rows skipped: blank Plant Code.")
    df = df[~blank_plant].reset_index(drop=True)

    # --- Clean "YTD Maintenance Cost" (must be a non-negative number) ---
    df["YTD Maintenance Cost"] = (
        df["YTD Maintenance Cost"].astype(str).str.replace(",", "", regex=False)
    )
    df["YTD Maintenance Cost"] = pd.to_numeric(
        df["YTD Maintenance Cost"], errors="coerce"
    )
    bad_cost = df["YTD Maintenance Cost"].isna()
    if bad_cost.sum() > 0:
        warnings.append(
            f"{bad_cost.sum()} rows skipped: non-numeric YTD Maintenance Cost."
        )
    df = df[~bad_cost].reset_index(drop=True)

    # --- Remove duplicate Plant Codes (keep last entry) ---
    dupes = df["Plant Code"].duplicated(keep="last").sum()
    if dupes > 0:
        warnings.append(
            f"{dupes} duplicate Plant Code(s) found — keeping the last entry each."
        )
        df = df.drop_duplicates(
            subset="Plant Code", keep="last"
        ).reset_index(drop=True)

    return df, warnings


def save_maintenance_cost(df: pd.DataFrame, month: int, year: int) -> int:
    """
    Save Maintenance Cost data for a specific month+year.
    Replaces any existing rows for that month+year, leaves other months untouched.
    Returns the number of rows saved.
    """
    now = _now()
    conn = database.get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            "DELETE FROM maintenance_cost WHERE month = ? AND year = ?",
            (month, year)
        )
        rows = [
            {
                "plant_code":           str(row["Plant Code"]),
                "month":                month,
                "year":                 year,
                "ytd_maintenance_cost": float(row["YTD Maintenance Cost"]),
                "uploaded_at":          now,
            }
            for _, row in df.iterrows()
        ]
        if rows:
            columns = list(rows[0].keys())
            placeholders = ", ".join(["?"] * len(columns))
            sql = (f"INSERT INTO maintenance_cost ({', '.join(columns)}) "
                   f"VALUES ({placeholders})")
            cur.executemany(sql, [list(r.values()) for r in rows])
        conn.commit()
        return len(rows)
    finally:
        conn.close()
