"""
google_sheets.py
================
Connects to Google Sheets and syncs the Master Data into SQLite.

TWO ACCESS MODES (chosen automatically):

  1. PUBLIC MODE (no credentials needed)
     If your Google Sheet is shared as "Anyone with the link can view",
     we fetch it as a CSV export — no API key required.
     This is the default mode.

  2. PRIVATE MODE (service account credentials)
     If credentials/service_account.json exists, we use the Google Sheets
     API via a service account. This lets you sync from private sheets.
     See credentials/README.txt for setup instructions.

The app picks mode 2 automatically when the credentials file is present,
and falls back to mode 1 otherwise. You never need to change code to switch.

ACCEPTS FULL URLS OR SHEET IDs:
  You can paste either of these into the Sheet ID field:
    - Full URL:  https://docs.google.com/spreadsheets/d/1BxiMVs0.../edit
    - Just the ID:  1BxiMVs0...
  Both work — the app extracts the ID from the URL automatically.
"""

import io
import os
import re
import urllib.parse
from datetime import datetime

import pandas as pd

import config
import database

# Path to the optional service account key file.
CREDS_PATH = os.path.join("credentials", "service_account.json")

# Google API scopes (only needed for private/service-account mode).
# Full spreadsheets scope is required for write-back (add/edit/delete rows).
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]


# ---------------------------------------------------------------------------
# 1. HELPERS
# ---------------------------------------------------------------------------

def credentials_exist() -> bool:
    """Return True if the service account JSON file is present."""
    return os.path.exists(CREDS_PATH)


def extract_sheet_id(url_or_id: str) -> str:
    """
    Accept any Google Sheets URL or raw ID and return just the sheet ID.

    Handles two URL formats:
      Published URL:  .../spreadsheets/d/e/{PUB_ID}/pubhtml  → returns {PUB_ID}
      Regular URL:    .../spreadsheets/d/{SHEET_ID}/edit     → returns {SHEET_ID}
      Raw ID:         1BxiMVs0...                            → returned as-is
    """
    url_or_id = url_or_id.strip()
    # Published-to-web URL: /d/e/{ID}/...  (must check before the regular pattern)
    match = re.search(r"/spreadsheets/d/e/([a-zA-Z0-9_-]+)", url_or_id)
    if match:
        return match.group(1)
    # Regular URL: /d/{ID}/...
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url_or_id)
    if match:
        return match.group(1)
    return url_or_id  # already a raw ID


def _is_published_id(sheet_id: str) -> bool:
    """
    Published-to-web IDs always start with '2PACX-'.
    Regular sheet IDs are shorter alphanumeric strings.
    """
    return sheet_id.startswith("2PACX-")


# ---------------------------------------------------------------------------
# 2. FETCH — PUBLIC MODE (no credentials)
# ---------------------------------------------------------------------------

def _fetch_public(sheet_id: str, worksheet_name: str) -> pd.DataFrame:
    """
    Fetch a publicly accessible Google Sheet and return a raw DataFrame.

    Tries three URL formats in order (most reliable first):
      1. export?format=csv  — works when sheet is "Anyone with link can view"
      2. gviz/tq CSV        — works when sheet is "Published to the web"
      3. pub?output=csv     — works when sheet is "Published to the web"

    Uses the requests library for reliable fetching and detects HTML error
    pages that Google returns when permissions are wrong.
    """
    try:
        import requests as _requests
    except ImportError:
        raise ImportError(
            "The 'requests' package is not installed.\n"
            "Run: python -m pip install requests"
        )

    encoded_name = urllib.parse.quote(worksheet_name)
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

    # Choose URL set based on the sheet ID type.
    if _is_published_id(sheet_id):
        # "Published to the web" sheets — use the /d/e/ path with pub?output=csv
        urls = [
            f"https://docs.google.com/spreadsheets/d/e/{sheet_id}/pub?output=csv&sheet={encoded_name}",
            f"https://docs.google.com/spreadsheets/d/e/{sheet_id}/pub?output=csv",
        ]
    else:
        # Regular sheets shared as "Anyone with the link can view".
        # IMPORTANT: the gviz endpoint is the only one that honours the
        # `sheet=<name>` parameter — the plain `export?format=csv&sheet=`
        # URL ignores the name and always returns the FIRST/default tab.
        # So gviz must be tried first whenever we need a specific tab.
        urls = [
            f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={encoded_name}",
            f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&sheet={encoded_name}",
            f"https://docs.google.com/spreadsheets/d/{sheet_id}/pub?output=csv&sheet={encoded_name}",
        ]

    last_error = "Unknown error"
    for url in urls:
        try:
            resp = _requests.get(url, timeout=15, headers=headers, allow_redirects=True)

            content = resp.text.strip()

            # Google returns an HTML page when permissions deny the request.
            if resp.status_code != 200 or content.startswith("<"):
                last_error = (
                    f"HTTP {resp.status_code} — Google returned an error page, "
                    "not CSV data. Check sheet permissions."
                )
                continue  # try next URL format

            # Parse as CSV — dtype=str keeps numeric codes like "818000000001"
            # from being converted to scientific notation (8.18E+11).
            df = pd.read_csv(io.StringIO(content), dtype=str)
            if df.empty:
                last_error = "The sheet returned empty data."
                continue

            return df  # success

        except Exception as exc:
            last_error = str(exc)
            continue

    # All three formats failed.
    raise ValueError(
        "Could not read your Google Sheet after trying all URL formats.\n\n"
        f"Last error: {last_error}\n\n"
        "How to fix:\n"
        "  1. Open your Google Sheet.\n"
        "  2. Click File → Share → Publish to the web.\n"
        "  3. Choose 'Entire Document' and 'CSV', then click Publish.\n"
        "  4. Come back and click Sync Now again.\n\n"
        "OR confirm the sheet is shared as 'Anyone with the link can view'."
    )


# ---------------------------------------------------------------------------
# 3. FETCH — PRIVATE MODE (service account credentials)
# ---------------------------------------------------------------------------

def _fetch_private(sheet_id: str, worksheet_name: str) -> pd.DataFrame:
    """
    Fetch a Google Sheet using the gspread library and a service account.
    Used automatically when credentials/service_account.json exists.
    """
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError:
        raise ImportError(
            "The 'gspread' and 'google-auth' packages are not installed.\n"
            "Run:  python -m pip install gspread google-auth"
        )

    creds = Credentials.from_service_account_file(CREDS_PATH, scopes=SCOPES)
    client = gspread.authorize(creds)

    spreadsheet = client.open_by_key(sheet_id)
    worksheet = spreadsheet.worksheet(worksheet_name)
    records = worksheet.get_all_records(expected_headers=None)

    if not records:
        return pd.DataFrame()

    df = pd.DataFrame(records)
    # Keep all columns as strings so numeric codes aren't turned into floats.
    return df.astype(str)


# ---------------------------------------------------------------------------
# 4. LAYOUT DETECTION & TRANSPOSE
# ---------------------------------------------------------------------------

def _is_transposed(df: pd.DataFrame) -> bool:
    """
    Detect whether the sheet uses a TRANSPOSED (row-based) layout.

    Normal layout  — field names are COLUMN HEADERS (first row):
        Employee Code | Employee Name | Designation | ...
        E001          | John          | SPE         | ...
        E002          | Jane          | MO          | ...

    Transposed layout — field names are ROW LABELS (first column):
        Employee Code | E001  | E002  | ...
        Employee Name | John  | Jane  | ...
        Designation   | SPE   | MO    | ...

    We detect it by counting how many of the required field names appear
    as values inside the first column vs. as column headers.
    """
    if df.empty:
        return False

    required_lower = {c.lower() for c in config.MASTER_DATA_COLUMNS}

    # How many required names appear as values in the first column?
    first_col_values = {str(v).strip().lower() for v in df.iloc[:, 0]}
    matches_in_rows = len(required_lower & first_col_values)

    # How many required names appear as column headers?
    header_values = {str(c).strip().lower() for c in df.columns}
    matches_in_headers = len(required_lower & header_values)

    # If the first column holds more field names than the headers do, it's transposed.
    return matches_in_rows > matches_in_headers


def _transpose_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    Pivot a transposed sheet back to the standard column-based layout.

    Before (transposed):
        col[0]="Employee Code" | col[1]="E001" | col[2]="E002"
        row 0: "Employee Name" | "John"         | "Jane"
        row 1: "Designation"   | "SPE"          | "MO"
        ...

    After (standard):
        Employee Code | Employee Name | Designation | ...
        E001          | John          | SPE         | ...
        E002          | Jane          | MO          | ...
    """
    # The name of column 0 is the first field label (e.g. "Employee Code").
    first_field = str(df.columns[0]).strip()

    # Make the first column the DataFrame index so it becomes column names after .T
    df = df.set_index(df.columns[0])

    # Transpose: old column names (E001, E002 …) become the row index.
    df = df.T

    # Bring the row index (employee codes) back as a normal column.
    df.index.name = first_field
    df = df.reset_index()

    # Clean up any whitespace in the new column names.
    df.columns = [str(c).strip() for c in df.columns]

    return df


# ---------------------------------------------------------------------------
# 5. SHARED CLEANING & VALIDATION
# ---------------------------------------------------------------------------

def _clean_and_validate(raw_df: pd.DataFrame) -> pd.DataFrame:
    """
    Take a raw DataFrame (from either public or private fetch) and:
      - Auto-detect and fix the transposed layout if needed.
      - Strip whitespace from column names and cell values.
      - Map sheet column names → config.MASTER_DATA_COLUMNS (case-insensitive).
      - Raise ValueError if a required column is missing.
      - Drop blank Employee Code rows.
      - Return a clean DataFrame with exactly the 6 required columns.
    """
    if raw_df.empty:
        return pd.DataFrame(columns=config.MASTER_DATA_COLUMNS)

    # Strip whitespace from column headers first (needed for detection).
    raw_df.columns = [str(c).strip() for c in raw_df.columns]

    # Auto-detect transposed layout and flip it back to standard.
    if _is_transposed(raw_df):
        raw_df = _transpose_df(raw_df)

    # Case-insensitive column matching.
    sheet_cols_lower = {c.lower(): c for c in raw_df.columns}
    rename_map = {}
    missing_cols = []

    for required_col in config.MASTER_DATA_COLUMNS:
        if required_col.lower() in sheet_cols_lower:
            rename_map[sheet_cols_lower[required_col.lower()]] = required_col
        else:
            missing_cols.append(required_col)

    if missing_cols:
        raise ValueError(
            f"The sheet is missing these required columns: {missing_cols}\n"
            f"Columns found in the sheet: {list(raw_df.columns)}\n\n"
            "Your sheet must have these labels (spelling matters, case does not):\n"
            f"{config.MASTER_DATA_COLUMNS}"
        )

    df = raw_df.rename(columns=rename_map)[config.MASTER_DATA_COLUMNS].copy()

    # Strip whitespace from every cell.
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()

    # Drop rows where Employee Code is blank or the literal string "nan".
    df = df[
        (df["Employee Code"].str.len() > 0) & (df["Employee Code"] != "nan")
    ].reset_index(drop=True)

    # Remove duplicate employee codes — keep the last occurrence so that
    # any corrections made lower in the sheet win over earlier entries.
    dupes = df["Employee Code"].duplicated(keep="last").sum()
    if dupes > 0:
        df = df.drop_duplicates(subset="Employee Code", keep="last").reset_index(drop=True)

    return df


# ---------------------------------------------------------------------------
# 5. PUBLIC FETCH ENTRY POINT
# ---------------------------------------------------------------------------

def fetch_master_data(sheet_id: str, worksheet_name: str) -> pd.DataFrame:
    """
    Fetch and return master data as a clean DataFrame.

    Tries service account mode first (if credentials exist), then falls back
    to public URL mode so sheets not shared with the service account still work.
    """
    sheet_id = extract_sheet_id(sheet_id)
    raw_df, _ = _fetch_with_fallback(sheet_id, worksheet_name)
    return _clean_and_validate(raw_df)


# ---------------------------------------------------------------------------
# 6. SYNC: FETCH + SAVE TO SQLITE
# ---------------------------------------------------------------------------

def sync_master_data(sheet_id: str, worksheet_name: str) -> dict:
    """
    Fetch the latest Master Data from Google Sheets and replace the SQLite
    master_data table with it.

    Returns:
        {
            "rows_synced": int,        # rows saved to SQLite
            "synced_at":  str | None,  # ISO timestamp
            "error":      str | None,  # None = success
            "mode":       str,         # "public" or "private"
        }

    Never raises — all errors are captured so the UI can show a message.
    """
    mode = "private" if credentials_exist() else "public"
    try:
        clean_id = extract_sheet_id(sheet_id)
        df = fetch_master_data(clean_id, worksheet_name)

        now = datetime.now().isoformat(timespec="seconds")

        import math as _math
        def _sv(v):
            if v is None: return ""
            if isinstance(v, float) and _math.isnan(v): return ""
            s = str(v).strip()
            return "" if s.lower() == "nan" else s

        rows = [
            {
                "employee_code": _sv(row["Employee Code"]),
                "employee_name": _sv(row["Employee Name"]),
                "designation":   _sv(row["Designation"]),
                "category":      _sv(row["Category"]),
                "plant":         _sv(row["Plant"]),
                "plant_code":    _sv(row["Plant Code"]),
                "updated_at":    now,
            }
            for _, row in df.iterrows()
        ]

        inserted = database.replace_table_rows("master_data", rows)

        # Persist sync details so the UI can show "Last synced …" info.
        database.set_setting("gsheet_id",         clean_id)
        database.set_setting("gsheet_worksheet",  worksheet_name)
        database.set_setting("gsheet_last_sync",  now)
        database.set_setting("gsheet_last_count", str(inserted))

        return {"rows_synced": inserted, "synced_at": now, "error": None, "mode": mode}

    except Exception as exc:
        return {"rows_synced": 0, "synced_at": None, "error": str(exc), "mode": mode}


# ---------------------------------------------------------------------------
# 7. RDC-TP: Sync plant data from "Plant Data for TP" sheet tab
# ---------------------------------------------------------------------------
# Expected columns (case-insensitive):
#   Plant Code | Exco Location | Plant | Business Head | Plant Manager/ Asst. PI | Mixer Theo. Capacity

_TP_PLANT_COLS = [
    "Plant Code",
    "Exco Location",
    "Plant",
    "Business Head",
    "Plant Manager/ Asst. PI",
    "Mixer Theo. Capacity",
]


def _clean_tp_plant_data(raw_df: pd.DataFrame) -> pd.DataFrame:
    """
    Validate and clean the raw 'Plant Data for TP' sheet DataFrame.
    Returns a clean DataFrame with exactly the 6 required columns.
    """
    if raw_df.empty:
        return pd.DataFrame(columns=_TP_PLANT_COLS)

    raw_df.columns = [str(c).strip() for c in raw_df.columns]

    # Case-insensitive match
    sheet_lower = {c.lower(): c for c in raw_df.columns}
    rename_map, missing = {}, []
    for req in _TP_PLANT_COLS:
        if req.lower() in sheet_lower:
            rename_map[sheet_lower[req.lower()]] = req
        else:
            missing.append(req)

    if missing:
        raise ValueError(
            f"'Plant Data for TP' sheet is missing columns: {missing}\n"
            f"Columns found: {list(raw_df.columns)}"
        )

    df = raw_df.rename(columns=rename_map)[_TP_PLANT_COLS].copy()
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()

    # Drop rows where Plant Code is blank
    df = df[(df["Plant Code"].str.len() > 0) & (df["Plant Code"] != "nan")].reset_index(drop=True)

    # Remove duplicate plant codes — keep last
    df = df.drop_duplicates(subset="Plant Code", keep="last").reset_index(drop=True)

    return df


def _fetch_with_fallback(sheet_id: str, worksheet_name: str) -> tuple:
    """
    Try private (service account) fetch first; fall back to public URL fetch.
    Returns (DataFrame, mode_string).
    """
    if credentials_exist():
        try:
            return _fetch_private(sheet_id, worksheet_name), "private"
        except Exception:
            pass  # service account doesn't have access to this sheet → try public
    return _fetch_public(sheet_id, worksheet_name), "public"


def sync_tp_plant_data(sheet_id: str, worksheet_name: str = "Plant Data for TP") -> dict:
    """
    Fetch the Plant Data for TP sheet and save to tp_plant_data table.
    Returns {"rows_synced": int, "synced_at": str|None, "error": str|None, "mode": str}
    """
    mode = "public"
    try:
        clean_id = extract_sheet_id(sheet_id)
        raw_df, mode = _fetch_with_fallback(clean_id, worksheet_name)

        df = _clean_tp_plant_data(raw_df)
        now = datetime.now().isoformat(timespec="seconds")

        rows = [
            {
                "plant_code":     str(r["Plant Code"]),
                "exco_location":  str(r["Exco Location"]),
                "plant_name":     str(r["Plant"]),
                "business_head":  str(r["Business Head"]),
                "plant_manager":  str(r["Plant Manager/ Asst. PI"]),
                "mixer_theo_cap": _to_float(r["Mixer Theo. Capacity"]),
                "updated_at":     now,
            }
            for _, r in df.iterrows()
        ]

        inserted = database.replace_table_rows("tp_plant_data", rows)
        database.set_module_setting("tp", "gsheet_last_sync",  now)
        database.set_module_setting("tp", "gsheet_last_count", str(inserted))
        database.set_module_setting("tp", "gsheet_worksheet",  worksheet_name)

        return {"rows_synced": inserted, "synced_at": now, "error": None, "mode": mode}

    except Exception as exc:
        return {"rows_synced": 0, "synced_at": None, "error": str(exc), "mode": mode}


def sync_btrtp_master_data(sheet_id: str, worksheet_name: str = "BT Master Data") -> dict:
    """
    Fetch the BT Master Data sheet and save to btrtp_master_data table.
    Expected sheet columns: Batcher ID, Batcher Name
    Returns {"rows_synced": int, "synced_at": str|None, "error": str|None, "mode": str}
    """
    mode = "public"
    try:
        clean_id = extract_sheet_id(sheet_id)
        raw_df, mode = _fetch_with_fallback(clean_id, worksheet_name)

        if raw_df is None or raw_df.empty:
            raise ValueError(
                f"Sheet '{worksheet_name}' returned no data. "
                "Check the tab name and sharing settings."
            )

        raw_df.columns = [str(c).strip() for c in raw_df.columns]
        cols_lower = {c.lower(): c for c in raw_df.columns}

        # Match ID column — tries several common naming conventions
        id_col = (
            cols_lower.get("batcher id")
            or cols_lower.get("employee code")
            or cols_lower.get("emp code")
            or cols_lower.get("employee id")
            or cols_lower.get("emp id")
            or next((c for c in raw_df.columns if "batcher" in c.lower() and "id" in c.lower()), None)
            or next((c for c in raw_df.columns if "employee" in c.lower() and "code" in c.lower()), None)
            or next((c for c in raw_df.columns if "employee" in c.lower() and "id" in c.lower()), None)
        )
        # Match Name column — tries several common naming conventions
        name_col = (
            cols_lower.get("batcher name")
            or cols_lower.get("employee name")
            or cols_lower.get("emp name")
            or next((c for c in raw_df.columns if "batcher" in c.lower() and "name" in c.lower()), None)
            or next((c for c in raw_df.columns if "employee" in c.lower() and "name" in c.lower()), None)
        )

        if not id_col or not name_col:
            raise ValueError(
                f"Could not find ID and Name columns. "
                f"Expected 'Employee Code'+'Employee Name' or 'Batcher ID'+'Batcher Name'. "
                f"Found columns: {list(raw_df.columns)}"
            )

        def _clean_id(val) -> str:
            """Strip whitespace; remove trailing '.0' that pandas adds to numeric codes."""
            s = str(val).strip()
            if s.endswith(".0"):
                try:
                    s = str(int(float(s)))
                except Exception:
                    pass
            return s

        now = datetime.now().isoformat(timespec="seconds")

        # Build rows — deduplicate by batcher_id (keep last) to avoid UNIQUE constraint errors
        seen: dict = {}
        for _, r in raw_df.iterrows():
            bid = _clean_id(r[id_col])
            if bid in ("", "nan"):
                continue
            seen[bid] = {
                "batcher_id":   bid,
                "batcher_name": str(r[name_col]).strip(),
                "updated_at":   now,
            }
        rows = list(seen.values())

        database.replace_table_rows("btrtp_master_data", rows)
        database.set_module_setting("btrtp", "gsheet_last_sync",  now)
        database.set_module_setting("btrtp", "gsheet_last_count", str(len(rows)))
        database.set_module_setting("btrtp", "gsheet_worksheet",  worksheet_name)

        return {"rows_synced": len(rows), "synced_at": now, "error": None, "mode": mode}

    except Exception as exc:
        return {"rows_synced": 0, "synced_at": None, "error": str(exc), "mode": mode}


def get_btrtp_last_sync_info() -> dict:
    """Return last BTRTP master sync info."""
    return {
        "last_sync":  database.get_module_setting("btrtp", "gsheet_last_sync",  ""),
        "last_count": database.get_module_setting("btrtp", "gsheet_last_count", "0"),
        "worksheet":  database.get_module_setting("btrtp", "gsheet_worksheet",  "BT Master"),
    }


def _to_float(val) -> float:
    """Convert a cell value to float, returning 0.0 on failure."""
    try:
        return float(str(val).replace(",", "").strip())
    except (ValueError, TypeError):
        return 0.0


def get_tp_last_sync_info() -> dict:
    """Return last TP plant data sync info."""
    return {
        "last_sync":  database.get_module_setting("tp", "gsheet_last_sync",  None),
        "last_count": database.get_module_setting("tp", "gsheet_last_count", "0"),
        "worksheet":  database.get_module_setting("tp", "gsheet_worksheet",  "Plant Data for TP"),
    }


# ---------------------------------------------------------------------------
# 8. LAST SYNC INFO (RDC-I&D master data)
# ---------------------------------------------------------------------------

def get_last_sync_info() -> dict:
    """Return the last sync details stored in app_settings."""
    return {
        "sheet_id":   database.get_setting("gsheet_id",        ""),
        "worksheet":  database.get_setting("gsheet_worksheet", ""),
        "last_sync":  database.get_setting("gsheet_last_sync", None),
        "last_count": database.get_setting("gsheet_last_count", "0"),
    }


# ---------------------------------------------------------------------------
# 9. RDC-TP: Write-back — push add / update / delete to Google Sheet
# ---------------------------------------------------------------------------
# Sheet column order (must match _TP_PLANT_COLS):
#   A: Plant Code | B: Exco Location | C: Plant | D: Business Head
#   E: Plant Manager/ Asst. PI | F: Mixer Theo. Capacity

_TP_SHEET_HEADERS = [
    "Plant Code", "Exco Location", "Plant", "Business Head",
    "Plant Manager/ Asst. PI", "Mixer Theo. Capacity",
]


def _gspread_worksheet(sheet_id: str, worksheet_name: str):
    """Open a gspread Worksheet object (write-capable). Raises if no credentials."""
    if not credentials_exist():
        raise RuntimeError(
            "No service account credentials found. "
            "Add credentials/service_account.json to enable write-back to Google Sheets."
        )
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError:
        raise ImportError("Run: python -m pip install gspread google-auth")

    creds = Credentials.from_service_account_file(CREDS_PATH, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(sheet_id).worksheet(worksheet_name)


def _find_plant_row(ws, plant_code: str) -> int:
    """
    Find the 1-based row index of a plant in the sheet by matching Plant Code
    in column A. Returns 0 if not found.
    """
    col_a = ws.col_values(1)   # all values in column A (1-indexed)
    for i, val in enumerate(col_a):
        if str(val).strip().upper() == str(plant_code).strip().upper():
            return i + 1       # gspread rows are 1-based
    return 0


def push_tp_plant_add(plant_data: dict) -> dict:
    """
    Append a new plant row to the Google Sheet.
    plant_data keys: plant_code, exco_location, plant_name,
                     business_head, plant_manager, mixer_theo_cap
    Returns {"ok": bool, "message": str}
    """
    try:
        sheet_id  = database.get_module_setting("tp", "gsheet_id",
                    database.get_setting("gsheet_id", ""))
        ws_name   = database.get_module_setting("tp", "gsheet_worksheet", "Plant Data for TP")
        if not sheet_id:
            return {"ok": False, "message": "Google Sheet ID not configured."}

        ws = _gspread_worksheet(extract_sheet_id(sheet_id), ws_name)
        row = [
            str(plant_data.get("plant_code", "")),
            str(plant_data.get("exco_location", "")),
            str(plant_data.get("plant_name", "")),
            str(plant_data.get("business_head", "")),
            str(plant_data.get("plant_manager", "")),
            str(plant_data.get("mixer_theo_cap", "")),
        ]
        ws.append_row(row, value_input_option="USER_ENTERED")
        return {"ok": True, "message": "Row added to Google Sheet."}
    except RuntimeError as exc:
        return {"ok": False, "message": str(exc)}
    except Exception as exc:
        return {"ok": False, "message": f"Sheet write failed: {exc}"}


def push_tp_plant_update(plant_data: dict) -> dict:
    """
    Update an existing plant row in the Google Sheet (matched by Plant Code).
    Returns {"ok": bool, "message": str}
    """
    try:
        sheet_id  = database.get_module_setting("tp", "gsheet_id",
                    database.get_setting("gsheet_id", ""))
        ws_name   = database.get_module_setting("tp", "gsheet_worksheet", "Plant Data for TP")
        if not sheet_id:
            return {"ok": False, "message": "Google Sheet ID not configured."}

        ws  = _gspread_worksheet(extract_sheet_id(sheet_id), ws_name)
        row_idx = _find_plant_row(ws, plant_data["plant_code"])
        if row_idx == 0:
            # Row not found in sheet — append as new row instead
            return push_tp_plant_add(plant_data)

        ws.update(f"A{row_idx}:F{row_idx}", [[
            str(plant_data.get("plant_code", "")),
            str(plant_data.get("exco_location", "")),
            str(plant_data.get("plant_name", "")),
            str(plant_data.get("business_head", "")),
            str(plant_data.get("plant_manager", "")),
            str(plant_data.get("mixer_theo_cap", "")),
        ]], value_input_option="USER_ENTERED")
        return {"ok": True, "message": "Row updated in Google Sheet."}
    except RuntimeError as exc:
        return {"ok": False, "message": str(exc)}
    except Exception as exc:
        return {"ok": False, "message": f"Sheet write failed: {exc}"}


def push_tp_plant_delete(plant_code: str) -> dict:
    """
    Delete a plant row from the Google Sheet (matched by Plant Code).
    Returns {"ok": bool, "message": str}
    """
    try:
        sheet_id  = database.get_module_setting("tp", "gsheet_id",
                    database.get_setting("gsheet_id", ""))
        ws_name   = database.get_module_setting("tp", "gsheet_worksheet", "Plant Data for TP")
        if not sheet_id:
            return {"ok": False, "message": "Google Sheet ID not configured."}

        ws = _gspread_worksheet(extract_sheet_id(sheet_id), ws_name)
        row_idx = _find_plant_row(ws, plant_code)
        if row_idx == 0:
            return {"ok": False, "message": f"Plant '{plant_code}' not found in sheet."}

        ws.delete_rows(row_idx)
        return {"ok": True, "message": f"Plant '{plant_code}' deleted from Google Sheet."}
    except RuntimeError as exc:
        return {"ok": False, "message": str(exc)}
    except Exception as exc:
        return {"ok": False, "message": f"Sheet write failed: {exc}"}


# ---------------------------------------------------------------------------
# 10. RDC-I&D: Write-back — push add / update / delete to Google Sheet
# ---------------------------------------------------------------------------
# Sheet column order (must match MASTER_DATA_COLUMNS in config.py):
#   A: Employee Code | B: Employee Name | C: Designation
#   D: Category | E: Plant | F: Plant Code

_ID_SHEET_HEADERS = [
    "Employee Code", "Employee Name", "Designation",
    "Category", "Plant", "Plant Code",
]


def _find_employee_row(ws, employee_code: str) -> int:
    """Find the 1-based row index of an employee matched by Employee Code in column A."""
    col_a = ws.col_values(1)
    for i, val in enumerate(col_a):
        if str(val).strip().upper() == str(employee_code).strip().upper():
            return i + 1
    return 0


def _get_id_worksheet() -> object:
    """
    Open the I&D master data worksheet (write-capable).
    The I&D module may store a published-to-web (2PACX-) ID which is read-only.
    For write-back we use the TP regular sheet ID instead (same workbook).
    """
    ws_name  = database.get_setting("gsheet_worksheet", "BT Master Data")
    # Prefer the TP regular sheet ID (writable); fall back to I&D setting
    sheet_id = database.get_module_setting("tp", "gsheet_id",
               database.get_setting("gsheet_id", ""))
    if not sheet_id or _is_published_id(sheet_id):
        # TP id not set either — nothing we can do
        raise ValueError(
            "Could not find a writable Google Sheet ID. "
            "Please sync Plant Data for TP first to register the sheet ID."
        )
    return _gspread_worksheet(extract_sheet_id(sheet_id), ws_name)


def push_id_employee_add(emp: dict) -> dict:
    """
    Append a new employee row to the I&D master data Google Sheet.
    emp keys: employee_code, employee_name, designation, category, plant, plant_code
    """
    try:
        ws = _get_id_worksheet()
        row = [
            str(emp.get("employee_code", "")),
            str(emp.get("employee_name", "")),
            str(emp.get("designation", "")),
            str(emp.get("category", "")),
            str(emp.get("plant", "")),
            str(emp.get("plant_code", "")),
        ]
        ws.append_row(row, value_input_option="USER_ENTERED")
        return {"ok": True, "message": "Row added to Google Sheet."}
    except RuntimeError as exc:
        return {"ok": False, "message": str(exc)}
    except Exception as exc:
        return {"ok": False, "message": f"Sheet write failed: {exc}"}


def push_id_employee_update(emp: dict) -> dict:
    """
    Update an existing employee row in the I&D master data Google Sheet.
    emp keys: employee_code, employee_name, designation, category, plant, plant_code
    """
    try:
        ws      = _get_id_worksheet()
        row_idx = _find_employee_row(ws, emp["employee_code"])
        if row_idx == 0:
            return push_id_employee_add(emp)
        ws.update(f"A{row_idx}:F{row_idx}", [[
            str(emp.get("employee_code", "")),
            str(emp.get("employee_name", "")),
            str(emp.get("designation", "")),
            str(emp.get("category", "")),
            str(emp.get("plant", "")),
            str(emp.get("plant_code", "")),
        ]], value_input_option="USER_ENTERED")
        return {"ok": True, "message": "Row updated in Google Sheet."}
    except RuntimeError as exc:
        return {"ok": False, "message": str(exc)}
    except Exception as exc:
        return {"ok": False, "message": f"Sheet write failed: {exc}"}


def push_id_employee_delete(employee_code: str) -> dict:
    """Delete an employee row from the I&D master data Google Sheet."""
    try:
        ws      = _get_id_worksheet()
        row_idx = _find_employee_row(ws, employee_code)
        if row_idx == 0:
            return {"ok": False, "message": f"Employee '{employee_code}' not found in sheet."}
        ws.delete_rows(row_idx)
        return {"ok": True, "message": f"Employee '{employee_code}' deleted from Google Sheet."}
    except RuntimeError as exc:
        return {"ok": False, "message": str(exc)}
    except Exception as exc:
        return {"ok": False, "message": f"Sheet write failed: {exc}"}
