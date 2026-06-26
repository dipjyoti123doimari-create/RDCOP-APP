"""
mapping_service.py
==================
Reads the existing Google Sheet to provide:
  - Plant Code  → Plant Name
  - Plant Code  → Plant Manager Email
  - Plant Code  → Business Manager Email
  - Plant Code  → Business Head Email
  - Plant Code  → CC Emails
  - Batcher Code → Batcher Name

Expected Google Sheet columns (flexible — fuzzy matched):
  Plant Code / Plant_Code / PlantCode
  Plant Name / Plant_Name / PlantName
  Plant Manager Email / PlantManagerEmail / PM Email
  Business Manager Email / BusinessManagerEmail / BM Email
  Business Head Email / BusinessHeadEmail / BH Email
  CC Email / CC Emails / CCEmail / CC
  Batcher Code / BatcherCode / Employee Code
  Batcher Name / BatcherName / Employee Name

Sheet IDs and tab names are read from module_settings (prefix "sla."):
  sla.gsheet_id           — Google Sheet ID (same sheet as TP/BTRTP if applicable)
  sla.plant_mapping_tab   — tab name for plant mapping (default: "SLA Plant Mapping")
  sla.batcher_mapping_tab — tab name for batcher mapping (default: "SLA Batcher Mapping")

If those tabs don't exist the user must configure them or add them to the
existing shared Google Sheet.
"""

import pandas as pd
import database
import google_sheets


# ── Fuzzy column finders ─────────────────────────────────────────────────────

def _find_col(df: pd.DataFrame, *candidates) -> str | None:
    """Return the first matching column name (case-insensitive, ignoring spaces/underscores)."""
    norm = {c.lower().replace(" ", "").replace("_", ""): c for c in df.columns}
    for cand in candidates:
        key = cand.lower().replace(" ", "").replace("_", "")
        if key in norm:
            return norm[key]
    return None


def _col_val(row, *candidates):
    for cand in candidates:
        v = row.get(cand, "")
        if v and str(v).strip():
            return str(v).strip()
    return ""


# ── Sheet fetching ────────────────────────────────────────────────────────────

def _get_cfg():
    def _ms(k, d=""):
        return (database.get_module_setting("sla", k, d) or d).strip()
    sheet_id   = _ms("gsheet_id", database.get_setting("tp.gsheet_id", "") or "")
    plant_tab  = _ms("plant_mapping_tab",   "SLA Plant Mapping")
    batcher_tab = _ms("batcher_mapping_tab", "SLA Batcher Mapping")
    return sheet_id, plant_tab, batcher_tab


def _fetch_tab(sheet_id: str, tab: str) -> pd.DataFrame:
    if not sheet_id:
        return pd.DataFrame()
    try:
        return google_sheets.fetch_sheet_tab(sheet_id, tab)
    except Exception as exc:
        print(f"[sla-mapping] could not fetch tab '{tab}': {exc}")
        return pd.DataFrame()


# ── Public API ────────────────────────────────────────────────────────────────

def get_plant_mapping() -> dict:
    """
    Return dict keyed by plant_code:
    {
      "P001": {
         "plant_name": "...",
         "pm_email": "...",
         "bm_email": "...",
         "bh_email": "...",
         "cc_emails": "...",
      },
      ...
    }
    """
    sheet_id, plant_tab, _ = _get_cfg()
    df = _fetch_tab(sheet_id, plant_tab)
    mapping = {}
    if df.empty:
        return mapping

    col_code = _find_col(df, "Plant Code", "PlantCode", "Plant_Code", "Code")
    col_name = _find_col(df, "Plant Name", "PlantName", "Plant_Name", "Name")
    col_pm   = _find_col(df, "Plant Manager Email", "PM Email", "PlantManagerEmail", "PMEmail")
    col_bm   = _find_col(df, "Business Manager Email", "BM Email", "BusinessManagerEmail", "BMEmail")
    col_bh   = _find_col(df, "Business Head Email", "BH Email", "BusinessHeadEmail", "BHEmail")
    col_cc   = _find_col(df, "CC Email", "CC Emails", "CCEmail", "CC", "CCEmails")

    for _, row in df.iterrows():
        code = str(row.get(col_code, "") if col_code else "").strip()
        if not code:
            continue
        mapping[code] = {
            "plant_name": str(row.get(col_name, "") if col_name else "").strip(),
            "pm_email":   str(row.get(col_pm,   "") if col_pm   else "").strip(),
            "bm_email":   str(row.get(col_bm,   "") if col_bm   else "").strip(),
            "bh_email":   str(row.get(col_bh,   "") if col_bh   else "").strip(),
            "cc_emails":  str(row.get(col_cc,   "") if col_cc   else "").strip(),
        }
    return mapping


def get_batcher_mapping() -> dict:
    """Return {batcher_code: batcher_name} from the batcher mapping tab."""
    sheet_id, _, batcher_tab = _get_cfg()
    df = _fetch_tab(sheet_id, batcher_tab)
    mapping = {}
    if df.empty:
        # Fallback: use existing btrtp_master_data from the local DB
        try:
            import database as _db
            conn = _db.get_connection()
            try:
                cur = conn.execute("SELECT batcher_id, batcher_name FROM btrtp_master_data")
                for r in cur.fetchall():
                    if r[0]:
                        mapping[str(r[0]).strip()] = str(r[1] or "").strip()
            finally:
                conn.close()
        except Exception:
            pass
        return mapping

    col_code = _find_col(df, "Batcher Code", "BatcherCode", "Employee Code", "EmployeeCode", "Code")
    col_name = _find_col(df, "Batcher Name", "BatcherName", "Employee Name", "EmployeeName", "Name")
    for _, row in df.iterrows():
        code = str(row.get(col_code, "") if col_code else "").strip()
        name = str(row.get(col_name, "") if col_name else "").strip()
        if code:
            mapping[code] = name
    return mapping


def apply_mappings(df: pd.DataFrame,
                   plant_map: dict,
                   batcher_map: dict) -> pd.DataFrame:
    """
    Enrich a DataFrame that has plant_code and batcher_code with:
      plant_name, batcher_name, pm_email, bm_email, bh_email, cc_emails
    """
    df = df.copy()

    def _pm(code): return plant_map.get(str(code), {}).get("plant_name", str(code))
    def _pm_e(code): return plant_map.get(str(code), {}).get("pm_email", "")
    def _bm_e(code): return plant_map.get(str(code), {}).get("bm_email", "")
    def _bh_e(code): return plant_map.get(str(code), {}).get("bh_email", "")
    def _cc_e(code): return plant_map.get(str(code), {}).get("cc_emails", "")
    def _bn(code):   return batcher_map.get(str(code), str(code))

    df["plant_name"]   = df["plant_code"].apply(_pm)
    df["batcher_name"] = df["batcher_code"].apply(_bn)
    df["pm_email"]     = df["plant_code"].apply(_pm_e)
    df["bm_email"]     = df["plant_code"].apply(_bm_e)
    df["bh_email"]     = df["plant_code"].apply(_bh_e)
    df["cc_emails"]    = df["plant_code"].apply(_cc_e)
    return df


def get_bh_plant_map(plant_map: dict) -> dict:
    """
    Return {bh_email: [plant_code, ...]} — for daily summary where BH receives
    summary of all plants under them.
    """
    bh_map = {}
    for code, info in plant_map.items():
        bh = info.get("bh_email", "").strip()
        if bh:
            bh_map.setdefault(bh, []).append(code)
    return bh_map
