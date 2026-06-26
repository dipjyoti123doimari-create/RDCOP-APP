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

def _strip_variant(code: str) -> str:
    """Return the bare plant code without a _BPn mixer-variant suffix.
    'MU7_BP1' -> 'MU7'  ;  'MU7' -> 'MU7'."""
    c = str(code or "").strip()
    if "_" in c:
        return c.split("_")[0]
    return c


def _fallback_plant_mapping_from_tp() -> dict:
    """
    Build plant mapping from the existing TP plant master (tp_plant_data).
    Provides plant_name + exco_location + mixer_capacity per plant code.

    Oracle reports the BARE plant code (e.g. 'MU7'), but the TP master stores
    plants under mixer-variant codes (e.g. 'MU7_BP1', 'MU7_BP2'). So in addition
    to the exact variant codes we ALSO register the bare code as an alias that
    points to the variant's location/name (all variants of a plant share one
    Exco Location). Mixer capacity for the bare alias uses the largest variant
    capacity (representative of the plant); the exact variant code keeps its own.

    Email fields are blank — they require the SLA Plant Mapping sheet tab.
    """
    mapping = {}
    bare_groups = {}  # bare_code -> list of (name, location, mixer_cap)
    try:
        conn = database.get_connection()
        try:
            cur = conn.execute(
                "SELECT plant_code, plant_name, exco_location, mixer_theo_cap "
                "FROM tp_plant_data"
            )
            for code, name, loc, cap in cur.fetchall():
                if not code:
                    continue
                c = str(code).strip()
                try:
                    mc = float(cap) if cap is not None else None
                except (TypeError, ValueError):
                    mc = None
                rec = {
                    "plant_name":     (name or "").strip(),
                    "exco_location":  (loc or "").strip(),
                    "pm_email":       "",   # TP master has names, not emails
                    "bm_email":       "",
                    "bh_email":       "",
                    "cc_emails":      "",
                    "mixer_capacity": mc,
                }
                mapping[c] = rec
                bare_groups.setdefault(_strip_variant(c), []).append(rec)
        finally:
            conn.close()
    except Exception as exc:
        print(f"[sla-mapping] TP fallback failed: {exc}")
        return mapping

    # Register bare-code aliases (only when the bare code isn't already a real row)
    for bare, recs in bare_groups.items():
        if bare in mapping:
            continue
        # Plant name without the _BPn suffix; first non-empty location.
        name = ""
        loc = ""
        caps = []
        for r in recs:
            if not name and r["plant_name"]:
                name = _strip_variant_name(r["plant_name"])
            if not loc and r["exco_location"]:
                loc = r["exco_location"]
            if r["mixer_capacity"] is not None:
                caps.append(r["mixer_capacity"])
        mapping[bare] = {
            "plant_name":     name or bare,
            "exco_location":  loc,
            "pm_email":       "",
            "bm_email":       "",
            "bh_email":       "",
            "cc_emails":      "",
            # representative capacity for the bare code = largest variant cap
            "mixer_capacity": max(caps) if caps else None,
        }
    return mapping


def _strip_variant_name(name: str) -> str:
    """'MUM-Kashimira_BP2' -> 'MUM-Kashimira'."""
    n = str(name or "").strip()
    if "_BP" in n:
        return n.split("_BP")[0]
    return n


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

    col_code = _find_col(df, "Plant Code", "PlantCode", "Plant_Code") if not df.empty else None
    col_name = _find_col(df, "Plant Name", "PlantName", "Plant_Name") if not df.empty else None
    col_pm   = _find_col(df, "Plant Manager Email", "PM Email", "PlantManagerEmail", "PMEmail") if not df.empty else None
    col_bm   = _find_col(df, "Business Manager Email", "BM Email", "BusinessManagerEmail", "BMEmail") if not df.empty else None
    col_bh   = _find_col(df, "Business Head Email", "BH Email", "BusinessHeadEmail", "BHEmail") if not df.empty else None
    col_cc   = _find_col(df, "CC Email", "CC Emails", "CCEmail", "CC", "CCEmails") if not df.empty else None
    col_mc   = _find_col(df, "Mixer Capacity", "MixerCapacity", "Mixer_Capacity",
                         "Mixer Cap", "MixerCap") if not df.empty else None

    # Guard against the public-CSV fallback returning the WRONG tab (e.g. the
    # default employee-master tab) when "SLA Plant Mapping" doesn't exist.
    # A valid SLA plant tab must have at least a plant code AND one email column.
    is_valid_sla_tab = bool(col_code and (col_pm or col_bm or col_bh))
    mapping = {}

    if not is_valid_sla_tab:
        # No proper SLA sheet tab → fall back to the existing TP plant master
        # for plant name + mixer capacity. Emails will be blank (must add the
        # SLA Plant Mapping tab to enable alert recipients).
        return _fallback_plant_mapping_from_tp()

    # TP master mixer caps — used to fill any plant the sheet leaves blank.
    tp_fallback = _fallback_plant_mapping_from_tp()

    for _, row in df.iterrows():
        code = str(row.get(col_code, "") if col_code else "").strip()
        if not code:
            continue
        # Mixer capacity may be blank in the sheet — parse leniently, then fall
        # back to the TP plant master capacity for that plant code.
        mc_raw = str(row.get(col_mc, "") if col_mc else "").strip()
        try:
            mixer_cap = float(mc_raw) if mc_raw else None
        except (TypeError, ValueError):
            mixer_cap = None
        if mixer_cap is None:
            mixer_cap = tp_fallback.get(code, {}).get("mixer_capacity")
        mapping[code] = {
            "plant_name":     str(row.get(col_name, "") if col_name else "").strip()
                              or tp_fallback.get(code, {}).get("plant_name", ""),
            "pm_email":       str(row.get(col_pm,   "") if col_pm   else "").strip(),
            "bm_email":       str(row.get(col_bm,   "") if col_bm   else "").strip(),
            "bh_email":       str(row.get(col_bh,   "") if col_bh   else "").strip(),
            "cc_emails":      str(row.get(col_cc,   "") if col_cc   else "").strip(),
            "mixer_capacity": mixer_cap,
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


def _plant_lookup(plant_map: dict, code) -> dict:
    """
    Resolve a plant_code to its mapping record, tolerant of bare↔variant
    mismatches between Oracle (bare 'MU7') and the TP master ('MU7_BP1'):
      1. exact code
      2. bare code (suffix stripped)
      3. any variant whose bare code matches (first one found)
    Returns {} if nothing matches.
    """
    c = str(code).strip()
    if c in plant_map:
        return plant_map[c]
    bare = _strip_variant(c)
    if bare in plant_map:
        return plant_map[bare]
    # Last resort: scan for any code sharing the same bare prefix.
    for k, v in plant_map.items():
        if _strip_variant(k) == bare:
            return v
    return {}


def get_location_map() -> dict:
    """
    Return {plant_code: exco_location} including bare-code aliases, so callers
    can group SLA records by location. Sourced from the plant mapping (which
    falls back to the TP master with variant handling).
    """
    pm = get_plant_mapping()
    return {code: info.get("exco_location", "") for code, info in pm.items()}


def resolve_location(plant_map: dict, code) -> str:
    """Return the Exco Location for a plant code (bare/variant tolerant)."""
    return _plant_lookup(plant_map, code).get("exco_location", "")


def apply_mappings(df: pd.DataFrame,
                   plant_map: dict,
                   batcher_map: dict) -> pd.DataFrame:
    """
    Enrich a DataFrame that has plant_code and batcher_code with:
      plant_name, exco_location, batcher_name, pm_email, bm_email,
      bh_email, cc_emails, mixer_capacity
    Tolerant of bare↔variant plant-code differences (see _plant_lookup).
    """
    df = df.copy()

    def _info(code): return _plant_lookup(plant_map, code)
    def _pm(code):   return _info(code).get("plant_name") or str(code)
    def _loc(code):  return _info(code).get("exco_location", "")
    def _pm_e(code): return _info(code).get("pm_email", "")
    def _bm_e(code): return _info(code).get("bm_email", "")
    def _bh_e(code): return _info(code).get("bh_email", "")
    def _cc_e(code): return _info(code).get("cc_emails", "")
    def _bn(code):   return batcher_map.get(str(code), str(code))

    def _mc(code, existing):
        # Mixer capacity: prefer the sheet/TP value (bare/variant tolerant),
        # else keep whatever was already on the row.
        cap = _info(code).get("mixer_capacity")
        return cap if cap is not None else existing

    df["plant_name"]    = df["plant_code"].apply(_pm)
    df["exco_location"] = df["plant_code"].apply(_loc)
    df["batcher_name"]  = df["batcher_code"].apply(_bn)
    df["pm_email"]      = df["plant_code"].apply(_pm_e)
    df["bm_email"]      = df["plant_code"].apply(_bm_e)
    df["bh_email"]      = df["plant_code"].apply(_bh_e)
    df["cc_emails"]     = df["plant_code"].apply(_cc_e)
    if "mixer_capacity" in df.columns:
        df["mixer_capacity"] = df.apply(
            lambda r: _mc(r["plant_code"], r.get("mixer_capacity")), axis=1)
    else:
        df["mixer_capacity"] = df["plant_code"].apply(lambda c: _mc(c, None))
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
