"""
mapping_service.py
==================
Reads the app's own database to provide:
  - Plant Code  -> Plant Name
  - Plant Code  -> Plant Manager Email
  - Plant Code  -> BM/EA to MD Email
  - Plant Code  -> Business Head Email
  - Plant Code  -> CC Emails
  - Batcher Code -> Batcher Name

Plant details and e-mails live in the tp_plant_data table (edit them in
RDC-TP > Data Uploader > Plant Data). Batcher names come from the RDC-I&D
Master Data (master_data table). There is no Google Sheet dependency.
"""

import pandas as pd
import database


# -- Plant mapping ------------------------------------------------------------

def _strip_variant(code: str) -> str:
    """Return the bare plant code without a _BPn mixer-variant suffix.
    'MU7_BP1' -> 'MU7'  ;  'MU7' -> 'MU7'."""
    c = str(code or "").strip()
    if "_" in c:
        return c.split("_")[0]
    return c


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
      "MU7_BP1": {
         "plant_name": "...", "exco_location": "...", "mixer_capacity": 7.0,
         "pm_email": "...", "bm_email": "...", "bh_email": "...", "cc_emails": "...",
      },
      ...
    }

    Oracle reports the BARE plant code (e.g. 'MU7'), but the plant master stores
    plants under mixer-variant codes (e.g. 'MU7_BP1', 'MU7_BP2'). So in addition
    to the exact codes we ALSO register the bare code as an alias that points to
    the first variant's details (all variants of a plant share one location and
    the same managers). Mixer capacity for the bare alias is the largest variant
    capacity; the exact variant code keeps its own.
    """
    mapping = {}
    bare_groups = {}  # bare_code -> list of records
    try:
        conn = database.get_connection()
        try:
            cur = conn.execute(
                "SELECT plant_code, plant_name, exco_location, mixer_theo_cap, "
                "pm_email, bm_email, bh_email, cc_emails FROM tp_plant_data"
            )
            for code, name, loc, cap, pm, bm, bh, cc in cur.fetchall():
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
                    "pm_email":       (pm or "").strip(),
                    "bm_email":       (bm or "").strip(),
                    "bh_email":       (bh or "").strip(),
                    "cc_emails":      (cc or "").strip(),
                    "mixer_capacity": mc,
                }
                mapping[c] = rec
                bare_groups.setdefault(_strip_variant(c), []).append(rec)
        finally:
            conn.close()
    except Exception as exc:
        print(f"[sla-mapping] could not read plant data: {exc}")
        return mapping

    # Register bare-code aliases (only when the bare code isn't already a real row)
    for bare, recs in bare_groups.items():
        if bare in mapping:
            continue
        first = recs[0]
        caps = [r["mixer_capacity"] for r in recs if r["mixer_capacity"] is not None]
        mapping[bare] = {
            "plant_name":     _strip_variant_name(first["plant_name"]) or bare,
            "exco_location":  next((r["exco_location"] for r in recs if r["exco_location"]), ""),
            "pm_email":       first["pm_email"],
            "bm_email":       first["bm_email"],
            "bh_email":       first["bh_email"],
            "cc_emails":      first["cc_emails"],
            "mixer_capacity": max(caps) if caps else None,
        }
    return mapping


# -- Batcher mapping ----------------------------------------------------------

def get_batcher_mapping() -> dict:
    """Return {batcher_code: batcher_name} from the I&D Master Data."""
    mapping = {}
    try:
        df = database.get_batcher_master()
        for _, r in df.iterrows():
            code = str(r["batcher_id"] or "").strip()
            if code:
                mapping[code] = str(r["batcher_name"] or "").strip()
    except Exception as exc:
        print(f"[sla-mapping] could not read batcher names: {exc}")
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
