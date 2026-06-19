"""
ecmd_calculator.py
==================
Calculation logic for RDC-ECMD (Energy Consumption & Mixer DG Ratio).

For each plant / month the formula chain is:
  EB KWh      = (eb_kwh_close - eb_kwh_open) * MF
  DG KWh      = (dg_kwh_close - dg_kwh_open)  if DG KWh readings given
              = DG running hrs * 20             otherwise
  Total KWh   = EB KWh + DG KWh
  DG run hrs  = dg_hr_close - dg_hr_open
  Mixer DG hrs= mixer_dg_hr_close - mixer_dg_hr_open
  Mixer DG %  = (Mixer DG hrs / DG run hrs) * 100
  L / hr      = diesel_issued_ltrs / DG run hrs
  Energy/MT   = Total KWh / Total Volume  (volume from Oracle by plant code)
"""

from __future__ import annotations

from datetime import date as _date
from typing import Optional

import database
import oracle_connector


def _f(v, default=0.0) -> float:
    """Safe float conversion with default."""
    try:
        return float(v) if v not in (None, "") else default
    except (TypeError, ValueError):
        return default


def _oracle_volume_for_plant(plant_code: str, from_date: str, to_date: str) -> float:
    """
    Fetch total production quantity from Oracle for a plant in the date range.
    Returns 0.0 if Oracle is not configured or the plant has no data.
    Uses the same tp_oracle_data table (already fetched by TP module).
    Falls back to in-memory Oracle fetch if not available in local store.
    """
    try:
        import pandas as pd
        conn = database.get_connection()
        try:
            df = pd.read_sql_query(
                "SELECT SUM(quantity) AS qty FROM tp_oracle_data "
                "WHERE plant_code = ? AND production_date BETWEEN ? AND ?",
                conn,
                params=(str(plant_code).strip(), from_date, to_date),
            )
        finally:
            conn.close()
        val = df["qty"].iloc[0] if not df.empty else None
        if val and float(val) > 0:
            return float(val)
    except Exception:
        pass
    return 0.0


def calculate_plant_row(
    reading: dict,
    plant_info: dict,
    from_date: str,
    to_date: str,
    oracle_vol_map: dict,
) -> dict:
    """
    Compute all ECMD metrics for one plant reading.

    Parameters
    ----------
    reading     : ecmd_readings row dict
    plant_info  : tp_plant_data row dict (name, exco, BH, PM)
    from_date   : ISO date string (period start)
    to_date     : ISO date string (period end)
    oracle_vol_map : {plant_code: total_volume} pre-fetched from Oracle

    Returns a result dict ready to insert into ecmd_results.
    """
    pc = str(reading.get("plant_code", "")).strip()

    # ── EB KWh ────────────────────────────────────────────────────────────────
    eb_open  = _f(reading.get("eb_kwh_open"))
    eb_close = _f(reading.get("eb_kwh_close"))
    mf       = _f(reading.get("mf"), 1.0) or 1.0
    eb_kwh   = (eb_close - eb_open) * mf

    # ── DG KWh ────────────────────────────────────────────────────────────────
    dg_hr_open  = _f(reading.get("dg_hr_open"))
    dg_hr_close = _f(reading.get("dg_hr_close"))
    dg_run_hrs  = max(dg_hr_close - dg_hr_open, 0.0)

    dg_kwh_open  = reading.get("dg_kwh_open")
    dg_kwh_close = reading.get("dg_kwh_close")
    if dg_kwh_open not in (None, "") and dg_kwh_close not in (None, ""):
        dg_kwh = max(_f(dg_kwh_close) - _f(dg_kwh_open), 0.0)
    else:
        dg_kwh = dg_run_hrs * 20.0

    # ── Total energy ──────────────────────────────────────────────────────────
    total_kwh = max(eb_kwh, 0.0) + dg_kwh

    # ── Mixer DG ratio ────────────────────────────────────────────────────────
    mx_open  = _f(reading.get("mixer_dg_hr_open"))
    mx_close = _f(reading.get("mixer_dg_hr_close"))
    mixer_dg_hrs  = max(mx_close - mx_open, 0.0)
    mixer_dg_ratio = (mixer_dg_hrs / dg_run_hrs * 100.0) if dg_run_hrs > 0 else 0.0

    # ── Diesel L/hr ───────────────────────────────────────────────────────────
    diesel_ltrs = _f(reading.get("diesel_issued_ltrs"))
    ltr_per_hr  = (diesel_ltrs / dg_run_hrs) if dg_run_hrs > 0 else 0.0

    # ── Volume & energy per MT ─────────────────────────────────────────────────
    volume_on_dg = _f(reading.get("volume_on_dg"))
    total_volume = oracle_vol_map.get(pc, 0.0)
    energy_per_mt = (total_kwh / total_volume) if total_volume > 0 else 0.0

    # ── Plant reference info ──────────────────────────────────────────────────
    month = int(reading.get("month", 0))
    year  = int(reading.get("year",  0))

    return {
        "month":            month,
        "year":             year,
        "plant_code":       pc,
        "plant_name":       plant_info.get("plant_name", pc),
        "exco_location":    plant_info.get("exco_location", ""),
        "business_head":    plant_info.get("business_head", ""),
        "plant_manager":    plant_info.get("plant_manager", ""),
        "eb_kwh":           round(eb_kwh,        2),
        "dg_kwh":           round(dg_kwh,        2),
        "total_kwh":        round(total_kwh,     2),
        "total_volume":     round(total_volume,  2),
        "energy_per_mt":    round(energy_per_mt, 4),
        "dg_run_hrs":       round(dg_run_hrs,    2),
        "mixer_dg_hrs":     round(mixer_dg_hrs,  2),
        "mixer_dg_ratio":   round(mixer_dg_ratio, 2),
        "diesel_issued_ltrs": round(diesel_ltrs, 2),
        "ltr_per_hr":       round(ltr_per_hr,    3),
        "volume_on_dg":     round(volume_on_dg,  2),
        "generated_at":     _date.today().isoformat(),
    }


def run_ecmd_calculation(
    month: int,
    year: int,
    from_date: str,
    to_date: str,
) -> tuple[list, list]:
    """
    Run the full ECMD calculation for the given month/year.

    Returns
    -------
    (result_rows, warnings)
      result_rows : list of dicts ready for ecmd_results table
      warnings    : list of human-readable warning strings
    """
    warnings: list[str] = []

    # Load all readings for this month
    readings = database.get_ecmd_readings_for_month(month, year)
    if not readings:
        warnings.append(f"No energy readings found for {month}/{year}. Enter readings first.")
        return [], warnings

    # Build plant info lookup from tp_plant_data
    plant_rows = database.get_tp_plants()
    plant_map  = {p["plant_code"]: p for p in plant_rows}

    # Pre-fetch Oracle volumes for all plant codes in one pass
    plant_codes = [r["plant_code"] for r in readings]
    oracle_vol_map: dict = {}
    if oracle_connector.is_configured():
        for pc in plant_codes:
            vol = _oracle_volume_for_plant(pc, from_date, to_date)
            oracle_vol_map[pc] = vol
            if vol == 0.0:
                warnings.append(f"Plant {pc}: no Oracle volume found for {from_date} → {to_date}. Energy/MT will be 0.")
    else:
        # Try to use locally saved tp_oracle_data
        for pc in plant_codes:
            vol = _oracle_volume_for_plant(pc, from_date, to_date)
            oracle_vol_map[pc] = vol
        if not any(oracle_vol_map.values()):
            warnings.append("Oracle not configured and no saved TP Oracle data found. "
                            "Total Volume will be 0 — Energy/MT cannot be calculated.")

    results: list[dict] = []
    for reading in readings:
        pc = str(reading.get("plant_code", "")).strip()
        plant_info = plant_map.get(pc, {"plant_name": pc, "exco_location": "",
                                        "business_head": "", "plant_manager": ""})
        if pc not in plant_map:
            warnings.append(f"Plant code '{pc}' not found in TP Plant Data — "
                            "plant name/location will be blank.")
        row = calculate_plant_row(reading, plant_info, from_date, to_date, oracle_vol_map)
        results.append(row)

    # Sort: by exco_location then plant_name
    results.sort(key=lambda r: (r.get("exco_location", ""), r.get("plant_name", "")))
    return results, warnings


def build_location_summary(plant_rows: list) -> list:
    """
    Aggregate plant-level rows into location-level summary rows.
    Returns list of location dicts, with a PAN India total appended.
    """
    from collections import defaultdict
    loc: dict = defaultdict(lambda: {
        "exco_location": "",
        "plant_count":   0,
        "total_kwh":     0.0,
        "total_volume":  0.0,
        "dg_run_hrs":    0.0,
        "mixer_dg_hrs":  0.0,
        "diesel_ltrs":   0.0,
        "volume_on_dg":  0.0,
    })

    for r in plant_rows:
        key = r.get("exco_location", "Unknown")
        g = loc[key]
        g["exco_location"] = key
        g["plant_count"]  += 1
        g["total_kwh"]    += _f(r.get("total_kwh"))
        g["total_volume"] += _f(r.get("total_volume"))
        g["dg_run_hrs"]   += _f(r.get("dg_run_hrs"))
        g["mixer_dg_hrs"] += _f(r.get("mixer_dg_hrs"))
        g["diesel_ltrs"]  += _f(r.get("diesel_issued_ltrs"))
        g["volume_on_dg"] += _f(r.get("volume_on_dg"))

    rows = sorted(loc.values(), key=lambda x: x["exco_location"])
    for g in rows:
        g["energy_per_mt"]   = round(g["total_kwh"] / g["total_volume"], 4) if g["total_volume"] > 0 else 0.0
        g["mixer_dg_ratio"]  = round(g["mixer_dg_hrs"] / g["dg_run_hrs"] * 100, 2) if g["dg_run_hrs"] > 0 else 0.0
        g["ltr_per_hr"]      = round(g["diesel_ltrs"] / g["dg_run_hrs"], 3) if g["dg_run_hrs"] > 0 else 0.0
        g["total_kwh"]       = round(g["total_kwh"],    2)
        g["total_volume"]    = round(g["total_volume"], 2)
        g["dg_run_hrs"]      = round(g["dg_run_hrs"],   2)
        g["mixer_dg_hrs"]    = round(g["mixer_dg_hrs"], 2)
        g["diesel_ltrs"]     = round(g["diesel_ltrs"],  2)
        g["volume_on_dg"]    = round(g["volume_on_dg"], 2)
        g["is_pan_india"]    = False

    # PAN India aggregate
    if rows:
        pan = {
            "exco_location": "PAN India",
            "plant_count":   sum(g["plant_count"]  for g in rows),
            "total_kwh":     round(sum(g["total_kwh"]    for g in rows), 2),
            "total_volume":  round(sum(g["total_volume"] for g in rows), 2),
            "dg_run_hrs":    round(sum(g["dg_run_hrs"]   for g in rows), 2),
            "mixer_dg_hrs":  round(sum(g["mixer_dg_hrs"] for g in rows), 2),
            "diesel_ltrs":   round(sum(g["diesel_ltrs"]  for g in rows), 2),
            "volume_on_dg":  round(sum(g["volume_on_dg"] for g in rows), 2),
            "is_pan_india":  True,
        }
        tv = pan["total_volume"]
        dh = pan["dg_run_hrs"]
        mh = pan["mixer_dg_hrs"]
        dl = pan["diesel_ltrs"]
        pan["energy_per_mt"]  = round(pan["total_kwh"] / tv, 4) if tv > 0 else 0.0
        pan["mixer_dg_ratio"] = round(mh / dh * 100, 2)         if dh > 0 else 0.0
        pan["ltr_per_hr"]     = round(dl / dh, 3)                if dh > 0 else 0.0
        rows.append(pan)

    return rows
