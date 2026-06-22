"""
btrtp_calculator.py
===================
Calculation logic for RDC-BTRTP (Batcher-wise Throughput).

Same TP% formula as RDC-TP but aggregated per individual batcher × plant:

    avg_rate      = total_quantity / (total_time_min / 60)
    throughput_% = avg_rate / mixer_theo_cap × 100

Batch parsing is shared with tp_calculator._parse_batch.
Plant master data (mixer capacity) is shared with RDC-TP (tp_plant_data table).
"""

from datetime import datetime

import pandas as pd

import database
from tp_calculator import _parse_batch


# ── Oracle data parsing ───────────────────────────────────────────────────────

def parse_btrtp_oracle_df(raw_df: pd.DataFrame) -> tuple:
    """
    Clean a raw Oracle BTRTP DataFrame (includes batcher_id column).
    Applies the same time/quantity rules as tp_calculator.parse_oracle_df.
    Returns (parsed_rows, skip_log).
    """
    rows = []
    skip_log = []

    for _, r in raw_df.iterrows():
        batch_ref  = str(r.get("batch_ref", "")).strip()
        batcher_id = str(r.get("batcher_id", "")).strip()

        if not batch_ref or batch_ref == "nan":
            skip_log.append({"reason": "Blank batch reference", "batch_ref": "", "batcher_id": batcher_id})
            continue
        if not batcher_id or batcher_id == "nan":
            skip_log.append({"reason": "Blank batcher ID", "batch_ref": batch_ref, "batcher_id": ""})
            continue

        try:
            time_min = abs(float(r.get("time_taken_min", 0)))
        except (TypeError, ValueError):
            time_min = 0.0

        qty = float(r.get("quantity", 0) or 0)

        if time_min == 0:
            skip_log.append({"reason": "Time taken = 0", "batch_ref": batch_ref, "batcher_id": batcher_id})
            continue
        if time_min > 100:
            skip_log.append({"reason": f"Time > 100 min ({time_min:.1f})", "batch_ref": batch_ref, "batcher_id": batcher_id})
            continue
        if qty <= 0:
            skip_log.append({"reason": f"Quantity ≤ 0 ({qty})", "batch_ref": batch_ref, "batcher_id": batcher_id})
            continue

        plant_code, mixer_variant, lookup_code = _parse_batch(batch_ref)

        rows.append({
            "production_date": str(r.get("production_date", "")),
            "batcher_id":      batcher_id,
            "plant_code":      plant_code,
            "mixer_variant":   mixer_variant or "",
            "lookup_code":     lookup_code,
            "batch_ref":       batch_ref,
            "quantity":        qty,
            "time_taken_min":  time_min,
        })

    return rows, skip_log


# ── Main calculation ──────────────────────────────────────────────────────────

def run_btrtp_calculation(month: int, year: int,
                           from_date: str = None, to_date: str = None) -> tuple:
    """
    Calculate batcher-wise throughput for the given date range.
    Returns (batcher_rows, warnings).

    batcher_rows — list of dicts, one row per (batcher × plant/mixer combo).
    """
    warnings = []

    # 1. Load BTRTP Oracle data from shared cache, fall back to legacy table
    conn = database.get_connection()
    try:
        fd_str = str(from_date) if from_date else "0000-00-00"
        td_str = str(to_date)   if to_date   else "9999-12-31"
        ora_df = pd.read_sql_query(
            """SELECT production_date, created_by AS batcher_id,
                      plant_code, batch_ref, quantity, time_taken_min
               FROM oracle_raw_data
               WHERE production_date >= ? AND production_date <= ?""",
            conn, params=(fd_str, td_str)
        )
        if ora_df.empty:
            ora_df = pd.read_sql_query(
                """SELECT production_date, batcher_id, plant_code,
                          batch_ref, quantity, time_taken_min
                   FROM btrtp_oracle_data
                   WHERE production_date >= ? AND production_date <= ?""",
                conn, params=(fd_str, td_str)
            )
    finally:
        conn.close()
    if ora_df.empty:
        warnings.append("No Oracle data found — fetch from Oracle first.")
        return [], warnings

    # 2. Load TP plant master (mixer capacity — shared with RDC-TP)
    plant_df = database.read_table("tp_plant_data")
    if plant_df.empty:
        warnings.append("No Plant Data found — sync Plant Data (via RDC-TP Data Uploader) first.")
        return [], warnings

    plant_map = {str(r["plant_code"]): r.to_dict() for _, r in plant_df.iterrows()}

    # 3. Load BT Master (batcher_id → batcher_name)
    master_df = database.read_table("btrtp_master_data")
    master_map = {}
    if not master_df.empty:
        for _, r in master_df.iterrows():
            master_map[str(r["batcher_id"]).strip().upper()] = str(r.get("batcher_name", ""))

    # 4. Aggregate by (batcher_id, lookup_code)
    ora_df["quantity"]       = pd.to_numeric(ora_df["quantity"],       errors="coerce").fillna(0)
    ora_df["time_taken_min"] = pd.to_numeric(ora_df["time_taken_min"], errors="coerce").fillna(0)

    # Derive lookup_code from batch_ref if not present (shared oracle_raw_data)
    if "lookup_code" not in ora_df.columns:
        ora_df["lookup_code"] = ora_df["batch_ref"].apply(
            lambda b: _parse_batch(str(b))[2]
        )

    # Drop zero-time and zero-qty rows
    ora_df = ora_df[(ora_df["time_taken_min"] > 0) & (ora_df["time_taken_min"] <= 100)
                    & (ora_df["quantity"] > 0)]

    grouped = (
        ora_df.groupby(["batcher_id", "lookup_code"], sort=True)
        .agg(
            total_quantity=("quantity",       "sum"),
            total_time_min=("time_taken_min", "sum"),
            batch_count=("quantity",          "count"),
        )
        .reset_index()
    )

    now = datetime.now().isoformat(timespec="seconds")
    batcher_rows = []

    for _, row in grouped.iterrows():
        batcher_id     = str(row["batcher_id"]).strip()
        lookup         = str(row["lookup_code"]).strip()
        total_qty      = float(row["total_quantity"])
        total_time_min = float(row["total_time_min"])
        batch_count    = int(row["batch_count"])

        info = plant_map.get(lookup)
        if not info:
            base = lookup.split("_")[0]
            info = plant_map.get(base)
        if not info:
            warnings.append(f"Plant '{lookup}' not found in Plant Data — skipped.")
            continue

        total_time_hrs = total_time_min / 60.0
        if total_time_hrs <= 0:
            continue

        mixer_cap = float(info.get("mixer_theo_cap") or 0)
        if mixer_cap <= 0:
            warnings.append(f"Mixer capacity = 0 for '{lookup}' — TP% set to 0%.")
            throughput_pct = 0.0
        else:
            avg_rate       = total_qty / total_time_hrs
            throughput_pct = (avg_rate / mixer_cap) * 100.0

        batcher_name = master_map.get(batcher_id.upper(), batcher_id)

        batcher_rows.append({
            "month":          month,
            "year":           year,
            "batcher_id":     batcher_id,
            "batcher_name":   batcher_name,
            "lookup_code":    lookup,
            "plant_name":     str(info.get("plant_name", "")),
            "exco_location":  str(info.get("exco_location", "")),
            "business_head":  str(info.get("business_head", "")),
            "plant_manager":  str(info.get("plant_manager", "")),
            "mixer_theo_cap": mixer_cap,
            "total_quantity": round(total_qty, 2),
            "total_time_hrs": round(total_time_hrs, 3),
            "throughput_pct": round(throughput_pct, 2),
            "batch_count":    batch_count,
            "generated_at":   now,
        })

    # Sort by plant then TP% descending (highest performer first within each plant)
    batcher_rows.sort(key=lambda r: (r["plant_name"], -r["throughput_pct"]))
    return batcher_rows, warnings


# ── Save results ──────────────────────────────────────────────────────────────

def save_btrtp_results(batcher_rows: list, month: int, year: int) -> int:
    """Delete existing results for this month/year and insert new ones."""
    conn = database.get_connection()
    try:
        cur = conn.cursor()
        cur.execute("DELETE FROM btrtp_results WHERE month = ? AND year = ?", (month, year))
        conn.commit()
    finally:
        conn.close()
    return database.insert_rows("btrtp_results", batcher_rows)
