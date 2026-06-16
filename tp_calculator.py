"""
tp_calculator.py
================
All calculation logic for the RDC-TP (Plant Throughput) module.

PUBLIC FUNCTIONS
    parse_oracle_df(df)           -> list[dict]   clean & parse raw Oracle rows
    run_tp_calculation(month, year) -> (plant_rows, location_rows, warnings)
    save_tp_results(plant_rows, month, year) -> int

THROUGHPUT FORMULA (per plant/mixer)
    throughput % = (Total Qty / Total Time in Hours) / Mixer Theo. Capacity × 100

TIME TAKEN RULES
    1. Take abs() — negative values are valid, just remove the sign
    2. Skip rows where abs(time) == 0
    3. Skip rows where abs(time) > 100 (outlier batches)
    Both quantity and time are excluded for skipped rows.

BATCH PARSING
    Format A: AY1/2026/2167        → plant_code=AY1, mixer=None,  lookup=AY1
    Format B: AY1/BP1/2026/2167    → plant_code=AY1, mixer=BP1,   lookup=AY1_BP1

LOCATION TP
    Simple average of throughput % of all plants under the same Exco Location.
    (Weighted average not applicable because mixer capacities differ per plant.)
"""

from datetime import datetime

import pandas as pd

import database


# ── Batch parsing ─────────────────────────────────────────────────────────────

def _parse_batch(batch_str: str) -> tuple:
    """
    Parse a batch reference string into (plant_code, mixer_variant, lookup_code).

    Returns:
        plant_code    – base plant code (e.g. "AY1")
        mixer_variant – "BP1"/"BP2"/"BP3" or None
        lookup_code   – key used to join with tp_plant_data (e.g. "AY1" or "AY1_BP1")
    """
    parts = str(batch_str).strip().split("/")
    plant_code = parts[0].strip() if parts else ""

    if len(parts) >= 4:
        # Format B: PLANT/BPn/YEAR/BATCHNUM
        second = parts[1].strip().upper()
        if second.startswith("BP"):
            return plant_code, second, f"{plant_code}_{second}"
        # Unusual 4-part without BP — treat like Format A
        return plant_code, None, plant_code

    # Format A: PLANT/YEAR/BATCHNUM (3 parts) or any other
    return plant_code, None, plant_code


# ── Oracle data cleaning & parsing ───────────────────────────────────────────

def parse_oracle_df(raw_df: pd.DataFrame) -> tuple:
    """
    Clean a raw Oracle DataFrame (from oracle_connector.fetch_tp_data) and
    return (parsed_rows, skip_log) ready for tp_oracle_data table.

    Rules applied:
        - batch_ref must be non-blank
        - time_taken_min: take abs(); skip row if == 0 or > 100
        - quantity: skip row if <= 0
    """
    rows = []
    skip_log = []   # list of dicts for tp_validation_errors style log

    for _, r in raw_df.iterrows():
        batch_ref = str(r.get("batch_ref", "")).strip()
        if not batch_ref or batch_ref == "nan":
            skip_log.append({"reason": "Blank batch reference", "batch_ref": batch_ref})
            continue

        raw_time = r.get("time_taken_min", 0)
        try:
            time_min = abs(float(raw_time))
        except (TypeError, ValueError):
            time_min = 0.0

        qty = float(r.get("quantity", 0) or 0)

        if time_min == 0:
            skip_log.append({"reason": "Time taken = 0", "batch_ref": batch_ref, "time": raw_time})
            continue
        if time_min > 100:
            skip_log.append({"reason": f"Time taken > 100 min ({time_min:.1f})", "batch_ref": batch_ref, "time": raw_time})
            continue
        if qty <= 0:
            skip_log.append({"reason": f"Quantity ≤ 0 ({qty})", "batch_ref": batch_ref, "qty": qty})
            continue

        plant_code, mixer_variant, lookup_code = _parse_batch(batch_ref)

        rows.append({
            "production_date": str(r.get("production_date", "")),
            "plant_code":      plant_code,
            "mixer_variant":   mixer_variant or "",
            "lookup_code":     lookup_code,
            "batch_ref":       batch_ref,
            "quantity":        qty,
            "time_taken_min":  time_min,
        })

    return rows, skip_log


# ── Main calculation ──────────────────────────────────────────────────────────

def run_tp_calculation(month: int, year: int,
                        from_date: str = None, to_date: str = None,
                        ora_df=None) -> tuple:
    """
    Calculate plant-wise and location-wise throughput.

    ora_df – optional in-memory DataFrame (already parsed). When supplied the
             DB is not read at all, enabling historical queries without
             persisting data. When None, data is read from tp_oracle_data.

    Returns:
        plant_rows    – list[dict]  one row per plant/mixer
        location_rows – list[dict]  one row per Exco Location (simple avg)
        warnings      – list[str]
    """
    warnings = []

    # 1. Load oracle data (from memory or DB)
    if ora_df is None:
        ora_df = database.read_table("tp_oracle_data")
        if ora_df.empty:
            warnings.append("No Oracle data found — fetch from Oracle first.")
            return [], [], warnings
        # Filter to the requested date range when reading from DB
        if from_date and to_date:
            pd_str = ora_df["production_date"].astype(str).str.slice(0, 10)
            ora_df = ora_df[(pd_str >= str(from_date)) & (pd_str <= str(to_date))]
            if ora_df.empty:
                warnings.append(f"No Oracle data in range {from_date} → {to_date}.")
                return [], [], warnings
    else:
        if ora_df.empty:
            warnings.append(f"No Oracle data in range {from_date} → {to_date}.")
            return [], [], warnings

    # 2. Load plant reference data
    plant_df = database.read_table("tp_plant_data")
    if plant_df.empty:
        warnings.append("No Plant Data found — sync from Google Sheets first.")
        return [], [], warnings

    plant_map = {
        str(r["plant_code"]): r.to_dict()
        for _, r in plant_df.iterrows()
    }

    # 3. Aggregate: total qty and total time per lookup_code
    ora_df["quantity"]      = pd.to_numeric(ora_df["quantity"],      errors="coerce").fillna(0)
    ora_df["time_taken_min"] = pd.to_numeric(ora_df["time_taken_min"], errors="coerce").fillna(0)

    grouped = (
        ora_df.groupby("lookup_code", sort=True)
        .agg(
            total_quantity=("quantity",      "sum"),
            total_time_min=("time_taken_min","sum"),
            batch_count=("quantity",         "count"),
        )
        .reset_index()
    )

    now = datetime.now().isoformat(timespec="seconds")
    plant_rows = []

    for _, row in grouped.iterrows():
        lookup = str(row["lookup_code"])
        info   = plant_map.get(lookup)

        if not info:
            warnings.append(f"Plant code '{lookup}' not found in Plant Data — skipped.")
            continue

        total_time_hrs = row["total_time_min"] / 60.0
        if total_time_hrs <= 0:
            warnings.append(f"'{lookup}' has zero total time after filtering — skipped.")
            continue

        mixer_cap = float(info.get("mixer_theo_cap") or 0)
        if mixer_cap <= 0:
            warnings.append(f"'{lookup}' has zero Mixer Theo. Capacity — throughput set to 0%.")
            throughput_pct = 0.0
        else:
            avg_rate       = row["total_quantity"] / total_time_hrs   # units/hr
            throughput_pct = (avg_rate / mixer_cap) * 100.0

        plant_rows.append({
            "month":          month,
            "year":           year,
            "lookup_code":    lookup,
            "plant_name":     str(info.get("plant_name", "")),
            "exco_location":  str(info.get("exco_location", "")),
            "business_head":  str(info.get("business_head", "")),
            "plant_manager":  str(info.get("plant_manager", "")),
            "mixer_theo_cap": mixer_cap,
            "total_quantity": round(float(row["total_quantity"]), 2),
            "total_time_min": round(float(row["total_time_min"]), 1),
            "total_time_hrs": round(total_time_hrs, 2),
            "throughput_pct": round(throughput_pct, 2),
            "batch_count":    int(row["batch_count"]),
            "generated_at":   now,
        })

    # 4. Sort plant rows lowest → highest and build location summary
    plant_rows.sort(key=lambda r: r["throughput_pct"])
    location_rows = build_location_rows(plant_rows, month, year)

    return plant_rows, location_rows, warnings


def build_location_rows(plant_rows: list, month: int, year: int) -> list:
    """
    Build Exco-Location rows (simple average of plant throughput %) from a list
    of plant rows, sorted lowest → highest with a PAN India total at the bottom.
    Reused by both the calculation and the filtered Reports view.
    """
    location_map = {}
    for pr in plant_rows:
        loc = pr.get("exco_location", "")
        if not loc:
            continue
        if loc not in location_map:
            location_map[loc] = {"plant_count": 0, "total_throughput_pct": 0.0,
                                 "total_quantity": 0.0}
        location_map[loc]["plant_count"]          += 1
        location_map[loc]["total_throughput_pct"] += pr["throughput_pct"]
        location_map[loc]["total_quantity"]        += pr["total_quantity"]

    location_rows = []
    all_pct_sum = all_qty_sum = 0.0
    all_count = 0
    for loc, d in sorted(location_map.items()):
        cnt = d["plant_count"]
        location_rows.append({
            "exco_location":      loc,
            "plant_count":        cnt,
            "avg_throughput_pct": round(d["total_throughput_pct"] / cnt, 2) if cnt else 0.0,
            "total_quantity":     round(d["total_quantity"], 2),
            "month": month, "year": year, "is_pan_india": False,
        })
        all_pct_sum += d["total_throughput_pct"]
        all_qty_sum += d["total_quantity"]
        all_count   += cnt

    location_rows.sort(key=lambda r: r["avg_throughput_pct"])

    if all_count:
        location_rows.append({
            "exco_location":      "PAN India",
            "plant_count":        all_count,
            "avg_throughput_pct": round(all_pct_sum / all_count, 2),
            "total_quantity":     round(all_qty_sum, 2),
            "month": month, "year": year, "is_pan_india": True,
        })

    return location_rows


# ── Persist results ───────────────────────────────────────────────────────────

def save_tp_results(plant_rows: list, month: int, year: int) -> int:
    """Replace stored results for this month/year with new plant_rows."""
    conn = database.get_connection()
    try:
        conn.execute(
            "DELETE FROM tp_results WHERE month = ? AND year = ?", (month, year)
        )
        conn.commit()
    finally:
        conn.close()

    # Only persist columns that exist in the tp_results table (total_time_min is
    # a display-only field kept in the in-memory rows, not stored).
    keep = {"month", "year", "lookup_code", "plant_name", "exco_location",
            "business_head", "plant_manager", "mixer_theo_cap", "total_quantity",
            "total_time_hrs", "throughput_pct", "batch_count", "generated_at"}
    db_rows = [{k: v for k, v in r.items() if k in keep} for r in plant_rows]
    return database.insert_rows("tp_results", db_rows)
