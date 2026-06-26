"""
alert_service.py
================
Core slow-loading detection logic.

Given a DataFrame of loading rows (already enriched with mappings),
detects slow loading events, builds alert records, deduplicates within
an hour, and returns lists ready for DB insert.
"""

import hashlib
from datetime import datetime as _dt

from modules.slow_loading_alert.threshold_service import get_allowed_time, classify_severity


def build_alert_key(plant_code: str, tm_number: str, grade: str, customer: str,
                     batched_quantity: float, loading_time: float,
                     alert_date: str, alert_hour: int) -> str:
    """Build a deterministic deduplication key for one alert event."""
    raw = (
        f"{plant_code}|{tm_number}|{grade}|{customer}|"
        f"{round(batched_quantity,2)}|{round(loading_time,2)}|"
        f"{alert_date}|{alert_hour}"
    )
    return hashlib.md5(raw.encode()).hexdigest()


def detect_slow_loading(df, alert_date: str, alert_hour: int) -> list:
    """
    Evaluate every row in df and return a list of alert record dicts
    for rows where loading_time_minutes > allowed_loading_minutes.

    df must have columns:
        plant_code, plant_name, customer, grade, batcher_code, batcher_name,
        tm_number, batched_quantity, loading_time_minutes, mixer_capacity,
        pm_email, bm_email, bh_email, cc_emails
    """
    records = []
    now_ts = _dt.now().isoformat(timespec="seconds")

    for _, row in df.iterrows():
        plant_code    = str(row.get("plant_code", "")).strip()
        plant_name    = str(row.get("plant_name", plant_code)).strip()
        customer      = str(row.get("customer", "")).strip()
        grade         = str(row.get("grade", "")).strip()
        batcher_code  = str(row.get("batcher_code", "")).strip()
        batcher_name  = str(row.get("batcher_name", batcher_code)).strip()
        tm_number     = str(row.get("tm_number", "")).strip()
        batched_qty   = float(row.get("batched_quantity", 0) or 0)
        load_time     = float(row.get("loading_time_minutes", 0) or 0)
        mixer_cap     = row.get("mixer_capacity")

        # Validation — skip blank/zero rows
        if not plant_code or not tm_number or batched_qty <= 0 or load_time <= 0:
            continue

        allowed_time, threshold_remark = get_allowed_time(mixer_cap, batched_qty)

        if load_time <= allowed_time:
            continue  # not slow — no alert

        delay = load_time - allowed_time
        severity = classify_severity(delay)

        key = build_alert_key(plant_code, tm_number, grade, customer,
                               batched_qty, load_time, alert_date, alert_hour)

        records.append({
            "alert_date":              alert_date,
            "alert_hour":              alert_hour,
            "plant_code":              plant_code,
            "plant_name":              plant_name,
            "customer":                customer,
            "grade":                   grade,
            "batcher_code":            batcher_code,
            "batcher_name":            batcher_name,
            "tm_number":               tm_number,
            "batched_quantity":        batched_qty,
            "mixer_capacity":          float(mixer_cap) if mixer_cap is not None else None,
            "loading_time_minutes":    load_time,
            "allowed_loading_minutes": allowed_time,
            "delay_minutes":           round(delay, 2),
            "alert_type":              "HOURLY",
            "status":                  "OPEN",
            "remarks":                 threshold_remark,
            "alert_key":               key,
            "severity":                severity,
            # email addresses carried for convenience (not stored in DB)
            "_pm_email":               str(row.get("pm_email", "")).strip(),
            "_bm_email":               str(row.get("bm_email", "")).strip(),
            "_bh_email":               str(row.get("bh_email", "")).strip(),
            "_cc_emails":              str(row.get("cc_emails", "")).strip(),
        })

    return records


def deduplicate_records(records: list) -> tuple[list, list]:
    """
    Split records into new (to insert + send) vs skipped (duplicate in same hour).
    Checks the DB for already-sent keys.
    """
    import database as _db
    new_records = []
    skipped = []
    for r in records:
        if _db.sla_alert_key_exists_this_hour(r["alert_key"], r["alert_date"], r["alert_hour"]):
            r["status"] = "SKIPPED_DUPLICATE"
            skipped.append(r)
        else:
            new_records.append(r)
    return new_records, skipped
