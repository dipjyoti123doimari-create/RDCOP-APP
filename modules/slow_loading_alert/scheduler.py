"""
scheduler.py
============
SLA background job functions — registered in app.py's _start_scheduler().

Job IDs (follow project convention):
  sla_hourly_alert         — runs every hour
  sla_daily_summary        — runs every day at 07:00 AM

Each function is self-contained and safe to call manually for testing.
"""

from datetime import datetime as _dt, date as _date, timedelta
import database


def run_hourly_alert_job(preview_only: bool = False) -> dict:
    """
    1. Fetch today's loading data from Oracle.
    2. Apply mappings (plant name, batcher name, emails).
    3. Detect slow loading cases.
    4. Deduplicate within the current hour.
    5. Send per-plant hourly alert emails.
    6. Persist alert records and logs.

    Returns a summary dict: {total_checked, total_alerts, total_sent, errors}
    """
    from modules.slow_loading_alert import oracle_service, mapping_service, alert_service, email_service

    started_at = _dt.now().isoformat(timespec="seconds")
    today_str  = str(_date.today())
    now_hour   = _dt.now().hour
    summary    = {"total_checked": 0, "total_alerts": 0, "total_sent": 0, "errors": []}

    # Guard: hourly alert enabled?
    if database.get_module_setting("sla", "hourly_alert_enabled", "true") != "true":
        database.sla_log_scheduler("sla_hourly_alert", started_at,
                                    _dt.now().isoformat(timespec="seconds"),
                                    "SKIPPED", error_message="Hourly alert disabled.")
        return summary

    try:
        # 1. Fetch Oracle data
        import oracle_connector as _oc
        if not _oc.is_configured():
            raise RuntimeError("Oracle not configured.")
        df, warnings = oracle_service.fetch_loading_data(today_str, today_str)
        for w in warnings:
            print(f"[sla-hourly] {w}")
        summary["total_checked"] = len(df)

        if df.empty:
            database.sla_log_scheduler("sla_hourly_alert", started_at,
                                        _dt.now().isoformat(timespec="seconds"),
                                        "SUCCESS", total_checked=0)
            return summary

        # 2. Enrich with mappings
        plant_map   = mapping_service.get_plant_mapping()
        batcher_map = mapping_service.get_batcher_mapping()
        df = mapping_service.apply_mappings(df, plant_map, batcher_map)

        # 3. Detect
        records = alert_service.detect_slow_loading(df, today_str, now_hour)
        summary["total_alerts"] = len(records)

        if not records:
            database.sla_log_scheduler("sla_hourly_alert", started_at,
                                        _dt.now().isoformat(timespec="seconds"),
                                        "SUCCESS", total_checked=summary["total_checked"])
            return summary

        # 4. Deduplicate
        new_records, skipped = alert_service.deduplicate_records(records)

        if preview_only:
            return {"preview": new_records, **summary}

        # 5. Persist alert records
        def _strip_private(r):
            return {k: v for k, v in r.items() if not k.startswith("_") and k != "severity"}

        if new_records:
            inserted_rows = [_strip_private(r) for r in new_records]
            database.sla_bulk_insert_records(inserted_rows)

        # Re-fetch inserted IDs by alert_key
        sent_ids = []
        failed_ids = []

        # 6. Group by plant and send
        from collections import defaultdict
        by_plant = defaultdict(list)
        for r in new_records:
            by_plant[r["plant_code"]].append(r)

        for plant_code, plant_recs in by_plant.items():
            plant_name = plant_recs[0].get("plant_name", plant_code)
            success = email_service.send_hourly_alert(
                plant_code, plant_name, plant_recs, today_str, now_hour
            )
            # Fetch DB IDs for these records (by key + hour + date)
            conn = database.get_connection()
            try:
                for pr in plant_recs:
                    cur = conn.execute(
                        "SELECT id FROM slow_loading_alert_records WHERE alert_key=? AND alert_date=? AND alert_hour=? ORDER BY id DESC LIMIT 1",
                        (pr["alert_key"], today_str, now_hour)
                    )
                    row = cur.fetchone()
                    if row:
                        (sent_ids if success else failed_ids).append(row[0])
            finally:
                conn.close()

            if success:
                summary["total_sent"] += 1

        database.sla_mark_records_sent(sent_ids)
        database.sla_mark_records_failed(failed_ids, "Email send failed")

        completed_at = _dt.now().isoformat(timespec="seconds")
        database.sla_log_scheduler(
            "sla_hourly_alert", started_at, completed_at, "SUCCESS",
            total_checked=summary["total_checked"],
            total_alerts=summary["total_alerts"],
            total_sent=summary["total_sent"],
        )

    except Exception as exc:
        import traceback
        err = f"{exc}\n{traceback.format_exc()}"
        summary["errors"].append(str(exc))
        database.sla_log_scheduler(
            "sla_hourly_alert", started_at,
            _dt.now().isoformat(timespec="seconds"),
            "FAILED", error_message=str(exc)[:1000]
        )
        print(f"[sla-hourly] ERROR: {err}")

    return summary


def run_daily_summary_job(summary_date: str = None, preview_only: bool = False) -> dict:
    """
    Send previous-day (or specified date) slow loading summary.

    BH receives summary for all their plants.
    PM + BM receive per-plant summary.
    """
    from modules.slow_loading_alert import mapping_service, email_service
    from collections import defaultdict

    started_at  = _dt.now().isoformat(timespec="seconds")
    summary     = {"total_checked": 0, "total_sent": 0, "errors": []}

    if database.get_module_setting("sla", "daily_summary_enabled", "true") != "true":
        database.sla_log_scheduler("sla_daily_summary", started_at,
                                    _dt.now().isoformat(timespec="seconds"),
                                    "SKIPPED", error_message="Daily summary disabled.")
        return summary

    if summary_date is None:
        summary_date = str(_date.today() - timedelta(days=1))

    try:
        # Load previous-day alert records from DB
        records = database.sla_get_report(
            from_date=summary_date, to_date=summary_date,
            status=None, limit=5000
        )
        summary["total_checked"] = len(records)

        if not records:
            database.sla_log_scheduler("sla_daily_summary", started_at,
                                        _dt.now().isoformat(timespec="seconds"),
                                        "SUCCESS", total_checked=0)
            return summary

        plant_map = mapping_service.get_plant_mapping()
        bh_plant_map = mapping_service.get_bh_plant_map(plant_map)  # {bh_email: [codes]}

        # Group records by plant_name for display
        by_plant_code = defaultdict(list)
        for r in records:
            by_plant_code[r["plant_code"]].append(r)

        if preview_only:
            summary["preview"] = records
            return summary

        # Send BH emails (each BH gets all their plants in one email)
        for bh_email, codes in bh_plant_map.items():
            bh_plant_records = {}
            for code in codes:
                recs = by_plant_code.get(code, [])
                if recs:
                    pname = recs[0].get("plant_name", code)
                    bh_plant_records[pname] = recs
            if not bh_plant_records:
                continue

            # Collect all CC emails across these plants
            cc_set = set()
            for code in codes:
                info = plant_map.get(code, {})
                for e in (info.get("cc_emails", "") or "").split(","):
                    e = e.strip()
                    if e:
                        cc_set.add(e)

            success = email_service.send_daily_summary(
                bh_email=bh_email, pm_emails=[], bm_emails=[],
                cc_email=",".join(cc_set),
                plant_records=bh_plant_records,
                summary_date=summary_date,
            )
            if success:
                summary["total_sent"] += 1

        # Send per-plant PM + BM emails
        for plant_code, recs in by_plant_code.items():
            info = plant_map.get(plant_code, {})
            pm_email = info.get("pm_email", "")
            bm_email = info.get("bm_email", "")
            bh_email = info.get("bh_email", "")
            cc_email = info.get("cc_emails", "")
            pname = recs[0].get("plant_name", plant_code)

            pm_bm_to = list(filter(None, [pm_email, bm_email]))
            if not pm_bm_to:
                continue
            success = email_service.send_daily_summary(
                bh_email="",
                pm_emails=[pm_email] if pm_email else [],
                bm_emails=[bm_email] if bm_email else [],
                cc_email=cc_email,
                plant_records={pname: recs},
                summary_date=summary_date,
            )
            if success:
                summary["total_sent"] += 1

        # Mark daily summary records
        ids = [r["id"] for r in records if r.get("id")]
        conn = database.get_connection()
        try:
            now_ts = _dt.now().isoformat(timespec="seconds")
            conn.executemany(
                "UPDATE slow_loading_alert_records SET alert_type='DAILY_SUMMARY', updated_at=? WHERE id=?",
                [(now_ts, i) for i in ids]
            )
            conn.commit()
        finally:
            conn.close()

        completed_at = _dt.now().isoformat(timespec="seconds")
        database.sla_log_scheduler(
            "sla_daily_summary", started_at, completed_at, "SUCCESS",
            total_checked=summary["total_checked"],
            total_sent=summary["total_sent"],
        )

    except Exception as exc:
        import traceback
        summary["errors"].append(str(exc))
        database.sla_log_scheduler(
            "sla_daily_summary", started_at,
            _dt.now().isoformat(timespec="seconds"),
            "FAILED", error_message=str(exc)[:1000]
        )
        print(f"[sla-daily] ERROR: {exc}\n{traceback.format_exc()}")

    return summary
