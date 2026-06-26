"""
email_service.py
================
Builds and sends SLA alert emails using the existing SMTP configuration
(email_helper.send_report_email with html_body=).

Two email types:
  1. HOURLY  — plant-wise compact HTML table, sent to PM + BM + CC
  2. DAILY_SUMMARY — previous-day summary grouped by plant, sent to BH + PM + BM + CC
"""

import smtplib
from datetime import datetime as _dt
from email.message import EmailMessage

import database
import email_helper


# ── Severity colours ──────────────────────────────────────────────────────────
_SEV_COLOR = {
    "RED":   "#FF4444",
    "AMBER": "#FFAA00",
    "":      "#333333",
}
_SEV_BG = {
    "RED":   "rgba(255,68,68,0.12)",
    "AMBER": "rgba(255,170,0,0.12)",
    "":      "transparent",
}


def _row_style(severity: str) -> str:
    bg = _SEV_BG.get(severity, "transparent")
    return f'background:{bg}'


def _delay_cell(delay: float, severity: str) -> str:
    color = _SEV_COLOR.get(severity, "#333")
    return f'<td style="text-align:right;color:{color};font-weight:600;padding:3px 5px;border:1px solid #999">{delay}</td>'


# ── Table builders ─────────────────────────────────────────────────────────────

def _html_table(records: list) -> str:
    headers = [
        "Plant", "Customer", "Grade", "Batcher",
        "TM Number", "Batched Qty", "Load Time", "Allowed Time", "Delay (min)"
    ]
    th = "".join(
        f'<th style="background:#082B49;color:#fff;padding:4px 6px;'
        f'border:1px solid #999;text-align:{"right" if h in ("Batched Qty","Load Time","Allowed Time","Delay (min)") else "left"}'
        f'">{h}</th>'
        for h in headers
    )
    rows_html = ""
    for r in records:
        sev = r.get("severity", "")
        td = f'<tr style="{_row_style(sev)}">'
        def _c(v, align="left"):
            return (f'<td style="padding:3px 5px;border:1px solid #999;'
                    f'text-align:{align};white-space:nowrap">{v}</td>')
        td += _c(r.get("plant_name", r.get("plant_code", "")))
        td += _c(r.get("customer", ""))
        td += _c(r.get("grade", ""))
        td += _c(r.get("batcher_name", r.get("batcher_code", "")))
        td += _c(r.get("tm_number", ""))
        td += _c(f'{float(r.get("batched_quantity",0)):.1f}', "right")
        td += _c(f'{float(r.get("loading_time_minutes",0)):.0f} min', "right")
        td += _c(f'{float(r.get("allowed_loading_minutes",0)):.0f} min', "right")
        td += _delay_cell(f'{float(r.get("delay_minutes",0)):.0f}', sev)
        td += "</tr>"
        rows_html += td

    return (
        '<table style="border-collapse:collapse;font-size:11px;width:100%;'
        'font-family:Arial,Helvetica,sans-serif">'
        f'<thead><tr>{th}</tr></thead>'
        f'<tbody>{rows_html}</tbody>'
        '</table>'
    )


def _base_html(greeting: str, intro: str, summary_html: str, table_html: str) -> str:
    return f"""<!DOCTYPE html><html><body>
<div style="font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#222;
     max-width:900px;margin:0 auto;padding:16px">
  <p style="margin:0 0 8px">{greeting}</p>
  <p style="margin:0 0 12px">{intro}</p>
  {summary_html}
  <div style="margin-top:12px">{table_html}</div>
  <p style="margin-top:16px;color:#333">Regards,<br><strong>RDC Operations</strong></p>
</div>
</body></html>"""


def _summary_box(items: list) -> str:
    rows = "".join(
        f'<tr><td style="padding:3px 8px;color:#555">{k}</td>'
        f'<td style="padding:3px 8px;font-weight:600">{v}</td></tr>'
        for k, v in items
    )
    return (
        '<table style="border-collapse:collapse;font-size:12px;'
        'background:#f7f9fc;border:1px solid #dde;border-radius:6px;margin-bottom:8px">'
        f'{rows}</table>'
    )


# ── Hourly email ──────────────────────────────────────────────────────────────

def build_hourly_html(plant_name: str, alert_date: str, alert_hour: int,
                       records: list) -> str:
    total = len(records)
    max_delay = max((r.get("delay_minutes", 0) for r in records), default=0)
    hour_label = f"{alert_hour:02d}:00"

    summary = _summary_box([
        ("Plant", plant_name),
        ("Alert Hour", hour_label),
        ("Slow Loading Cases", total),
        ("Maximum Delay", f"{max_delay:.0f} min"),
    ])
    table = _html_table(records)
    return _base_html(
        greeting="Dear Sir/Ma'am,",
        intro=(f"This is an automated <strong>Slow Loading Alert</strong> for "
               f"<strong>{plant_name}</strong> at <strong>{hour_label}</strong> "
               f"on <strong>{alert_date}</strong>. "
               f"<strong>{total}</strong> vehicle(s) exceeded the allowed loading time."),
        summary_html=summary,
        table_html=table,
    )


def send_hourly_alert(plant_code: str, plant_name: str, records: list,
                       alert_date: str, alert_hour: int) -> bool:
    """Send hourly alert for one plant. Returns True on success."""
    if not records:
        return False

    pm_email = records[0].get("_pm_email", "")
    bm_email = records[0].get("_bm_email", "")
    cc_email = records[0].get("_cc_emails", "")

    to_emails = ",".join(filter(None, [pm_email, bm_email]))
    if not to_emails:
        database.sla_log_email(
            "HOURLY", alert_date, alert_hour, plant_code, plant_name,
            "", cc_email,
            f"Slow Loading Alert | {plant_name} | {alert_date} {alert_hour:02d}:00",
            len(records), "FAILED",
            "No Plant Manager or Business Manager email configured."
        )
        return False

    subject = f"Slow Loading Alert | {plant_name} | {alert_date} {alert_hour:02d}:00"
    html_body = build_hourly_html(plant_name, alert_date, alert_hour, records)

    try:
        _send_html(to_emails, cc_email, subject, html_body)
        database.sla_log_email(
            "HOURLY", alert_date, alert_hour, plant_code, plant_name,
            to_emails, cc_email, subject, len(records), "SENT"
        )
        return True
    except Exception as exc:
        database.sla_log_email(
            "HOURLY", alert_date, alert_hour, plant_code, plant_name,
            to_emails, cc_email, subject, len(records), "FAILED", str(exc)
        )
        return False


# ── Daily summary email ───────────────────────────────────────────────────────

def build_daily_html(summary_date: str, plant_records: dict) -> str:
    """
    plant_records: {plant_name: [record, ...]}
    """
    total = sum(len(v) for v in plant_records.values())
    plant_counts = [(pn, len(rs)) for pn, rs in plant_records.items()]
    plant_summary_rows = "".join(
        f'<tr><td style="padding:2px 8px">{pn}</td>'
        f'<td style="padding:2px 8px;text-align:right;font-weight:600">{cnt}</td></tr>'
        for pn, cnt in plant_counts
    )
    plant_summary = (
        '<table style="border-collapse:collapse;font-size:11px;margin-bottom:10px">'
        '<tr><th style="background:#082B49;color:#fff;padding:3px 8px">Plant</th>'
        '<th style="background:#082B49;color:#fff;padding:3px 8px">Cases</th></tr>'
        f'{plant_summary_rows}</table>'
    )

    all_tables = ""
    for pn, recs in plant_records.items():
        all_tables += (
            f'<p style="margin:14px 0 4px;font-weight:600;font-size:13px">{pn}</p>'
            + _html_table(recs)
        )

    summary = _summary_box([
        ("Summary Date", summary_date),
        ("Total Slow Loading Cases", total),
    ])
    return _base_html(
        greeting="Dear Sir/Ma'am,",
        intro=(f"Please find below the <strong>Daily Slow Loading Summary</strong> "
               f"for <strong>{summary_date}</strong>."),
        summary_html=summary + plant_summary,
        table_html=all_tables,
    )


def send_daily_summary(bh_email: str, pm_emails: list, bm_emails: list,
                        cc_email: str, plant_records: dict,
                        summary_date: str) -> bool:
    """Send the daily summary to BH, all PMs and BMs for the plants in plant_records."""
    if not plant_records:
        return False

    plant_names = list(plant_records.keys())
    to_list = list(filter(None, [bh_email] + pm_emails + bm_emails))
    if not to_list:
        return False

    to_emails = ",".join(to_list)
    region_label = plant_names[0] if len(plant_names) == 1 else f"{len(plant_names)} Plants"
    subject = f"Daily Slow Loading Summary | {region_label} | {summary_date}"
    html_body = build_daily_html(summary_date, plant_records)

    total_cases = sum(len(v) for v in plant_records.values())
    try:
        _send_html(to_emails, cc_email, subject, html_body)
        database.sla_log_email(
            "DAILY_SUMMARY", summary_date, 0, "", region_label,
            to_emails, cc_email, subject, total_cases, "SENT"
        )
        return True
    except Exception as exc:
        database.sla_log_email(
            "DAILY_SUMMARY", summary_date, 0, "", region_label,
            to_emails, cc_email, subject, total_cases, "FAILED", str(exc)
        )
        return False


# ── Low-level sender ─────────────────────────────────────────────────────────

def _send_html(to_emails: str, cc_emails: str, subject: str, html_body: str):
    """Send HTML email using the existing global SMTP config."""
    cfg = email_helper.get_smtp_config()
    if not (cfg["host"] and cfg["sender"] and cfg["password"]):
        raise RuntimeError("SMTP is not configured.")

    to_list = email_helper._split_emails(to_emails)
    cc_list = email_helper._split_emails(cc_emails)
    if not to_list:
        raise RuntimeError("No recipient.")

    msg = EmailMessage()
    msg["From"]    = cfg["sender"]
    msg["To"]      = ", ".join(to_list)
    if cc_list:
        msg["Cc"]  = ", ".join(cc_list)
    msg["Subject"] = subject
    msg.set_content("Please view this email in an HTML-capable client.")
    msg.add_alternative(html_body, subtype="html")

    with smtplib.SMTP(cfg["host"], cfg["port"], timeout=30) as server:
        if cfg["use_tls"]:
            server.starttls()
        server.login(cfg["sender"], cfg["password"])
        server.send_message(msg, to_addrs=to_list + cc_list)
