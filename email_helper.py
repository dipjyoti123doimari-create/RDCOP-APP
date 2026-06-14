"""
email_helper.py
===============
Phase 12 — send the generated Excel report by email.

How the SMTP settings are found (in priority order):
  - The password comes from the environment variable SMTP_PASSWORD if set,
    otherwise from the value saved on the Settings page (stored locally in the
    app database). The password is NEVER hardcoded and NEVER printed/logged.
  - Host, port, sender, TLS and default recipients come from the Settings page
    (the app_settings table).

Public functions:
  get_smtp_config()   -> dict of the current settings (password included only
                         so we can actually log in; we never display it)
  is_configured()     -> True when host + sender + password are all present
  send_report_email(to, cc, subject, body, attachment_bytes, attachment_name)
                      -> {"success": bool, "error": str | None}
                         Every attempt is recorded in the email_log table.

All email logic lives HERE so the rest of the app stays clean.
"""

import os
import smtplib
from email.message import EmailMessage

import database


_DEFAULT_PORT = 587
# MIME type for a .xlsx file
_XLSX_MAINTYPE = "application"
_XLSX_SUBTYPE = "vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def get_smtp_config() -> dict:
    """Read all email settings. Password falls back to the SMTP_PASSWORD env var.

    All keys are read in ONE database call (get_all_settings) instead of one
    connection per key — this keeps the Settings page snappy.
    """
    s = database.get_all_settings()
    password = os.environ.get("SMTP_PASSWORD") or s.get("smtp_password", "")
    raw_port = s.get("smtp_port", str(_DEFAULT_PORT))
    try:
        port = int(raw_port)
    except (TypeError, ValueError):
        port = _DEFAULT_PORT

    return {
        "host":       (s.get("smtp_host", "") or "").strip(),
        "port":       port,
        "sender":     (s.get("smtp_sender", "") or "").strip(),
        "password":   password,
        "use_tls":    s.get("smtp_use_tls", "true") == "true",
        "default_to": s.get("email_default_to", "") or "",
        "default_cc": s.get("email_default_cc", "") or "",
        "subject":    s.get(
            "email_subject", "Batching Incentive & Deduction Report"),
    }


def is_configured() -> bool:
    """True when the essentials (host, sender, password) are all set."""
    c = get_smtp_config()
    return bool(c["host"] and c["sender"] and c["password"])


# The seven report sections, in the exact order and wording used in the
# company's standard monthly mail (see the Mail Reference). Keeping them here
# means the email body always matches the agreed format.
REPORT_SECTIONS = [
    "Production report of all Trainees",
    "Production report of PM/API",
    "Production Report of QC Person",
    "Production report of MO",
    "Production report of SPE",
    "Production report of Officer Production (Onroll Batcher)",
    "Production report of TL Batcher, Mechanic",
]


def compose_report_subject(month_label) -> str:
    """The standard subject line, with the report month filled in."""
    return ("Report of Production, Penalty & Incentive summary for "
            f"(PI/QCI, MO, Teamlease Employee & All Trainees) - {month_label}")


def compose_report_body(month_label, sections=None) -> str:
    """
    Build the default email body in the company's standard format:

        Dear Sir,

        Please find below the compiled report for the month of <month>,
        covering the following:

        1. Production report of all Trainees
        2. Production report of PM/API
        ... (the seven standard sections) ...

        Kindly review the attached Excel file below and in case of any
        clarifications or corrections, please mail me and Kanhaiya sir

    `month_label` is a friendly month string (e.g. "May 2026"). The user can
    still edit the text before sending.
    """
    sections = sections or REPORT_SECTIONS
    numbered = "\n".join(f"{i}. {s}" for i, s in enumerate(sections, start=1))
    return (
        "Dear Sir,\n\n"
        f"Please find below the compiled report for the month of {month_label}, "
        "covering the following:\n\n"
        f"{numbered}\n\n"
        "Kindly review the attached Excel file below and in case of any "
        "clarifications or corrections, please mail me and Kanhaiya sir"
    )


def _split_emails(text) -> list:
    """Turn 'a@x.com, b@y.com; c@z.com' into a clean list of addresses."""
    if not text:
        return []
    parts = [p.strip() for p in str(text).replace(";", ",").split(",")]
    return [p for p in parts if p]


def wrap_html_body(message_text, tables_html="") -> str:
    """
    Turn the plain-text message into an HTML email body and append the report
    tables (if any) AFTER the text — matching the company format where the
    section tables follow the closing line.
    """
    import html as _html
    safe = _html.escape(message_text or "").replace("\n", "<br>")
    return (
        '<div style="font-family:Arial,Helvetica,sans-serif;font-size:13px;'
        'color:#0A2540;line-height:1.45">'
        f'<div>{safe}</div>'
        f'{tables_html or ""}'
        '</div>'
    )


def send_report_email(to_emails, cc_emails, subject, body,
                      attachment_bytes=None, attachment_name=None,
                      html_body=None) -> dict:
    """
    Send an email (optionally with the Excel report attached) and log the
    attempt. Returns {"success": bool, "error": str | None}.

    If `html_body` is given, the email is sent as HTML (with `body` kept as the
    plain-text fallback for clients that don't render HTML). This is how the
    report tables get shown inside the email body.
    """
    cfg = get_smtp_config()
    to_list = _split_emails(to_emails)
    cc_list = _split_emails(cc_emails)
    result = {"success": False, "error": None}

    # ---- Validate before trying to connect ----
    if not is_configured():
        result["error"] = "SMTP is not configured. Add settings on the Settings page."
    elif not to_list:
        result["error"] = "Please enter at least one 'To' email address."
    else:
        try:
            msg = EmailMessage()
            msg["From"] = cfg["sender"]
            msg["To"] = ", ".join(to_list)
            if cc_list:
                msg["Cc"] = ", ".join(cc_list)
            msg["Subject"] = subject or cfg["subject"]
            msg.set_content(body or "Please find the attached report.")
            if html_body:
                # Adds an HTML alternative; email clients show this over the text.
                msg.add_alternative(html_body, subtype="html")

            if attachment_bytes:
                msg.add_attachment(
                    attachment_bytes,
                    maintype=_XLSX_MAINTYPE, subtype=_XLSX_SUBTYPE,
                    filename=attachment_name or "report.xlsx",
                )

            recipients = to_list + cc_list
            with smtplib.SMTP(cfg["host"], cfg["port"], timeout=30) as server:
                if cfg["use_tls"]:
                    server.starttls()
                server.login(cfg["sender"], cfg["password"])
                server.send_message(msg, to_addrs=recipients)

            result["success"] = True
        except Exception as exc:  # noqa: BLE001 - report any send failure
            result["error"] = str(exc)

    # ---- Log every attempt (never store the password) ----
    database.log_email(
        report_file_name=attachment_name or "",
        to_emails=", ".join(to_list),
        cc_emails=", ".join(cc_list),
        subject=subject or cfg["subject"],
        status="Success" if result["success"] else "Failed",
        error_message=result["error"],
    )
    return result
