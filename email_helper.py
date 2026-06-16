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
  get_smtp_config()              -> dict of current SMTP settings
  is_configured()                -> True when host + sender + password are set
  wrap_html_body_with_image(...) -> HTML body string with inline CID image
  create_tp_preview_image(...)   -> generate TP report PNG, return path or None
  create_report_preview_image(.) -> generate I&D report PNG, return path or None
  send_report_email(...)         -> {"success": bool, "error": str | None}

Email design: PNG preview image embedded inline (CID) so colors and borders
display correctly in Gmail/Outlook regardless of CSS stripping. Full Excel
report is attached separately.
"""

import os
import smtplib
from email.message import EmailMessage

import database


_DEFAULT_PORT  = 587
_XLSX_MAINTYPE = "application"
_XLSX_SUBTYPE  = "vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def get_smtp_config() -> dict:
    """Read all email settings. Password falls back to the SMTP_PASSWORD env var."""
    s = database.get_all_settings()
    password = os.environ.get("SMTP_PASSWORD") or s.get("smtp_password", "")
    raw_port = s.get("smtp_port", str(_DEFAULT_PORT))
    try:
        port = int(raw_port)
    except (TypeError, ValueError):
        port = _DEFAULT_PORT

    return {
        "host":       (s.get("smtp_host",   "") or "").strip(),
        "port":       port,
        "sender":     (s.get("smtp_sender", "") or "").strip(),
        "password":   password,
        "use_tls":    s.get("smtp_use_tls", "true") == "true",
        "default_to": s.get("email_default_to", "") or "",
        "default_cc": s.get("email_default_cc", "") or "",
        "subject":    s.get("email_subject", "Batching Incentive & Deduction Report"),
    }


def is_configured() -> bool:
    """True when the essentials (host, sender, password) are all set."""
    c = get_smtp_config()
    return bool(c["host"] and c["sender"] and c["password"])


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
    return ("Report of Production, Penalty & Incentive summary for "
            f"(PI/QCI, MO, Teamlease Employee & All Trainees) - {month_label}")


def compose_report_body(month_label, sections=None) -> str:
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
    if not text:
        return []
    parts = [p.strip() for p in str(text).replace(";", ",").split(",")]
    return [p for p in parts if p]


def wrap_html_body(message_text, tables_html="") -> str:
    """Legacy plain HTML body (used as fallback when image generation fails)."""
    import html as _html
    safe = _html.escape(message_text or "").replace("\n", "<br>")
    return (
        '<div style="font-family:Arial,Helvetica,sans-serif;font-size:13px;'
        'color:#0A2540;line-height:1.45">'
        f'<div>{safe}</div>'
        f'{tables_html or ""}'
        '</div>'
    )


def wrap_html_body_with_image(message_text, excel_attached=True) -> str:
    """
    HTML email body that references an inline PNG via CID (Content-ID).
    The caller must attach the image separately with Content-ID <report_preview>.
    """
    import html as _html
    safe = _html.escape(message_text or "").replace("\n", "<br>")
    note = (
        '<p style="color:#555;font-size:12px;margin-top:14px">'
        '&#128206; Full detailed report is attached as an Excel file.</p>'
        if excel_attached else ""
    )
    return (
        '<!DOCTYPE html><html><body>'
        '<div style="font-family:Arial,Helvetica,sans-serif;font-size:13px;'
        'color:#0A2540;line-height:1.6;max-width:960px;margin:0 auto;padding:20px">'
        f'<div style="margin-bottom:18px">{safe}</div>'
        '<div style="margin:18px 0">'
        '<img src="cid:report_preview" alt="Report Preview" '
        'style="max-width:100%;height:auto;border:1px solid #ddd;border-radius:6px">'
        '</div>'
        f'{note}'
        '<p style="color:#333;margin-top:18px">Regards</p>'
        '</div></body></html>'
    )


# ─────────────────────────────────────────────────────────────────────────────
# PNG PREVIEW IMAGE GENERATION (Pillow)
# Generates a high-DPI table image so colors and gridlines survive any email
# client regardless of CSS handling.
# ─────────────────────────────────────────────────────────────────────────────

_S   = 2     # 2× scale for Retina / high-DPI quality
_RH  = 26    # base row height in virtual px (multiplied by _S)
_THH = 36    # section title row height
_PAD = 6     # cell inner padding

_C = {
    "bg":      (255, 255, 255),
    "hdr_bg":  ( 10,  37,  64),   # #0A2540 — matches app thead
    "hdr_txt": (255, 255, 255),
    "bdr":     (154, 154, 154),   # #9A9A9A — reference-mail border
    "red":     (255, 179, 179),   # #FFB3B3
    "yellow":  (255, 224, 102),   # #FFE066
    "green":   (146, 212, 146),   # #92D492
    "pan":     (210, 210, 210),   # PAN India separator row
    "text":    ( 20,  20,  20),
    "muted":   (110, 110, 110),
}


def _pil_font(pt, bold=False):
    from PIL import ImageFont
    face = "arialbd.ttf" if bold else "arial.ttf"
    for path in [
        f"C:/Windows/Fonts/{face}",
        f"/usr/share/fonts/truetype/liberation/Liberation{'Bold' if bold else 'Regular'}.ttf",
    ]:
        try:
            return ImageFont.truetype(path, pt * _S)
        except OSError:
            pass
    return ImageFont.load_default()


def _pil_tw(draw, text, font):
    try:
        return int(draw.textlength(str(text), font=font))
    except AttributeError:
        return draw.textsize(str(text), font=font)[0]


def _pil_clip(text, max_px, draw, font):
    s = str(text)
    if _pil_tw(draw, s, font) <= max_px:
        return s
    while s and _pil_tw(draw, s + "…", font) > max_px:
        s = s[:-1]
    return (s + "…") if s else ""


def _img_fmt(val):
    try:
        f = float(val)
        return f"{int(f):,}" if f == int(f) else f"{f:,.1f}"
    except (TypeError, ValueError):
        return str(val) if val is not None else ""


def _draw_img_table(draw, title, col_defs, data_rows, color_fn, y_start,
                    f_title, f_hdr, f_data):
    """
    Draw one titled table section onto the ImageDraw canvas.
    col_defs   = [(header_text, scaled_px_width), ...]
    data_rows  = list of lists of str values
    color_fn(row_index, row_list) -> RGB tuple
    Returns new y position after drawing.
    """
    pad     = _PAD * _S
    rh      = _RH  * _S
    th      = _THH * _S
    bdr     = _C["bdr"]
    total_w = sum(w for _, w in col_defs)
    y       = y_start

    # Section title row
    draw.rectangle([0, y, total_w, y + th], fill=_C["hdr_bg"])
    draw.text((pad, y + (th - 11 * _S) // 2), str(title),
              fill=_C["hdr_txt"], font=f_title)
    y += th

    # Column header row
    x = 0
    for label, w in col_defs:
        draw.rectangle([x, y, x + w, y + rh], fill=_C["hdr_bg"])
        tw = _pil_tw(draw, label, f_hdr)
        tx = x + max(pad, (w - tw) // 2)
        draw.text((tx, y + pad // 2), label, fill=_C["hdr_txt"], font=f_hdr)
        x += w
    draw.line([0, y + rh, total_w, y + rh], fill=bdr, width=_S)
    y += rh

    # Data rows
    for ri, row in enumerate(data_rows):
        bg = color_fn(ri, row)
        x = 0
        for ci, (_, w) in enumerate(col_defs):
            draw.rectangle([x, y, x + w, y + rh], fill=bg)
            val = _pil_clip(row[ci] if ci < len(row) else "", w - pad * 2, draw, f_data)
            draw.text((x + pad, y + pad // 2), val, fill=_C["text"], font=f_data)
            draw.line([x + w, y, x + w, y + rh], fill=bdr, width=1)
            x += w
        draw.line([0, y + rh, total_w, y + rh], fill=bdr, width=1)
        y += rh

    # Outer border box (top + left + right + bottom)
    t0 = y_start
    draw.line([0,       t0, total_w, t0], fill=bdr, width=_S)
    draw.line([0,       t0, 0,       y],  fill=bdr, width=_S)
    draw.line([total_w, t0, total_w, y],  fill=bdr, width=_S)
    draw.line([0,       y,  total_w, y],  fill=bdr, width=_S)

    return y + _S * 16   # small gap after the table


def create_tp_preview_image(plant_rows, location_rows, month, year,
                            output_path, max_rows=30):
    """
    Generate TP Plant Throughput Report preview PNG.
    Shows the location table + plant table (limited to max_rows plants).
    Returns output_path on success, None on failure.
    """
    try:
        import calendar
        from PIL import Image, ImageDraw

        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

        mon_name = calendar.month_name[month]
        mon_tag  = f"{calendar.month_abbr[month]}'{str(year)[-2:]}"

        LOC_COLS = [
            ("#",                 28),
            ("Exco Location",    140),
            ("Plants",            50),
            ("Total Qty",         85),
            ("Total Time (min)", 115),
            ("Avg TP %",          65),
        ]
        PLT_COLS = [
            ("#",                 28),
            ("Plant",            190),
            ("Exco Location",     95),
            ("Business Head",    115),
            ("Total Qty",         85),
            ("Time (min)",        80),
            ("TP %",              55),
        ]
        loc_defs = [(h, w * _S) for h, w in LOC_COLS]
        plt_defs = [(h, w * _S) for h, w in PLT_COLS]
        loc_w    = sum(w for _, w in loc_defs)
        plt_w    = sum(w for _, w in plt_defs)
        img_w    = max(loc_w, plt_w)

        # Build location rows
        loc_rows_out = []
        for i, r in enumerate(location_rows, 1):
            pan = r.get("is_pan_india", False)
            pct = round(float(r.get("avg_throughput_pct", 0) or 0))
            loc_rows_out.append([
                "—" if pan else str(i),
                ("■ " if pan else "") + str(r.get("exco_location", "")),
                str(r.get("plant_count", "")),
                _img_fmt(r.get("total_quantity", 0)),
                _img_fmt(r.get("total_time_min", 0)),
                f"{pct}%",
            ])

        # Build plant rows (limited)
        plt_limited  = plant_rows[:max_rows]
        plt_rows_out = []
        for i, r in enumerate(plt_limited, 1):
            pct = round(float(r.get("throughput_pct", 0) or 0))
            plt_rows_out.append([
                str(i),
                str(r.get("plant_name", "")),
                str(r.get("exco_location", "")),
                str(r.get("business_head", "")),
                _img_fmt(r.get("total_quantity", 0)),
                _img_fmt(r.get("total_time_min", 0)),
                f"{pct}%",
            ])

        banner_h = 52 * _S
        gap      = 16 * _S
        rh       = _RH  * _S
        th       = _THH * _S
        footer_h = 28 * _S if len(plant_rows) > max_rows else 0

        total_h = (banner_h + gap
                   + th + rh + rh * len(loc_rows_out) + _S * 16
                   + th + rh + rh * len(plt_rows_out) + _S * 16
                   + footer_h + 20 * _S)

        img  = Image.new("RGB", (img_w, total_h), _C["bg"])
        draw = ImageDraw.Draw(img)

        f_ttl  = _pil_font(13, bold=True)
        f_hdr  = _pil_font(9,  bold=True)
        f_data = _pil_font(9,  bold=False)
        f_foot = _pil_font(8,  bold=False)

        # Main banner
        draw.rectangle([0, 0, img_w - 1, banner_h - 1], fill=_C["hdr_bg"])
        draw.text((_PAD * _S, (banner_h - 13 * _S) // 2),
                  f"RDC-TP Plant Throughput Report — {mon_name} {year}",
                  fill=_C["hdr_txt"], font=f_ttl)
        y = banner_h + gap

        def _loc_color(ri, row):
            name    = row[1] if len(row) > 1 else ""
            pct_str = row[-1].replace("%", "").strip()
            if "■" in name or "PAN" in name.upper():
                return _C["pan"]
            try:
                pct = float(pct_str)
            except ValueError:
                return _C["bg"]
            if pct < 60: return _C["red"]
            if pct < 75: return _C["yellow"]
            return _C["green"]

        def _plt_color(ri, row):
            pct_str = row[-1].replace("%", "").strip()
            try:
                pct = float(pct_str)
            except ValueError:
                return _C["bg"]
            if pct < 60: return _C["red"]
            if pct < 75: return _C["yellow"]
            return _C["green"]

        y = _draw_img_table(draw, f"Location wise Throughput — {mon_tag}",
                            loc_defs, loc_rows_out, _loc_color, y,
                            f_ttl, f_hdr, f_data)
        y = _draw_img_table(draw, f"Plant Throughput report — {mon_tag}",
                            plt_defs, plt_rows_out, _plt_color, y,
                            f_ttl, f_hdr, f_data)

        if len(plant_rows) > max_rows:
            draw.text((_PAD * _S, y + 4 * _S),
                      (f"Showing first {max_rows} of {len(plant_rows)} plants. "
                       "Full report in attached Excel."),
                      fill=_C["muted"], font=f_foot)

        img.save(output_path, "PNG", dpi=(144, 144))
        return output_path

    except Exception as exc:
        print(f"[email_helper] TP preview image error: {exc}")
        return None


def create_report_preview_image(df, output_path, max_rows=30, month_label=""):
    """
    Generate I&D Incentive & Deduction Report preview PNG.
    Shows category-wise sections from the results DataFrame.
    df uses internal column names (employee_name, category, ...).
    Returns output_path on success, None on failure.
    """
    try:
        from PIL import Image, ImageDraw
        from report_generator import EMAIL_SECTIONS, _prepare_category_df

        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

        # Preview columns: (internal/renamed key, display header, base px width)
        COLS = [
            ("#",               "#",              28),
            ("Employee Name",   "Employee Name",  148),
            ("Category",        "Category",       120),
            ("Plant",           "Plant",           68),
            ("Total Quantity",  "Total Qty",       80),
            ("Incentive Amount","Incentive Amt",   98),
            ("Deduction Amount","Deduction Amt",   98),
            ("Remarks",         "Remarks",          65),
        ]
        col_keys = [k for k, _, _ in COLS]
        col_defs = [(h, w * _S) for _, h, w in COLS]
        img_w    = sum(w for _, w in col_defs)

        ded_idx = col_keys.index("Deduction Amount")
        inc_idx = col_keys.index("Incentive Amount")

        def _id_color(ri, row):
            try:
                ded = float(str(row[ded_idx]).replace(",", "")) if ded_idx < len(row) else 0
            except ValueError:
                ded = 0
            try:
                inc = float(str(row[inc_idx]).replace(",", "")) if inc_idx < len(row) else 0
            except ValueError:
                inc = 0
            if ded > 0: return _C["red"]
            if inc > 0: return _C["green"]
            return _C["bg"]

        # Gather section data (total across all sections <= max_rows)
        sections_out = []
        total_shown  = 0
        for title, cats in EMAIL_SECTIONS:
            if total_shown >= max_rows:
                break
            sec_df = _prepare_category_df(df, cats)
            if sec_df.empty:
                continue
            sliced = sec_df.head(max_rows - total_shown)
            total_shown += len(sliced)

            rows_out = []
            for rn, (_, row) in enumerate(sliced.iterrows(), 1):
                cells = []
                for k in col_keys:
                    if k == "#":
                        cells.append(str(rn))
                    elif k in ("Total Quantity", "Incentive Amount", "Deduction Amount"):
                        val = row.get(k, "") if hasattr(row, "get") else getattr(row, k, "")
                        cells.append(_img_fmt(val))
                    else:
                        val = row.get(k, "") if hasattr(row, "get") else getattr(row, k, "")
                        cells.append(str(val) if val is not None else "")
                rows_out.append(cells)
            sections_out.append((title, rows_out))

        # Calculate image height
        banner_h = 52 * _S
        gap      = 16 * _S
        rh       = _RH  * _S
        th       = _THH * _S
        footer_h = 28 * _S if total_shown >= max_rows else 0

        total_h = banner_h + gap
        for _, rows in sections_out:
            total_h += th + rh + rh * len(rows) + _S * 16
        total_h += footer_h + 20 * _S
        if not sections_out:
            total_h += 60 * _S  # empty state

        img  = Image.new("RGB", (img_w, total_h), _C["bg"])
        draw = ImageDraw.Draw(img)

        f_ttl  = _pil_font(13, bold=True)
        f_hdr  = _pil_font(9,  bold=True)
        f_data = _pil_font(9,  bold=False)
        f_foot = _pil_font(8,  bold=False)

        # Main banner
        draw.rectangle([0, 0, img_w - 1, banner_h - 1], fill=_C["hdr_bg"])
        draw.text((_PAD * _S, (banner_h - 13 * _S) // 2),
                  ("Batching Incentive & Deduction Report"
                   + (f" — {month_label}" if month_label else "")),
                  fill=_C["hdr_txt"], font=f_ttl)
        y = banner_h + gap

        for title, rows in sections_out:
            y = _draw_img_table(draw, title, col_defs, rows, _id_color, y,
                                f_ttl, f_hdr, f_data)

        if total_shown >= max_rows:
            draw.text((_PAD * _S, y + 4 * _S),
                      (f"Showing first {max_rows} rows only. "
                       "Full report is in the attached Excel file."),
                      fill=_C["muted"], font=f_foot)

        img.save(output_path, "PNG", dpi=(144, 144))
        return output_path

    except Exception as exc:
        print(f"[email_helper] I&D preview image error: {exc}")
        return None


# ─────────────────────────────────────────────────────────────────────────────
# EMAIL SENDING
# ─────────────────────────────────────────────────────────────────────────────

def send_report_email(to_emails, cc_emails, subject, body,
                      attachment_bytes=None, attachment_name=None,
                      html_body=None, inline_image_path=None) -> dict:
    """
    Send the report email with optional inline PNG preview and Excel attachment.

    When inline_image_path points to an existing PNG file the message uses a
    multipart/related structure so the image is embedded inline (CID reference).
    html_body must include <img src="cid:report_preview"> to display it.

    Falls back to a plain EmailMessage (no inline image) when the image file
    is missing or image generation failed.

    Returns {"success": bool, "error": str | None}.
    Every attempt is recorded in the email_log table.
    """
    cfg     = get_smtp_config()
    to_list = _split_emails(to_emails)
    cc_list = _split_emails(cc_emails)
    result  = {"success": False, "error": None}

    if not is_configured():
        result["error"] = "SMTP is not configured. Add settings on the Settings page."
    elif not to_list:
        result["error"] = "Please enter at least one 'To' email address."
    else:
        try:
            use_image = bool(inline_image_path and os.path.isfile(inline_image_path))

            if use_image:
                msg = _build_mime_with_inline_image(
                    cfg, to_list, cc_list, subject, body,
                    html_body, inline_image_path,
                    attachment_bytes, attachment_name)
            else:
                # Fallback: standard EmailMessage (no inline image)
                msg = EmailMessage()
                msg["From"]    = cfg["sender"]
                msg["To"]      = ", ".join(to_list)
                if cc_list:
                    msg["Cc"]  = ", ".join(cc_list)
                msg["Subject"] = subject or cfg["subject"]
                msg.set_content(body or "Please find the attached report.")
                if html_body:
                    msg.add_alternative(html_body, subtype="html")
                if attachment_bytes:
                    msg.add_attachment(
                        attachment_bytes,
                        maintype=_XLSX_MAINTYPE, subtype=_XLSX_SUBTYPE,
                        filename=attachment_name or "report.xlsx")

            recipients = to_list + cc_list
            with smtplib.SMTP(cfg["host"], cfg["port"], timeout=30) as server:
                if cfg["use_tls"]:
                    server.starttls()
                server.login(cfg["sender"], cfg["password"])
                server.send_message(msg, to_addrs=recipients)

            result["success"] = True
        except Exception as exc:
            result["error"] = str(exc)

    database.log_email(
        report_file_name=attachment_name or "",
        to_emails=", ".join(to_list),
        cc_emails=", ".join(cc_list),
        subject=subject or cfg["subject"],
        status="Success" if result["success"] else "Failed",
        error_message=result["error"],
    )
    return result


def _build_mime_with_inline_image(cfg, to_list, cc_list, subject, body,
                                  html_body, image_path,
                                  attachment_bytes, attachment_name):
    """
    Build multipart/mixed message:
      mixed
      ├── related
      │   ├── HTML (with <img src="cid:report_preview">)
      │   └── PNG  (Content-ID: <report_preview>, inline)
      └── Excel attachment (if provided)
    """
    from email.mime.multipart import MIMEMultipart
    from email.mime.text      import MIMEText
    from email.mime.image     import MIMEImage
    from email.mime.base      import MIMEBase
    from email                import encoders

    outer = MIMEMultipart("mixed")
    outer["From"]    = cfg["sender"]
    outer["To"]      = ", ".join(to_list)
    if cc_list:
        outer["Cc"] = ", ".join(cc_list)
    outer["Subject"] = subject or cfg["subject"]

    # related = HTML body + inline image
    related = MIMEMultipart("related")
    html = html_body or f"<p>{body or 'Please find the report preview below.'}</p>"
    related.attach(MIMEText(html, "html", "utf-8"))

    with open(image_path, "rb") as fh:
        img_data = fh.read()
    img_part = MIMEImage(img_data, _subtype="png")
    img_part.add_header("Content-ID", "<report_preview>")
    img_part.add_header("Content-Disposition", "inline",
                        filename="report_preview.png")
    related.attach(img_part)
    outer.attach(related)

    # Excel attachment
    if attachment_bytes:
        xlsx = MIMEBase(_XLSX_MAINTYPE, _XLSX_SUBTYPE)
        xlsx.set_payload(attachment_bytes)
        encoders.encode_base64(xlsx)
        xlsx.add_header("Content-Disposition", "attachment",
                        filename=attachment_name or "report.xlsx")
        outer.attach(xlsx)

    return outer
