# Batching Incentive & Deduction Calculator

A **local** web app (Python + Streamlit + SQLite) that calculates month-wise
batching **incentives** and **deductions** for all employees across all plants.

> **Status: All 14 phases complete — app is fully built and production-ready.**

---

## 1. What you need installed

Only **Python 3.9 or newer**. Check with:

```bash
python --version
```

If not installed, download from https://www.python.org/downloads/ and tick
**"Add Python to PATH"** during setup.

---

## 2. First-time setup (do once)

Open a terminal in the project folder (the folder that contains `app.py`).

**Step 1 — create a virtual environment (recommended)**

```bash
python -m venv .venv
```

**Step 2 — activate it**

- Windows PowerShell:
  ```powershell
  .venv\Scripts\Activate.ps1
  ```
  If blocked, run once: `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`

- macOS / Linux:
  ```bash
  source .venv/bin/activate
  ```

You'll see `(.venv)` at the start of the terminal line when active.

**Step 3 — install packages**

```bash
pip install -r requirements.txt
```

---

## 3. How to run the app

```bash
streamlit run app.py
```

The app opens at **http://localhost:8501** (or port 2001 if configured).
Stop with `Ctrl + C`.

---

## 4. What the app does

| Page | Description |
|------|-------------|
| **Dashboard** | Live KPI cards (master data rows, backend rows, maintenance rows, calculated rows) and database status |
| **Data Uploader** | Three tabs: sync Master Data from Google Sheets, upload Backend Data Excel, upload Maintenance Cost Excel |
| **Calculate Incentive & Deduction** | Pick a date range, run the full category-wise calculation, browse results by category with colour-coded rows |
| **View Reports** | Browse any date range with multi-select filters (Category, Designation, Plant), generate a report snapshot, download Excel / CSV, or email the report |
| **Error / Validation Report** | Run data quality checks across all three data sources; review and clear validation errors |
| **Settings** | Animated background controls, SMTP email settings, cache & database compaction |

### Colour coding
- 🟢 **Green row** — incentive earned
- 🔴 **Red row** — deduction applied (takes priority if both apply)
- In reports: red rows first (largest first), then green, then plain

---

## 5. Email setup

Go to **Settings → Email (SMTP) settings**:

| Field | Value |
|-------|-------|
| SMTP host | `smtp.gmail.com` |
| Port | `587` |
| Sender email | your Gmail address |
| Password | Gmail **App Password** (not your login password) |
| Use TLS | ON |

To create a Gmail App Password: Google Account → Security → 2-Step Verification
→ App Passwords → create one for "Mail".

---

## 6. Google Sheets setup (Master Data)

Go to **Data Uploader → Master Data Sync & Management**.

**Public mode (easiest):** Share your sheet as "Anyone with the link can view",
paste the URL, click Sync Now.

**Private mode:** Add `credentials/service_account.json` (Google Cloud service
account). See `credentials/README.txt` for setup steps.

Your sheet must have these columns (spelling matters, case does not):
`Employee Code · Employee Name · Designation · Category · Plant · Plant Code`

---

## 7. Project structure

```
Incentive_Calculator/
│
├── app.py                # Main app: page routing and all page functions
├── config.py             # All constants: categories, rates, column names, colours
├── ui_helpers.py         # Theme/CSS, headers, cards, messages, progress tracker
│
├── database.py           # SQLite database — all read/write/table helpers
├── google_sheets.py      # Google Sheets sync (public URL or service account)
├── data_loader.py        # Read and clean uploaded Excel files
├── validations.py        # Data quality checks across all three data sources
├── calculator.py         # Category-wise incentive & deduction calculation engine
├── report_generator.py   # Excel workbook builder (multi-sheet, colour-coded)
├── email_helper.py       # SMTP email sender with HTML/plain-text + attachment
├── cache_helpers.py      # Streamlit cache clear + SQLite VACUUM compaction
│
├── requirements.txt      # Python package list
├── README.md             # This file
│
├── data/                 # SQLite database file (app.db) — not committed to git
├── exports/              # Generated Excel reports — not committed to git
├── credentials/          # Google service_account.json (if used) — not in git
├── assets/               # Optional app_logo.png
└── utils/
    └── animated_background.py   # Stripe-inspired animated ray-burst background
```

---

## 8. Animated background

A Stripe-inspired fan of thin rays that reacts to your mouse.

**6 themes:** Pre-dawn · Sunrise · Daytime · Dusk · Sunset · Night

**Automatic mode (default):** theme follows your system clock:
`04:00–05:59 Pre-dawn · 06:00–08:59 Sunrise · 09:00–16:59 Daytime ·
17:00–18:29 Dusk · 18:30–20:00 Sunset · 20:01–03:59 Night`

**Manual mode:** Settings page → turn off Auto → pick any theme.
A quick-select dropdown also lives in the top-right corner of the page.

**Performance:** FPS capped at 40, DPR capped at 2, gradient cached.
Settings: **Animation On/Off** and **Intensity** (Low 70 / Medium 120 / High 180 rays).

---

## 9. Build history (all phases complete)

| Phase | Feature |
|------:|---------|
| 1 | Project skeleton, theme, logo, sidebar navigation |
| 2 | SQLite database and all 9 tables |
| 3 | Google Sheets Master Data sync (public + private modes) |
| 4 | Upload Backend Data and Maintenance Cost Excel files |
| 5 | Data validation with friendly error messages |
| 6 | Incentive & deduction calculation rules (all categories) |
| 7 | Date-range filtering |
| 8 | Dynamic multi-select filters from Master Data |
| 9 | View Reports page with report snapshot (Generate button) |
| 10 | Excel export (multi-sheet, colour-coded, header frozen) |
| 11 | Master Data add / edit / delete + change log |
| 12 | Email report via SMTP + email log |
| 13 | Cache clear + SQLite VACUUM compaction |
| 14 | Final cleanup and documentation |

---

Made to be simple and local-first. No Django, FastAPI, React, Docker, or cloud
deployment needed — just Python and a browser.
