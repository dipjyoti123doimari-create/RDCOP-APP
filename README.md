# Batching Incentive & Deduction Calculator

A simple, **local** web app (built with Python + Streamlit) that calculates
month-wise batching **incentives** and **deductions** based on the quantity each
employee produced and any shortfall, across all plants.

> **Status: Phase 1 complete** — this is the project skeleton: folder
> structure, a modern theme, logo support, sidebar navigation, and placeholder
> pages. The real features (data upload, calculation, reports, email) are added
> in the later phases listed at the bottom of this file.

---

## 1. What you need installed

You only need **Python** (version 3.9 or newer). To check if you already have it,
open the VS Code terminal and run:

```bash
python --version
```

If you see something like `Python 3.11.x`, you are ready. If not, install Python
from https://www.python.org/downloads/ and tick **"Add Python to PATH"** during
installation.

---

## 2. How to set up the project (one time)

Do these steps once, inside VS Code, in the terminal
(**Terminal → New Terminal**). Make sure the terminal is in the project folder
(the folder that contains `app.py`).

**Step 1 — (Recommended) create a "virtual environment".**
A virtual environment is a private box for this project's packages, so they do
not mix with other Python projects.

```bash
python -m venv .venv
```

**Step 2 — activate the virtual environment.**

- On **Windows PowerShell** (the default VS Code terminal on Windows):
  ```powershell
  .venv\Scripts\Activate.ps1
  ```
  If PowerShell blocks the script, run this once and try again:
  ```powershell
  Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
  ```

- On **macOS / Linux**:
  ```bash
  source .venv/bin/activate
  ```

When it is active you will see `(.venv)` at the start of the terminal line.

**Step 3 — install the required packages.**

```bash
pip install -r requirements.txt
```

This downloads Streamlit, pandas, and the other tools listed in
`requirements.txt`.

---

## 3. How to RUN the app

Still in the project folder (with the virtual environment active), run:

```bash
streamlit run app.py
```

Streamlit will start and automatically open the app in your web browser
(usually at http://localhost:8501). If it does not open by itself, copy that
address from the terminal into your browser.

To **stop** the app, click the terminal and press `Ctrl + C`.

---

## 4. What you should see (Phase 1 test)

When the app opens you should see:

- A **left sidebar** with the app name (or your logo) and a navigation menu.
- A colourful **gradient header** at the top of each page.
- The **Dashboard** page with four KPI cards (all showing 0 for now).
- Eight other pages in the sidebar. Clicking each one shows a "Coming soon"
  card — that is expected in Phase 1.

If all of that works, **Phase 1 is successful.** 🎉

---

## 5. Optional: add a logo

Add a PNG image named `app_logo.png` into the `assets/` folder. The app will
then show it in the sidebar and use it as the browser tab icon. If you skip
this, the app still works and uses a default icon.

---

## 6. Project structure

```
Incentive_Calculator/
│
├── app.py                # Main app: sidebar + navigation (keep it small)
├── config.py             # All constants: columns, categories, rates, colors
├── ui_helpers.py         # The theme/look: CSS, headers, cards, messages
│
├── database.py           # SQLite database (Phase 2)         [placeholder]
├── google_sheets.py      # Google Sheets master-data sync (Phase 3) [placeholder]
├── data_loader.py        # Read/clean uploaded Excel (Phase 4) [placeholder]
├── validations.py        # Validate data (Phase 5)            [done]
├── calculator.py         # Incentive/deduction rules (Phase 6-7) [done]
├── report_generator.py   # Excel export & formatting (Phase 10) [done]
├── email_helper.py       # Email the report (Phase 12)        [done]
├── cache_helpers.py      # Performance/cache helpers (Phase 13)[done]
│
├── requirements.txt      # The list of Python packages to install
├── README.md             # This file
│
├── data/                 # The SQLite database file (app.db) will live here
├── exports/              # Generated Excel reports will be saved here
├── credentials/          # Google service_account.json goes here (Phase 3)
└── assets/               # Optional app_logo.png goes here
```

The `[placeholder]` files are intentionally empty for now. Each one has a
comment at the top explaining exactly what it will do and in which phase.

---

## 7. The build plan (phases)

| Phase | What it adds |
|------:|--------------|
| **1** | **Project skeleton, theme, logo, navigation (done ✅)** |
| 2  | SQLite database and tables |
| 3  | Google Sheets Master Data sync |
| 4  | Upload Backend Data and Maintenance Cost Excel files |
| 5  | Data validation with friendly error messages |
| 6  | Incentive & deduction calculation rules |
| 7  | Date-range filtering |
| 8  | Dynamic multi-select filters from Master Data |
| 9  | View Reports page |
| 10 | Excel export (multiple sheets, colours, formatting, progress bar) |
| 11 | Master Data add / edit / delete + change log |
| 12 | Email circulation of the report + email log |
| 13 | Caching / performance + "Clear Cache" button |
| 14 | Final cleanup and documentation |

A Stripe-inspired **animated background** theme system is also planned as an
add-on after the core app is working.

---

## 8. Settings that come later (just so you know)

- **Google Sheets (Phase 3):** see `credentials/README.txt`. You will create a
  Google Cloud service account, enable the Google Sheets API, download a JSON
  key into `credentials/service_account.json`, and share your sheet with the
  service account's email.
- **Email / SMTP (Phase 12):** the Settings page will let you enter SMTP host,
  port, sender email, app password, and TLS. **Passwords are never hardcoded
  and never printed** — they are kept in Streamlit secrets or environment
  variables and masked in the UI.

---

## 9. Animated background (Stripe-inspired)

The app has a reusable animated background: a fan of thin rays bursting up from
the bottom of the screen that gently reacts to your mouse. It lives in
`utils/animated_background.py` and is applied app-wide from `app.py`.

**Themes (6):** Pre-dawn, Sunrise, Daytime, Dusk, Sunset, Night.

**How the theme is chosen** — controlled on the **Settings** page:
- **Automatic (default):** the theme follows your computer's clock, detected in
  the browser:
  `04:00-05:59 Pre-dawn · 06:00-08:59 Sunrise · 09:00-16:59 Daytime ·
   17:00-18:29 Dusk · 18:30-20:00 Sunset · 20:01-03:59 Night`.
- **Manual:** turn automatic off and pick a theme from the dropdown.
- **Theme dropdown (LOV):** a dropdown in the **top-right corner** lets you pick
  "Auto" or any of the six themes directly, at any time.

**Disable / performance:**
- Settings has **Background animation On/Off** (off = calm static gradient) and
  **Intensity Low / Medium / High** (70 / 120 / 180 rays).
- If your system has "reduce motion" enabled, the app automatically shows a
  static gradient instead of the animation.
- The background sits *behind* everything and never blocks buttons, forms,
  uploads, filters or tables. It uses no external libraries or CDNs.

> Tip: the dark themes (Night, Pre-dawn) darken the page. If any text is hard to
> read, pick a lighter theme or switch the animation off.

---

Made to be simple and local first. No Django, FastAPI, React, Docker, or cloud
deployment is used.
