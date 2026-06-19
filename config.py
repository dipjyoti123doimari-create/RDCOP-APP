"""
config.py
=========
This file holds all the "constants" for the app.

A constant is a value that does not change while the app runs, for example:
- the names of the columns we expect in the uploaded files,
- the list of allowed employee categories,
- the incentive rates and deduction targets,
- the names of the pages and Excel sheets,
- the theme colors.

Keeping these in one place means that if a rule or a column name ever
changes, you only edit it HERE and the whole app updates automatically.
You do NOT have to hunt through every file.

NOTE: In Phase 1 we only USE the app name, page list and theme colors.
The other values (rates, targets, columns) are written down now so they
are ready for the later phases (validation, calculation, export).
"""

# ---------------------------------------------------------------------------
# 1. BASIC APP INFORMATION
# ---------------------------------------------------------------------------

APP_NAME = "Batching Incentive & Deduction Calculator"
APP_TAGLINE = "Month-wise batching incentives and deductions across all plants"

# Path (relative to this project folder) where an optional logo can live.
# If this file does not exist, the app simply skips the logo (it will not crash).
LOGO_PATH = "assets/app_logo.png"


# ---------------------------------------------------------------------------
# 2. SIDEBAR PAGES
# ---------------------------------------------------------------------------
# These are the pages shown in the left sidebar menu, in order.
# In Phase 1 each page is just a placeholder. We will fill them in later phases.

PAGES = [
    "Dashboard",
    "Data Uploader",
    "Calculate Incentive & Deduction",
    "View Reports",
    "Error / Validation Report",
    "Settings",
]


# ---------------------------------------------------------------------------
# 3. EXPECTED COLUMN NAMES (used later for validation & calculation)
# ---------------------------------------------------------------------------
# We list the columns each data source MUST contain. Validation (Phase 5)
# will check the uploaded files against these lists.

MASTER_DATA_COLUMNS = [
    "Employee Code",
    "Employee Name",
    "Designation",
    "Category",
    "Plant",
    "Plant Code",
]

# Backend Data may have many columns; these are the ones we actually need.
BACKEND_REQUIRED_COLUMNS = [
    "Created by",   # employee code of the batcher
    "Quantity",     # quantity produced in that vehicle
    "Date",         # used for month-wise / date-range reports
]

MAINTENANCE_REQUIRED_COLUMNS = [
    "Plant Code",
    "YTD Maintenance Cost",   # already in Rs per cum
]


# ---------------------------------------------------------------------------
# 4. EMPLOYEE CATEGORIES
# ---------------------------------------------------------------------------
# These are the only valid values allowed in the Master Data "Category" column.

CATEGORIES = [
    "Civil Trainee",
    "Non-Civil Trainee",
    "PM & API",
    "QCI",
    "MO",
    "SPE",
    "Production Officer",
    "TL BPO",
    "NA",
]


# ---------------------------------------------------------------------------
# 5. INCENTIVE RULES
# ---------------------------------------------------------------------------
# Every eligible category (except Production Officer) needs Total Quantity >= 1000
# to qualify for an incentive.

INCENTIVE_MIN_QUANTITY = 1000          # minimum quantity to be eligible
MAINTENANCE_COST_THRESHOLD = 18        # <= 18 is "low cost", > 18 is "high cost"

# Incentive rate (Rs per quantity) for SPE category.
SPE_RATE_LOW_COST = 5.0     # when YTD maintenance cost <= 18
SPE_RATE_HIGH_COST = 2.5    # when YTD maintenance cost  > 18

# Incentive rate (Rs per quantity) for all other eligible categories
# (everyone except SPE and Production Officer).
OTHER_RATE_LOW_COST = 3.0   # when YTD maintenance cost <= 18
OTHER_RATE_HIGH_COST = 1.5  # when YTD maintenance cost  > 18

# Production Officer never gets an incentive.
NO_INCENTIVE_CATEGORIES = ["Production Officer"]


# ---------------------------------------------------------------------------
# 6. DEDUCTION RULES
# ---------------------------------------------------------------------------
# For each category we store:
#   "target" -> the minimum quantity expected
#   "rate"   -> Rs deducted for every unit of shortfall below the target
#
# Deduction Amount = (target - total_quantity) * rate, but only when
# total_quantity is BELOW the target. Otherwise the deduction is 0.
#
# "NA" has no deduction at all (target 0).

DEDUCTION_RULES = {
    "Civil Trainee":      {"target": 300, "rate": 10},
    "Non-Civil Trainee":  {"target": 500, "rate": 10},
    "PM & API":           {"target": 50,  "rate": 20},
    "QCI":                {"target": 50,  "rate": 20},
    "MO":                 {"target": 500, "rate": 20},
    "SPE":                {"target": 500, "rate": 20},
    "Production Officer":  {"target": 500, "rate": 20},
    "TL BPO":             {"target": 500, "rate": 20},
    "NA":                 {"target": 0,   "rate": 0},
}


# ---------------------------------------------------------------------------
# 7. EXCEL REPORT SHEET NAMES (used later in Phase 10)
# ---------------------------------------------------------------------------
# Each report sheet name is mapped to the categories that belong in it.

REPORT_SHEETS = {
    "Summary": [],  # special sheet, filled with totals
    "All Trainees Incentive & Deduction": ["Civil Trainee", "Non-Civil Trainee"],
    "Plant Manager and PI Incentive and Deduction": ["PM & API"],
    "QCI Incentive & Deduction": ["QCI"],
    "All MO Incentive & Deduction": ["MO"],
    "SPE Incentive & Deduction": ["SPE"],
    "TL Employee Incentive & Deduction": ["TL BPO"],
    "Production Officer Deduction": ["Production Officer"],
    "NA Incentive": ["NA"],
    "Unmapped Employees": [],   # special sheet
    "Validation Errors": [],    # special sheet
}

# The exact columns every report sheet should contain (in this order).
REPORT_OUTPUT_COLUMNS = [
    "Employee Code",
    "Employee Name",
    "Designation",
    "Category",
    "Plant",
    "Batching Quantity",
    "YTD Maintenance Cost",
    "Incentive Eligible",
    "Incentive Rate",
    "Incentive Amount",
    "Deduction Target",
    "Shortfall Quantity",
    "Deduction Amount",
    "Remarks",
]


# ---------------------------------------------------------------------------
# 8. THEME COLORS (used by ui_helpers.py for the modern look)
# ---------------------------------------------------------------------------
# These follow the "Modern light SaaS dashboard" direction.

THEME = {
    "bg":            "#0A2540",                          # dark navy — Stripe's main bg
    "card_bg":       "rgba(255, 255, 255, 0.04)",        # frosted glass card
    "card_border":   "rgba(255, 255, 255, 0.08)",        # subtle card border
    "text_primary":  "#FFFFFF",                          # white text
    "text_secondary":"rgba(255, 255, 255, 0.70)",            # soft white — readable over any blob
    "accent_purple": "#635BFF",                          # Stripe purple
    "accent_blue":   "#00D4FF",                          # Stripe cyan
    "success":       "rgba(16, 185, 129, 0.12)",         # dark green tint
    "error":         "rgba(239, 68, 68, 0.12)",          # dark red tint
    "warning":       "rgba(245, 158, 11, 0.12)",         # dark amber tint
}
