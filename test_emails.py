"""
Automated test: verify both TP and I&D email HTML builders.
Checks: color bands, header style, border style, PAN India row, empty sections.
"""
import sys, re
sys.path.insert(0, 'd:/Dipjyoti Doimari/AI/Incentive_Calculator')
from app import _tp_build_html_tables
from report_generator import build_email_tables_html
import pandas as pd

PASS = "[PASS]"
FAIL = "[FAIL]"
results = []

def check(label, condition):
    tag = PASS if condition else FAIL
    results.append(f"{tag} {label}")
    if not condition:
        print(f"  !! FAILED: {label}")

# ── TP test data ──────────────────────────────────────────────────────────────
plant_rows = [
    {'plant_name':'NCR-Gurgaon_BP2','exco_location':'NCR','business_head':'Japjit Singh',
     'plant_manager':'Vishal S','mixer_theo_cap':60,'total_quantity':466,
     'total_time_min':2638,'throughput_pct':18,'batch_count':10},  # RED
    {'plant_name':'GOA-Goa 2','exco_location':'Goa','business_head':'Adnan Khan',
     'plant_manager':'Satyam Jha','mixer_theo_cap':56,'total_quantity':982,
     'total_time_min':4958,'throughput_pct':67,'batch_count':18},  # YELLOW
    {'plant_name':'BLR-Whitefield','exco_location':'Bangalore','business_head':'Rahul',
     'plant_manager':'Sharma','mixer_theo_cap':50,'total_quantity':4200,
     'total_time_min':2100,'throughput_pct':88,'batch_count':80},  # GREEN
]
location_rows = [
    {'exco_location':'NCR','plant_count':2,'total_quantity':583,'total_time_min':3169,
     'avg_throughput_pct':19,'is_pan_india':False},    # RED
    {'exco_location':'Goa','plant_count':1,'total_quantity':983,'total_time_min':4958,
     'avg_throughput_pct':67,'is_pan_india':False},    # YELLOW
    {'exco_location':'Bangalore','plant_count':17,'total_quantity':28598,'total_time_min':18000,
     'avg_throughput_pct':76,'is_pan_india':False},    # GREEN
    {'exco_location':'PAN India','plant_count':20,'total_quantity':30164,'total_time_min':26127,
     'avg_throughput_pct':47,'is_pan_india':True},
]

tp_html = _tp_build_html_tables(plant_rows, location_rows, 6, 2026)

# Color checks
check("TP red row color #FFD5D5",    "#FFD5D5" in tp_html)
check("TP yellow row color #FFF3CC", "#FFF3CC" in tp_html)
check("TP green row color #D5F0D5",  "#D5F0D5" in tp_html)
check("TP no old rgba red",          "rgba(239,68,68,0.38)" not in tp_html)
check("TP no old rgba green",        "rgba(16,185,129,0.38)" not in tp_html)

# Header/border checks
check("TP dark navy header #0A2540", "#0A2540" in tp_html)
check("TP reference-mail border #9A9A9A", "9A9A9A" in tp_html)
check("TP white header text color:#fff", "color:#fff" in tp_html)

# Title row
check("TP location table title present", "Location wise Throughput - Jun'26" in tp_html)
check("TP plant table title present",    "Plant Throughput report - Jun'26" in tp_html)

# PAN India checks
check("TP PAN India flag emoji",     "PAN India" in tp_html)
check("TP PAN India #EEEEEE row",    "#EEEEEE" in tp_html)
check("TP PAN India top border",     "border-top:2px solid #555" in tp_html)
check("TP PAN India dash in # col",  ">—<" in tp_html)

# Column checks
check("TP Total Time (min) col",     "Total Time (min)" in tp_html)
check("TP no Plant Code col",        "Plant Code" not in tp_html)
check("TP numbering present",        ">1<" in tp_html and ">2<" in tp_html)

print("\n-- TP Results --")
for r in results: print(r)
results.clear()

# ── I&D test data ─────────────────────────────────────────────────────────────
data = {
    'month': [6,6,6,6],
    'year':  [2026]*4,
    'employee_code': ['E001','E002','E003','E004'],
    'employee_name': ['Alice','Bob','Carol','Dave'],
    'designation':   ['Trainee','Trainee','PM & API','QCI'],
    'category':      ['Civil Trainee','Civil Trainee','PM & API','QCI'],
    'plant':         ['Plant A','Plant B','Plant A','Plant C'],
    'plant_code':    ['PA','PB','PA','PC'],
    'total_quantity':  [120, 80, 200, 90],
    'ytd_maintenance_cost': [0,0,0,0],
    'incentive_eligible':   [1, 0, 1, 0],
    'incentive_rate':       [10, 0, 15, 0],
    'incentive_amount':     [1200, 0, 3000, 0],
    'deduction_target':     [0, 100, 0, 50],
    'shortfall_quantity':   [0,  20, 0, 10],
    'deduction_amount':     [0, 400, 0, 200],
    'remarks':              ['OK','Short','OK','Short'],
}
df = pd.DataFrame(data)
id_html = build_email_tables_html(df)

# Color checks
check("I&D red row #FFD5D5 (deduction)",  "#FFD5D5" in id_html)
check("I&D green row #D5F0D5 (incentive)","#D5F0D5" in id_html)
check("I&D no old rgba red",              "rgba(239,68,68" not in id_html)
check("I&D no old rgba green",            "rgba(16,185,129" not in id_html)
check("I&D white rows #ffffff",           "#ffffff" in id_html)

# Header / border
check("I&D dark navy header #0A2540",     "#0A2540" in id_html)
check("I&D reference-mail border 9A9A9A", "9A9A9A" in id_html)
check("I&D no light gray EEEEEE header",  "background:#EEEEEE" not in id_html)

# Section titles
check("I&D section 1 Trainees title",     "Production report of all Trainees" in id_html)
check("I&D section 2 PM/API title",       "Production report of PM/API" in id_html)
check("I&D section 7 TL Batcher title",   "Production report of TL Batcher" in id_html)
check("I&D empty sections handled",       "No records for this section" in id_html)

# Deduction label
check("I&D Deduction Amount @ Rs label",  "Deduction Amount @ Rs" in id_html)

# No <style> block (all inline)
check("I&D no CSS class .rpt",            'class="rpt"' not in id_html)
check("I&D no <style> block",             '<style>' not in id_html)

print("\n-- I&D Results --")
for r in results: print(r)
results.clear()

# ── Summary ───────────────────────────────────────────────────────────────────
all_passed = all(r.startswith("[PASS]") for r in
    [l for l in open(__file__).read().split('\n') if False])  # placeholder

# Re-collect by running checks again
print("\n-- Done. Check for any [FAIL] above. --")
