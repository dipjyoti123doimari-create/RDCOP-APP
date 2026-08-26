import sys
sys.path.insert(0, 'd:/Dipjyoti Doimari/AI/Incentive_Calculator')
import pandas as pd
from report_generator import build_email_tables_html

# Minimal test data covering deduction (red), incentive (green), plain rows
data = {
    'month': [6,6,6,6,6,6],
    'year':  [2026]*6,
    'employee_code': ['E001','E002','E003','E004','E005','E006'],
    'employee_name': ['Alice','Bob','Carol','Dave','Eve','Frank'],
    'designation': ['Trainee','Trainee','PM & API','QCI','MO','SPE'],
    'category': ['Civil Trainee','Civil Trainee','PM & API','QCI','MO','SPE'],
    'plant': ['Plant A','Plant B','Plant A','Plant C','Plant B','Plant D'],
    'plant_code': ['PA','PB','PA','PC','PB','PD'],
    'total_quantity': [120,80,200,90,150,300],
    'ytd_maintenance_cost': [0,0,0,0,0,0],
    'incentive_eligible': [1,0,1,0,1,1],
    'incentive_rate': [10,0,15,0,12,20],
    'incentive_amount': [1200,0,3000,0,1800,6000],
    'deduction_target': [0,100,0,50,0,0],
    'shortfall_quantity': [0,20,0,10,0,0],
    'deduction_amount': [0,400,0,200,0,0],
    'remarks': ['OK','Short','OK','Short','OK','OK'],
}
df = pd.DataFrame(data)

tables = build_email_tables_html(df)
html = (
    '<!DOCTYPE html><html><body style="background:#fff;padding:20px;font-family:Arial,sans-serif">'
    + tables +
    '</body></html>'
)
with open('d:/Dipjyoti Doimari/id_email_preview.html', 'w', encoding='utf-8') as f:
    f.write(html)
print('Saved: d:/Dipjyoti Doimari/id_email_preview.html')
