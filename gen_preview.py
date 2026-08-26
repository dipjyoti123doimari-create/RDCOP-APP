import sys
sys.path.insert(0, 'd:/Dipjyoti Doimari/AI/Incentive_Calculator')
from app import _tp_build_html_tables

plant_rows = [
    {'plant_name':'NCR-Gurgaon Badshahpur_BP2','exco_location':'NCR','business_head':'Japjit Singh','plant_manager':'Vishal Saraswat','mixer_theo_cap':60,'total_quantity':466,'total_time_min':2638,'throughput_pct':18,'batch_count':10},
    {'plant_name':'ODI-TPL Talcher','exco_location':'Odisha','business_head':'Naresh C','plant_manager':'Rajeshwar S','mixer_theo_cap':60,'total_quantity':107,'total_time_min':531,'throughput_pct':20,'batch_count':5},
    {'plant_name':'GOA-Goa 2','exco_location':'Goa','business_head':'Adnan Khan','plant_manager':'Satyam Jha','mixer_theo_cap':56,'total_quantity':982.8,'total_time_min':4958,'throughput_pct':21,'batch_count':18},
    {'plant_name':'MUM-Deonar','exco_location':'Mumbai','business_head':'Raj Kumar','plant_manager':'Kumar Singh','mixer_theo_cap':50,'total_quantity':3200,'total_time_min':1800,'throughput_pct':67,'batch_count':64},
    {'plant_name':'PUN-TPL Metro 1_BP2','exco_location':'Pune','business_head':'Mangesh/Adnan','plant_manager':'Abhijit Pawar','mixer_theo_cap':40,'total_quantity':2100,'total_time_min':1400,'throughput_pct':78,'batch_count':45},
    {'plant_name':'BLR-Whitefield','exco_location':'Bangalore','business_head':'Rahul','plant_manager':'Sharma','mixer_theo_cap':50,'total_quantity':4200,'total_time_min':2100,'throughput_pct':88,'batch_count':80},
]
location_rows = [
    {'exco_location':'NCR','plant_count':2,'total_quantity':583,'total_time_min':3169,'avg_throughput_pct':19,'is_pan_india':False},
    {'exco_location':'Goa','plant_count':1,'total_quantity':983,'total_time_min':4958,'avg_throughput_pct':21,'is_pan_india':False},
    {'exco_location':'Mumbai','plant_count':14,'total_quantity':30249,'total_time_min':15000,'avg_throughput_pct':41,'is_pan_india':False},
    {'exco_location':'Pune','plant_count':9,'total_quantity':10789,'total_time_min':8000,'avg_throughput_pct':56,'is_pan_india':False},
    {'exco_location':'Bangalore','plant_count':17,'total_quantity':28598,'total_time_min':18000,'avg_throughput_pct':76,'is_pan_india':False},
    {'exco_location':'PAN India','plant_count':43,'total_quantity':71202,'total_time_min':49127,'avg_throughput_pct':47,'is_pan_india':True},
]

tables = _tp_build_html_tables(plant_rows, location_rows, 6, 2026)
html = (
    '<!DOCTYPE html><html><body style="background:#fff;padding:20px;font-family:Arial,sans-serif">'
    + tables +
    '</body></html>'
)
with open('d:/Dipjyoti Doimari/email_preview.html', 'w', encoding='utf-8') as f:
    f.write(html)
print('Saved: d:/Dipjyoti Doimari/email_preview.html')
