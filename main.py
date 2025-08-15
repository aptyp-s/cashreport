from cbr_exchange import get_rates, get_keyrate
from daily import update_daily_sheet
from helper import get_filename, date_extract, date_fallback, find_excel_file_in_current_dir, file_save
from table import table_new_column
import openpyxl
from datetime import datetime

usd = 'USD/RUB'
eur = 'EUR/RUB'
cny = 'CNY/RUB'
thb = 'THB/RUB'
sheet_1 = 'Daily'
sheet_2 = 'Cash in bank report'
sheet_3 = 'Table'
default_keyrate = 18.0

filename = get_filename()
date_str = date_extract(filename)
if date_str is None:
    date_str = date_fallback()
KR_date = datetime.strptime(date_str, '%d/%m/%Y').date()
currency = get_rates(date_str)

keyrate = get_keyrate(KR_date)
if keyrate is None:
    keyrate = default_keyrate

try:
    excel_path = find_excel_file_in_current_dir()
    wb_formulas = openpyxl.load_workbook(excel_path)
    wb_values = openpyxl.load_workbook(excel_path, data_only=True)
    
    update_daily_sheet(
        wb_formulas,
        wb_values,
        sheet_1,
        KR_date,
        keyrate,
        currency[usd],
        currency[eur],
        currency[cny]
    )
    table_new_column(wb_formulas, wb_values, sheet_3, KR_date)
    file_save(excel_path, KR_date, wb_formulas)
except FileNotFoundError as e:
    print(e)
except Exception as e:
    print(f"An unexpected error occurred during the process: {e}")
