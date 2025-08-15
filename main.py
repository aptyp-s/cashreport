from cbr_exchange import get_rates, get_keyrate
from daily import update_daily_sheet
from helper import get_filename, date_extract, date_fallback, find_excel_file_in_current_dir, file_save
from table import table_new_column
import openpyxl
from datetime import datetime
filename = get_filename()
date_str = date_extract(filename)
if date_str is None:
    date_str = date_fallback()
currency = get_rates(date_str)
print(currency)
usd = 'USD/RUB'
eur = 'EUR/RUB'
cny = 'CNY/RUB'
thb = 'THB/RUB'

KR_date = datetime.strptime(date_str, '%d/%m/%Y').date()
default_keyrate = "18.00"
keyrate_str = get_keyrate(KR_date)
if keyrate_str is None:
    print("\nCouldn't get key rate from the web. Defaults to 18%.")
    keyrate_str = default_keyrate
keyrate = float(keyrate_str)
print(f"CBR key rate for the date: {keyrate}%.")

try:
    excel_path = find_excel_file_in_current_dir()
    wb_formulas = openpyxl.load_workbook(excel_path)
    wb_values = openpyxl.load_workbook(excel_path, data_only=True)
    
    update_daily_sheet(
        wb_formulas,
        wb_values,
        KR_date,
        keyrate,
        currency[usd],
        currency[eur],
        currency[cny]
    )
    table_new_column(wb_formulas, wb_values, KR_date)
    file_save(excel_path, KR_date, wb_formulas)
except FileNotFoundError as e:
    print(e)
except Exception as e:
    print(f"An unexpected error occurred during the process: {e}")
