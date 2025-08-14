from cbr_exchange import get_rates, get_keyrate
from daily import *
from date import *
from table import *
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
print(f"CBR key rate at the date (dd/mm/YYYY) {date_str} is {keyrate}%.")

first_update = update_daily_sheet(
    'Daily',
    KR_date,
    keyrate,
    currency[usd],
    currency[eur],
    currency[cny]
)
if first_update:
    table_new_column(first_update, KR_date)

