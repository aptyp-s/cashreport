from cbr_exchange import get_rates, get_keyrate
from date import get_filename, date_extract, date_fallback
from datetime import datetime
filename = get_filename()
date_str = date_extract(filename)
if date_str == "":
    date_str = date_fallback()
currency = get_rates(date_str)
print(currency)

KR_date = datetime.strptime(date_str, '%d/%m/%Y').date()
keyrate_str = get_keyrate(KR_date)
if keyrate_str is None:
    print("\nCouldn't get key rate from the web. Defaults to 18%.")
    keyrate_str = "18.00"
keyrate = float(keyrate_str)/100
print(f"Key rate at the date (dd/mm/YYYY) {date_str} is {keyrate*100}%.")