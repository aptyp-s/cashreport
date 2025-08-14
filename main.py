from cbr_exchange import get_rates

date = "11/08/2025"
currency = get_rates(date)
print(currency)