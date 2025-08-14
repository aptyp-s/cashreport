import requests
import datetime as DT
import xml.etree.ElementTree as ET

def get_rates(date):
    # parse CBR website to get XML response for current date
    url_fixed = 'https://www.cbr.ru/scripts/XML_daily.asp?date_req='
    # # enter date below (will import later)
    # print(f"Today is {DT.date.today()}.\n")
    # prompt = "Input a number x (0-7) for rates x days ago (0 = today) "
    # prompt += "or a date in ISO format (YYYY-MM-DD): "
    # date = None
    # while date is None:
    #     user_date = input(prompt)
    #     try:
    #         delta = int(user_date)
    #         if 0 <= delta <= 7:
    #             date_temp = DT.date.today() - DT.timedelta(days=delta)
    #             date = str(date_temp.strftime("%d/%m/%Y"))
    #             if delta == 0:
    #                 print('Using today as date.')
    #             elif delta == 1:
    #                 print("Using yesterday as date.")
    #             else:
    #                 print(f"Using {delta} days ago as date.")
    #         else:
    #             print('Number out of range (0-7).')
    #     except ValueError:
    #         try:
    #             date_temp = DT.date.fromisoformat(user_date)
    #             date = str(date_temp.strftime("%d/%m/%Y"))
    #         except ValueError:
    #             print('Wrong date format, try again: e.g. 2025-06-30 for '
    #                 'absolute date or a number (0-7) for relative date.')  

    print(f"CBR exchange rates for the date (dd/mm/YYYY): {date}")
    url_full = url_fixed + date
    url_headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
                    '(KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    try:
        response = requests.get(url_full, headers=url_headers, timeout=10)
        response.raise_for_status()
        data = response.text
    except requests.exceptions.ConnectionError as e:
        print(f"ERROR: Connection error occurred: {e}")
        data = '0'
    except requests.exceptions.Timeout as e:
        print(f"ERROR: Request timed out after 10 seconds: {e}")
        data = '0'
    except requests.exceptions.HTTPError as e:
        print(f"ERROR: HTTP error occurred - Status Code: {e.response.status_code}")
        print(f"Response content: {e.response.text}")
        data = '0'
    except requests.exceptions.RequestException as e:
        print(f"ERROR: An unexpected requests error occurred: {e}")
        data = '0'
    except Exception as e:
        print(f"ERROR: An unhandled exception occurred: {e}")
        data = '0'

    try:
        root = ET.fromstring(data)
    except ET.ParseError as e:
        print(f"Error parsing XML: {e}")
        root = None

    # find target IDs in XML and extract their values
    id_to_friendly = {
        'R01235': 'USD/RUB',
        'R01239': 'EUR/RUB',
        'R01375': 'CNY/RUB',
        'R01675': 'THB/RUB'
    }

    target_ids = list(id_to_friendly.keys())

    def_values=[79.0,90.0,11.0,2.4]
    valute_match = dict(zip(target_ids,def_values))

    if root is not None:
        for id_element in root:
            valute_id = id_element.get('ID')
            if valute_id in target_ids:
                vunit_rate_element = id_element.find('VunitRate')
                
                if vunit_rate_element is not None:
                    try:
                        if vunit_rate_element.text:
                            clean_text = vunit_rate_element.text.replace(',', '.')
                        else: 
                            clean_text = '0'
                        valute_match[valute_id] = float(clean_text)
                    except ValueError as e:
                        print(f"Could not convert {vunit_rate_element.text} "
                            f"to float for ID {valute_id}: {e}")
                        valute_match[valute_id] = 0.0
                else:
                    print(f"Rate not found for ID {valute_id}, using default value")
    else:
        print("Failed to process XML data due to parsing error")

    # make a new dictionary with friendly names
    currency = {}

    for tech_id, friendly_name in id_to_friendly.items():
        currency[friendly_name] = valute_match.get(tech_id)

    return currency