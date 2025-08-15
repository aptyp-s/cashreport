import requests
import datetime as DT
import xml.etree.ElementTree as ET

def get_rates(date):
    # parse CBR website to get XML response for current date
    url_fixed = 'https://www.cbr.ru/scripts/XML_daily.asp?date_req='
    # # enter date below (will import later)
    
    print(f"CBR exchange rates for the date: {date}")
    url_full = url_fixed + date
    url_headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
                    '(KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
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

    try:
        response = requests.get(url_full, headers=url_headers, timeout=10)
        response.raise_for_status()
        data = response.text
    except requests.exceptions.RequestException as e:
        print(f"ERROR: An request error occurred: {e}")
        print("Returning default exchange rates.")
        currency = {}
        for tech_id, friendly_name in id_to_friendly.items():
            currency[friendly_name] = valute_match.get(tech_id)
        return currency
    except Exception as e:
        print(f"ERROR: An unhandled exception occurred: {e}")
        currency = {}
        for tech_id, friendly_name in id_to_friendly.items():
            currency[friendly_name] = valute_match.get(tech_id)
        return currency
    
    try:
        root = ET.fromstring(data)
    except ET.ParseError as e:
        print(f"Error parsing XML: {e}")
        root = None
    
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
        print("Failed to process XML data, using default values")
        currency = {}
        for tech_id, friendly_name in id_to_friendly.items():
            currency[friendly_name] = valute_match.get(tech_id)
        return currency

    # make a new dictionary with friendly names
    currency = {}
    for tech_id, friendly_name in id_to_friendly.items():
        currency[friendly_name] = valute_match.get(tech_id)
    return currency

# key rate will use a different method
def get_keyrate(target_date):
    datetime_string = target_date.strftime('%Y-%m-%dT%H:%M:%S')
    url = "http://www.cbr.ru/DailyInfoWebServ/DailyInfo.asmx"
    headers = {
    "Host": "www.cbr.ru",
    "Content-Type": "application/soap+xml; charset=utf-8",
    "Content-Length": "length",
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
                    '(KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    soap_body = f"""<?xml version="1.0" encoding="utf-8"?>
<soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">
  <soap12:Body>
    <KeyRateXML xmlns="http://web.cbr.ru/">
      <fromDate>{datetime_string}</fromDate>
      <ToDate>{datetime_string}</ToDate>
    </KeyRateXML>
  </soap12:Body>
</soap12:Envelope>
"""
    response = None
    try:
        response = requests.post(url, headers=headers, data=soap_body.encode('utf-8'))
        response.raise_for_status()
        namespaces = {
            'soap': 'http://www.w3.org/2003/05/soap-envelope',
        }

        root = ET.fromstring(response.content)
        
        # Ищем нужный элемент с учетом пространства имен.
        rate_element = root.find('.//Rate', namespaces)

        if rate_element is not None and rate_element.text:
            # Преобразуем найденное значение в число
            key_rate_value = float(rate_element.text)
            return key_rate_value
        else:
            print("Элемент с ключевой ставкой не найден в ответе.")
            # Это может означать, что на запрашиваемую дату ставка не публиковалась.
            # Можно посмотреть XML ответа для анализа:
            # print(response.text)
            return None

    except requests.exceptions.HTTPError as err:
        # 5. УЛУЧШЕННАЯ ОБРАБОТКА ОШИБОК: Печатаем не только статус, но и тело ответа.
        #    В теле ответа сервер часто присылает детальное описание проблемы.
        print(f"Ошибка HTTP: {err}")
        print("--- Тело ответа сервера ---")
        print(err.response.text)
        print("---------------------------")
        return None
    except requests.exceptions.RequestException as e:
        print(f"Произошла ошибка при отправке запроса: {e}")
        return None
    except ET.ParseError as e:
        print(f"Ошибка парсинга XML: {e}")
        if response:
            print("--- Полученный ответ ---")
            print(response.text)
            print("------------------------")
        return None