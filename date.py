import re
import glob
import datetime as DT

def get_filename():
    folder_path = "source"
    fixed_part = "Cash report_"
    pattern = f"{folder_path}/{fixed_part}*.xlsx"
    matching_files = glob.glob(pattern)
    if matching_files:
        found_filename = matching_files[0]
        print(f"Found file to extract date: {found_filename}")
    else:
        found_filename = ""
        print("File not found")
    return found_filename

def date_extract(filename):
    pattern = r"\d{8}"
    match = re.search(pattern, filename)
    if match:
        date_str = match.group(0)
        date_object = DT.datetime.strptime(date_str, '%d%m%Y')
        new_date_str = date_object.strftime('%d/%m/%Y')
        print(f"Report date: {new_date_str} ")
    else:
        new_date_str = ""
        print("Date not found")
    return new_date_str

def date_fallback():
    print(f"Today is {DT.date.today()}.\n")
    prompt = "Input a number x (0-7) for rates x days ago (0 = today) "
    prompt += "or a date in ISO format (YYYY-MM-DD): "
    date = None
    while date is None:
        user_date = input(prompt)
        try:
            delta = int(user_date)
            if 0 <= delta <= 7:
                date_temp = DT.date.today() - DT.timedelta(days=delta)
                date = str(date_temp.strftime("%d/%m/%Y"))
                if delta == 0:
                    print('Using today as date.')
                elif delta == 1:
                    print("Using yesterday as date.")
                else:
                    print(f"Using {delta} days ago as date.")
            else:
                print('Number out of range (0-7).')
        except ValueError:
            try:
                date_temp = DT.date.fromisoformat(user_date)
                date = str(date_temp.strftime("%d/%m/%Y"))
            except ValueError:
                print('Wrong date format, try again: e.g. 2025-06-30 for '
                    'absolute date or a number (0-7) for relative date.') 
    return date