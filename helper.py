import re
import glob
import os
import datetime as DT
from copy import copy

def get_filename(fixed_part):
    folder_path = "source"
    pattern = f"{folder_path}/{fixed_part}*.xlsx"
    matching_files = glob.glob(pattern)
    if matching_files:
        found_filename = matching_files[0]
        print(f"Найден файл с датой отчета: {found_filename}")
    else:
        found_filename = ""
        print("Файл не найден.")
    return found_filename

def date_extract(filename):
    pattern = r"\d{8}"
    match = re.search(pattern, filename)
    if match:
        date_str = match.group(0)
        date_object = DT.datetime.strptime(date_str, '%d%m%Y')
        new_date_str = date_object.strftime('%d/%m/%Y')
        print(f"Дата отчета: {date_object.strftime('%d.%m.%Y')} ")
        return new_date_str
    else:
        print("Date not found")
        return None

def date_fallback():
    print(f"Сегодня {DT.date.today().strftime('%d.%m.%Y')}.\n")
    prompt = "Введите число x (0-7), чтобы получить курсы валют на дату за x дней до текущего дня (0 = на сегодня) или дату в формате ГГГГ-ММ-ДД."
    date = None
    while date is None:
        user_date = input(prompt)
        try:
            delta = int(user_date)
            if 0 <= delta <= 7:
                date_temp = DT.date.today() - DT.timedelta(days=delta)
                date = str(date_temp.strftime("%d/%m/%Y"))
                if delta == 0:
                    print('Используется текущая дата.')
                elif delta == 1:
                    print("Используется вчерашняя дата")
                else:
                    print(f"Используется дата за {delta} дней до текущего.")
            else:
                print('Число за пределами диапазона (0-7).')
        except ValueError:
            try:
                date_temp = DT.date.fromisoformat(user_date)
                date = str(date_temp.strftime("%d/%m/%Y"))
            except ValueError:
                print('Неверный формат даты, попробуйте снова: введите ГГГГ-ММ-ДД (например, 2025-06-30) для конкретной даты или число (0-7) для даты относительно текущей.') 
    return date

def find_excel_file_in_current_dir():
    excel_files = [f for f in os.listdir('.') if f.endswith('.xlsx') and not f.startswith('~')]
    if len(excel_files) == 0:
        raise FileNotFoundError("Исходный файл Excel не найден.")
    if len(excel_files) > 1:
        print(f"Найдено более одного Excel файла: {excel_files}")
        print(f"Используется первый файл: {excel_files[0]}")
    else:
        print(f"Используется файл: '{excel_files[0]}'")
    
    return excel_files[0]

def file_save(excel_path, new_date, wb_formulas):
    pattern = r"\d{2}.\d{2}.\d{4}"
    match = re.search(pattern, excel_path)
    if match:
        output_path = excel_path.replace(match.group(0), new_date.strftime("%d.%m.%Y"))
    else:
        output_path = excel_path.replace('.xlsx', '_updated.xlsx')
    try:
        wb_formulas.save(output_path)
        print(f"\nОбновленный отчет {output_path} успешно сохранен.")
    except PermissionError:
        print(f"\nОшибка доступа к файлу '{output_path}'. Пожалуйста, закройте его в Excel и попробуйте снова.")
        return
    
def copy_cell_style(source_cell, target_cell):
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = source_cell.number_format
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)