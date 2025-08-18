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
        print(f"Найден файл: {found_filename}")
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

def divide(value):
    """Преобразует числовое значение, деля его на 1,000,000."""
    # Если ячейка пустая (None) или содержит нечисловые данные, возвращаем 0
    if not isinstance(value, (int, float)):
        print(f"Предупреждение: значение '{value}' не является числом. Используется 0.")
        return 0.0
    return value / 1_000_000

def clean_and_convert_to_float(value):
    """Очищает строку от пробелов, заменяет запятую на точку и преобразует в float."""
    if value is None:
        return 0.0
    
    # Преобразуем значение в строку для безопасной обработки
    s_value = str(value)
    
    # Убираем пробелы (включая неразрывные \xa0) и меняем запятую
    cleaned_s_value = s_value.replace('\xa0', '').replace(' ', '').replace(',', '.')
    
    try:
        return float(cleaned_s_value)
    except (ValueError, TypeError):
        # Если значение не является числом (например, текст), возвращаем 0
        return 0.0
    
def update_formula_and_compare(target_ws, cell_address, new_value, currency_name):

    target_cell = target_ws[cell_address]
    existing_content = target_cell.value
    old_value = 0.0

    if not existing_content:
        print(f"  - Предупреждение: Ячейка {cell_address} пуста. Обновление пропущено.")
        return False

    try:
        # --- ЛОГИКА ОПРЕДЕЛЕНИЯ ТИПА СОДЕРЖИМОГО ---
        if str(existing_content).startswith('='):
            # --- СЛУЧАЙ 1: ЭТО ФОРМУЛА ---
            parts = str(existing_content).split('*', 1)
            if len(parts) < 2:
                print(f"  - Предупреждение: Неожиданный формат формулы в {cell_address}. Обновление пропущено.")
                return False
            
            old_value_str = parts[0].replace('=', '').strip()
            old_value = float(old_value_str)
            
            static_part = parts[1]
            new_formula = f"={new_value}*{static_part}"
            target_cell.value = new_formula
            print(f"  - {currency_name}: Формула в {cell_address} обновлена. Новое значение {new_value:.2f} vs старое {old_value:.2f}.")

        else:
            # --- СЛУЧАЙ 2: ЭТО ПРОСТОЕ ЗНАЧЕНИЕ ---
            old_value = float(existing_content)
            target_cell.value = new_value
            print(f"  - {currency_name}: Значение в {cell_address} обновлено. Новое {new_value:.2f} vs старое {old_value:.2f}.")

        # Возвращаем результат сравнения
        return new_value > old_value

    except (ValueError, IndexError) as e:
        print(f"  - Ошибка при обработке ячейки {cell_address}: {e}. Обновление пропущено.")
        return False