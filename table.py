import openpyxl
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter, column_index_from_string
from helper import copy_cell_style, get_filename

def table_new_column(wb_formulas, wb_values, sheet_name, report_date):
    if sheet_name not in wb_formulas.sheetnames:
        print(f"Лист с именем '{sheet_name}' не найден.")
        return None
           
    ws = wb_formulas[sheet_name]
    sheet_values = wb_values[sheet_name]
    print(f"Работаю с листом {sheet_name}...")

    last_col_idx = ws.max_column
    new_col_idx = last_col_idx + 1
    new_column_name = get_column_letter(new_col_idx)
    print(f"Последний столбец с данными: {get_column_letter(last_col_idx)}. Новый столбец: {new_column_name}.")
    
    for row_num in range(1, 51):
        source_cell = ws.cell(row=row_num, column=last_col_idx)
        dest_cell = ws.cell(row=row_num, column=new_col_idx)
        
        source_cell_static_value = sheet_values.cell(row=row_num, column=last_col_idx).value

        copy_cell_style(source_cell, dest_cell)

        # Временно для сипи трейдинга
        if row_num == 47:
            dest_cell.value = source_cell_static_value
        
        elif source_cell.data_type == 'f':
            formula = source_cell.value
            translator = Translator(formula, origin=source_cell.coordinate)
            dest_cell.value = translator.translate_formula(dest_cell.coordinate)

            if source_cell_static_value is not None:
                source_cell.value = source_cell_static_value
        
    date_row = 1
    date_dest_cell = ws.cell(row=date_row, column=new_col_idx)
    date_dest_cell.value = report_date
    print("Готово!")
    return new_column_name

def divide(value):
    """Преобразует числовое значение, деля его на 1,000,000."""
    # Если ячейка пустая (None) или содержит нечисловые данные, возвращаем 0
    if not isinstance(value, (int, float)):
        print(f"Предупреждение: значение '{value}' не является числом. Используется 0.")
        return 0.0
    return value / 1_000_000

def copy_cpfo(wb_formulas, column, sheet_name):
    folder_path = "source"
    source_filename = get_filename(fixed_part = "Cash report_")
    try:
        source_wb = openpyxl.load_workbook(source_filename)
        source_ws = source_wb.active
        if source_ws is None:
            print(f"Ошибка: Не удалось найти активный лист в файле '{source_filename}'.")
            return
        source_range = 'B3:G3'
        target_ws = wb_formulas[sheet_name]
        source_values_raw = [cell.value for cell in source_ws[source_range][0]]
        print(f"Прочитанные значения: {source_values_raw}")
        processed_values = [divide(val) for val in source_values_raw]
        print(f"Обработанные значения (разделенные на 1 млн): {processed_values}")
        start_row_idx = 5
        start_col_idx = column_index_from_string(column)
        for i, value in enumerate(processed_values):
            target_ws.cell(row=start_row_idx + i, column=start_col_idx, value=value)
        print(f"Данные успешно скопированы в столбец {column} листа '{sheet_name}'.")
    except FileNotFoundError:
        print(f"Файл {source_filename} не найден.")
    
