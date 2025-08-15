from openpyxl.formula.translate import Translator
from openpyxl.cell import MergedCell
from openpyxl.utils import get_column_letter
from daily import copy_cell_style

def table_new_column(wb_formulas, wb_values, sheet_name, report_date):
    if sheet_name not in wb_formulas.sheetnames:
        print(f"Лист с именем '{sheet_name}' не найден.")
        return None
           
    ws = wb_formulas[sheet_name]
    sheet_values = wb_values[sheet_name]
    print(f"Работаю с листом {sheet_name}...")

    last_col_idx = ws.max_column
    new_col_idx = last_col_idx + 1

    print(f"Последний столбец с данными: {get_column_letter(last_col_idx)}. Новый столбец: {get_column_letter(new_col_idx)}.")
    
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
