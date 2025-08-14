import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.cell import MergedCell
from openpyxl.formula.translate import Translator
from copy import copy
import os
import re

def find_excel_file_in_current_dir():
    excel_files = [f for f in os.listdir('.') if f.endswith('.xlsx') and not f.startswith('~')]
    if len(excel_files) == 0:
        raise FileNotFoundError("Source file not found.")
    if len(excel_files) > 1:
        print(f"More than one Excel file found: {excel_files}")
        print(f"Using the first file: {excel_files[0]}")
    
    return excel_files[0]

def find_anchor_column(sheet, anchor_text="Rate from CBR"):
    for row in sheet.iter_rows(max_row=10):
        for cell in row:
            if cell.value and anchor_text in str(cell.value):
                print(f"Target column is {cell.column}.")
                return cell.column
    return None

def copy_cell_style(source_cell, target_cell):
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = source_cell.number_format
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)

def update_daily_sheet(
    sheet_name,
    new_date,
    new_key_rate,
    new_rub_usd,
    new_rub_eur,
    new_rub_cny
):
    """Main function to update the report."""
    try:
        excel_path = find_excel_file_in_current_dir()
        print(f"Working with file: '{excel_path}'")
    except FileNotFoundError as e:
        print(e)
        return None

    print("Loading data from Excel...")
    wb_formulas = openpyxl.load_workbook(excel_path)
    if sheet_name not in wb_formulas.sheetnames:
        print(f"Error: Sheet with name '{sheet_name}' not found.")
        return None
    sheet_formulas = wb_formulas[sheet_name]
    
    wb_values = openpyxl.load_workbook(excel_path, data_only=True)
    sheet_values = wb_values[sheet_name]

    anchor_col_index = find_anchor_column(sheet_values)
    if not anchor_col_index:
        print(f"Error: Anchor text 'Rate from CBR' not found.")
        return None

    new_column_index = anchor_col_index 
    previous_column_index = new_column_index - 1
    
    print(f"Previous column: {get_column_letter(previous_column_index)}, New column: {get_column_letter(new_column_index)}")

    sheet_formulas.insert_cols(new_column_index)
    print("New column inserted. Processing rows from 3 to 30...")
    
    for row_idx in range(3, 31):
        previous_cell_formulas = sheet_formulas.cell(row=row_idx, column=previous_column_index)
        new_cell_formulas = sheet_formulas.cell(row=row_idx, column=new_column_index)
        
        if isinstance(previous_cell_formulas, MergedCell):
            continue

        previous_cell_static_value = sheet_values.cell(row=row_idx, column=previous_column_index).value
        
        copy_cell_style(previous_cell_formulas, new_cell_formulas)

        # Step 1: Fill the new column
        if row_idx == 3:
            new_cell_formulas.value = new_date
            new_cell_formulas.number_format = 'M/D/YYYY'
        elif row_idx == 4:
            new_cell_formulas.value = new_key_rate / 100 
            new_cell_formulas.number_format = '0.0%' # Excel will add hundredths if necessary
        elif row_idx == 7:
            new_cell_formulas.value = new_rub_usd
        elif row_idx == 8:
            new_cell_formulas.value = new_rub_eur
        elif row_idx == 9:
            new_cell_formulas.value = new_rub_cny
        # --- This is the key block for formulas ---
        else:
            # If the source cell has a formula (e.g., SUMIF), we need to translate it.
            # openpyxl does not automatically shift relative references on simple copy.
            if previous_cell_formulas.data_type == 'f':
                # Create a translator for the formula from the previous cell
                translator = Translator(previous_cell_formulas.value, origin=previous_cell_formulas.coordinate)
                # Translate the formula to the new cell's coordinate (1 column to the right)
                new_cell_formulas.value = translator.translate_formula(new_cell_formulas.coordinate)
            else:
                # If it's not a formula (just text or a number), copy the value.
                new_cell_formulas.value = previous_cell_formulas.value # type: ignore

        # Step 2: "Freeze" the previous column by replacing the formula/value with its static value
        previous_cell_formulas.value = previous_cell_static_value
            
    # Set the column width
    new_col_letter = get_column_letter(new_column_index)
    prev_col_letter = get_column_letter(previous_column_index)
    if sheet_formulas.column_dimensions[prev_col_letter].width:
        sheet_formulas.column_dimensions[new_col_letter].width = sheet_formulas.column_dimensions[prev_col_letter].width
    
    pattern = r"\d{2}.\d{2}.\d{4}"
    match = re.search(pattern, excel_path)
    if match:
        output_path = excel_path.replace(match.group(0), new_date.strftime("%d.%m.%Y"))
    else:
        output_path = excel_path.replace('.xlsx', '_updated.xlsx')
    try:
        wb_formulas.save(output_path)
        print(f"\n{sheet_name} sheet has been successfully updated.")
    except PermissionError:
        print(f"\nError: Could not save file '{output_path}'. Please close Excel and try again.")
        return
    return output_path