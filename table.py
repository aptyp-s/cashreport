import openpyxl
from openpyxl.formula.translate import Translator
from openpyxl.cell import MergedCell
from datetime import date
from daily import copy_cell_style

def table_new_column(filename: str, report_date: date):
    """
    Adds a new data column to the 'Table' sheet of the specified Excel file.
    It copies formulas to the new column (and a specific value for row 47), 
    and replaces formulas with static values in the source column.
    Only processes the first 50 rows.
    """
    sheet_name = 'Table'
    try:
        # Load workbook to read formulas and write changes
        wb = openpyxl.load_workbook(filename)
        
        if sheet_name not in wb.sheetnames:
            print(f"Error: Sheet {sheet_name} not found in the workbook '{filename}'.")
            return
            
        ws = wb[sheet_name]

        # Load a separate instance of the workbook in data_only mode to get calculated values
        # This is crucial for correctly "freezing" the formulas
        wb_values = openpyxl.load_workbook(filename, data_only=True)
        sheet_values = wb_values[sheet_name]

        print(f"Updating {sheet_name} sheet...")

        last_col_idx = ws.max_column
        new_col_idx = last_col_idx + 1

        print(f"Last data column found is: {last_col_idx}. New column will be: {new_col_idx}.")
        
        # Replicate merged cell structure if it exists in the last column
        for merged_range in list(ws.merged_cells.ranges):
            if merged_range.min_col <= last_col_idx <= merged_range.max_col:
                col_offset = new_col_idx - last_col_idx
                new_range_start_col = merged_range.min_col + col_offset
                new_range_end_col = merged_range.max_col + col_offset
                ws.merge_cells(
                    start_row=merged_range.min_row,
                    end_row=merged_range.max_row,
                    start_column=new_range_start_col,
                    end_column=new_range_end_col
                )

        # Loop through rows 1 to 50
        for row_num in range(1, 51):
            source_cell = ws.cell(row=row_num, column=last_col_idx)
            dest_cell = ws.cell(row=row_num, column=new_col_idx)
            
            if isinstance(dest_cell, MergedCell):
                continue
            
            # --- THIS IS THE FIX ---
            # Get the pre-calculated value BEFORE modifying the source cell
            source_cell_static_value = sheet_values.cell(row=row_num, column=last_col_idx).value

            # 1) Copy cell style
            copy_cell_style(source_cell, dest_cell)

            # Special case for row 47: copy only the value
            if row_num == 47:
                # Copy the static value to the new column
                dest_cell.value = source_cell_static_value
                continue

            # Process only formula cells for other rows
            if source_cell.data_type == 'f':
                # a) Copy and translate the formula to the new column
                formula = source_cell.value
                translator = Translator(formula, origin=source_cell.coordinate)
                dest_cell.value = translator.translate_formula(dest_cell.coordinate)
                
                # b) Replace the formula in the source cell with its static value
                source_cell.value = source_cell_static_value # type: ignore
                
        # 2) Insert the predefined date into the specified row
        date_row = 1
        date_dest_cell = ws.cell(row=date_row, column=new_col_idx)
        
        if not isinstance(date_dest_cell, MergedCell):
            date_dest_cell.value = report_date
            # Style for the date cell should also be copied
            date_source_cell = ws.cell(row=date_row, column=last_col_idx)
            copy_cell_style(date_source_cell, date_dest_cell)

        # Overwrite the original file
        output_filename = filename
        wb.save(output_filename)
        print(f"Successfully updated {sheet_name} sheet in '{output_filename}'.")

    except FileNotFoundError:
        print(f"Error: The file '{filename}' was not found.")
    except Exception as e:
        print(f"An unexpected error occurred on sheet {sheet_name}: {e}")