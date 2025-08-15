def update_cash_in_bank_core(wb_formulas, sheet_name, report_date, thb_rate):
    if sheet_name not in wb_formulas.sheetnames:
        print(f"Лист с именем '{sheet_name}' не найден.")
        return None
    
    sheet = wb_formulas[sheet_name]
    sheet['B2'] = report_date
    sheet['E5'] = thb_rate
