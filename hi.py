import openpyxl
wb = openpyxl.load_workbook('testing.xlsx',read_only = True, data_only = True)
print wb.get_sheet_names()