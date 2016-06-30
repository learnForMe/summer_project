import openpyxl
wb = openpyxl.Workbook('testing.xlxs')
sheet = wb.get_sheet_by_name('Sheet1')
sheet['A1'] = 'Hello world!'
sheet['A1'].value