import openpyxl

from openpyxl import load_workbook
from openpyxl.compat import range
from openpyxl.utils import get_column_letter, column_index_from_string



wb=load_workbook('testing.xlsx',read_only = False, data_only = True)
sheet = wb.get_sheet_by_name('Sheet1')
ws=wb.active

max_row = sheet.get_highest_row()
insert_name=max_row+1
print insert_name


row=1
column=4

for i in range(1,max_row):
	new=(i,sheet.cell(row=i,column=1).value) 
	if new != 10:
		print "Fresh Meat"



