import openpyxl

from openpyxl import load_workbook
from openpyxl.compat import range
from openpyxl.utils import get_column_letter, column_index_from_string

while True:

	wb=load_workbook('testing.xlsx')
	sheet = wb.get_sheet_by_name('Sheet1')
	ws=wb.active

	row=1
	column=4

	for i in range(1,11):
		row+=1
		col=get_column_letter(column)
		ws['%s%d' % (col,row)] = 'Hello world!'
		wb.save('testing.xlsx')



