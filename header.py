import openpyxl
from openpyxl import load_workbook
from openpyxl.compat import range
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils import coordinate_from_string
from openpyxl.styles import Font


def header():
	italic24Font = Font(size=18, italic=False)
	wb=load_workbook('testing.xlsx',read_only = False, data_only = True)
	sheet =wb.get_sheet_by_name('Sheet')
	ws=wb.active
	col=get_column_letter(1)# convert column number to letter and use for first column (ID card data)
	col2=get_column_letter(2)# use for second column (Student name)
	col3=get_column_letter(3)#use for third column(occurance)
	sheet['a1'].font=italic24Font
	sheet['b1'].font=italic24Font
	sheet.column_dimensions['A'].width = 20
	sheet.column_dimensions['B'].width = 20
	sheet['a1']="UID"
	sheet['b1']="Name"
	wb.save('testing.xlsx')

