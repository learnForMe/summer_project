import openpyxl
import datetime
from openpyxl import load_workbook
from openpyxl.compat import range
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils import coordinate_from_string
from openpyxl.styles import Font


def header():
	month="{:%B %Y}".format(datetime.date.today())
	italic24Font = Font(size=18, italic=False)
	wb=load_workbook('testing.xlsx',read_only = False, data_only = True)
	sheet =wb.get_sheet_by_name('Sheet')
	wb.active
	sheet['a1'].font=italic24Font
	sheet['b1'].font=italic24Font
	sheet.column_dimensions['A'].width = 20
	sheet.column_dimensions['B'].width = 20
	sheet['a1']="UID"
	sheet['b1']="Name"
	curr_col = sheet.max_column
	new_col = curr_col +1 
	new_col= get_column_letter(new_col)
	this_month=get_column_letter(curr_col)
	if sheet.cell('%s1' % this_month).value != month:
		sheet['%s1' % this_month] = month
	else:
		pass
	#print sheet.cell('%s1' % this_month).value 
	sheet.column_dimensions['%s' % new_col].width = 20
	wb.save('testing.xlsx')

#header()
