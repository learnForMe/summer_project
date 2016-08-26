import openpyxl
import datetime
from datetime import timedelta
from calendar import mdays
from openpyxl import load_workbook
from openpyxl.compat import range
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils import coordinate_from_string
from openpyxl.styles import Font

def logs (a,b):
	wb=load_workbook('report.xlsx',read_only = False, data_only = True)
	sheet =wb.get_sheet_by_name('Sheet')
	wb.active
	sheet['a1']="August 2016"
	curr_col= sheet.max_column
	curr_row = sheet.max_row
	curr_col = get_column_letter(curr_col)
	sheet.column_dimensions['%s' %  curr_col].width = 50
	time=datetime.datetime.now()
	if a != None:
		new_row=curr_row+1
		sheet['%s%d' % (curr_col,new_row)] = a+ " " + b +" "+ str(time)
	
	wb.save('report.xlsx')
	now = datetime.datetime.now().time()
	print (now)



