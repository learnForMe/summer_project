import openpyxl
import datetime
from openpyxl import load_workbook
from openpyxl.compat import range
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils import coordinate_from_string
# use crontab -e
# 0 0 1 * *  python /Users/garytsai/Desktop/rfid-reader-http/summer_project/autoEmail.py
def add_column():
	wb=load_workbook('testing.xlsx',read_only = False, data_only = True)
	sheet = wb.get_sheet_by_name('Sheet')
	ws=wb.active
	sheet_count = sheet.max_column
	sheet_count += 1
	#new_col=sheet_count
	new_col= get_column_letter(sheet_count)
	month="{:%B %Y}".format(datetime.date.today())
	#print new_col
	ws['%s1' % new_col] = month
	wb.save('testing.xlsx')