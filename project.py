
import openpyxl
from openpyxl import load_workbook
from openpyxl.compat import range
from openpyxl.utils import get_column_letter, column_index_from_string
#from openpyxl.cell import get_column_letter

f=open ('lastday.txt', 'r')
text=f.read()
texting=text.split('/n')

wb=load_workbook('testing.xlsx',read_only = False, data_only = True)
sheet = wb.get_sheet_by_name('Sheet1')
ws=wb.active

print texting

row=1
column=4

for i in range(1,11):
	row+=1
	col=get_column_letter(column)
	ws['%s%d' % (col,row)] =str(texting)
	wb.save('testing.xlsx')



f.closed
