
import openpyxl
from openpyxl import load_workbook
from openpyxl.compat import range
from openpyxl.utils import get_column_letter, column_index_from_string
#from openpyxl.cell import get_column_letter

f=open ('lastday.txt', 'r')
text=f.read()
texting=text.split('/n')

print texting






wb= load_workbook("testing.xlsx")
sheet1=wb.get_sheet_by_name('Sheet1')
for i in range(1,12):
	print (i, sheet1.cell(row=i, column=1).value)
	print column_index_from_string('c')
#print sheet1.max_column

ws=wb.active
ws['B2']="heyyyyyy"

wb.save("testing.xlsx")

f.closed
