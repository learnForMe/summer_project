import openpyxl

from openpyxl import load_workbook
from openpyxl.compat import range
from openpyxl.utils import get_column_letter, column_index_from_string

f=open ('lastday.txt', 'r')
text=f.read()
#texting=text.split('/n')
texting ="g"

wb=load_workbook('testing.xlsx',read_only = False, data_only = True)
sheet = wb.get_sheet_by_name('Sheet1')
ws=wb.active
count =0
column=1
col=get_column_letter(column)

max_row = sheet.max_row
insert_name=max_row+1

#insert_name=max_row+1
#print insert_name



#row=1
#column=4
'''
for i in range(1,insert_name):
	new=(i,sheet.cell(row=i,column=1).value)
	print new
	'''
for row in sheet.iter_rows(): 
        for cell in row:
            data = cell.value 
            print data
            if data == str(texting):
				count+=1
				
if count< 1:
	print "This Student is NEW"
	ws['%s%d' % (col,insert_name)] =str(texting)
	wb.save('testing.xlsx')
else:
	print "Not New"            	

            	
#wb.save('testing.xlsx')
f.closed
