import openpyxl
import re
import os
from openpyxl import load_workbook
from openpyxl.compat import range
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils import coordinate_from_string


def remove(y):
    
    cell_location=re.sub(r'[\D]','',y)
    return cell_location

def search_Student (x):
    wb=load_workbook('testing.xlsx',read_only = False, data_only = True)
    worksheet= wb.get_sheet_names()
    wb.active
    sheet = wb.get_sheet_by_name('Sheet')
    max_row = sheet.max_row
    insert_name=max_row+1

    max_col=sheet.max_column-1
    if max_col != 3:
        max_col=3

    col=get_column_letter(1)# convert column number to letter and use for first column (ID card data)
    col2=get_column_letter(2)# use for second column (Student name)
    col3=get_column_letter(max_col)#use for third column(occurance)
    count =0
    first_time =1
    for row in sheet.iter_rows():
        for cell in row:
            data = cell.value
            
            if data == str(x):
                cell = str(cell)
                cool=int(remove(cell))
                stop_by= sheet.cell('%s%d' % (col3,cool)).value
                if stop_by == None:
                    stop_by = 0
                #print (stop_by)
                
                sheet['%s%d' % (col3,cool)] = stop_by+1
                wb.save('testing.xlsx')
                count+=1
              
    if count< 1:
        print ('\a\a\a\a\a\a')
        os.system ('clear')
        #os.system ('echo "New Student"')
        print ("NEW STUDENT\n")
        #print ("This Student is NEW")
        student_name=raw_input("Enter Name -> ")
        #student_name= input("Enter Name -> ")
        print (student_name + " added to database")
        
        sheet['%s%d' % (col2,insert_name)] =str(student_name)
        sheet['%s%d' % (col,insert_name)] =str(x)
        sheet['%s%d' % (col3,insert_name)] =first_time
        print (x)
        wb.save('testing.xlsx')
    else:
        print ('\a')
        again=sheet.cell('%s%d' % (col2,cool)).value
        print ("Welcome Back! "+ again)
        #os.system ('say What is up%s' % again)  # use for prank (aka April Foo)
