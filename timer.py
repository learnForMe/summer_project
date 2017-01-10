from time import sleep
import openpyxl
from openpyxl import load_workbook
from openpyxl.compat import range
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils import coordinate_from_string


def prompt_with_timeout():
  print ("\t######### MISSING Name or Email ##########\n")

  print('Please press Ctrl-C to enter Name and Email')
  try:
    for i in range(0, 10): # 30 minutes is 30*60 seconds
      sleep(1)
    print("No Name or Email was logged")
    n = ""
    e = ""
    
  except KeyboardInterrupt:
    n = input("Student's Name -> ")
    e = input("Student's Email -> ")
    if n == "" or  e == "" or n == "None" or  e == "None":
      print("No Name or Email was logged")
    else:
      print (n + " and " + e +" added to database")
  return n, e


