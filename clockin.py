import openpyxl
import datetime
from datetime import timedelta
from calendar import mdays
from openpyxl import load_workbook
from openpyxl.compat import range
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils import coordinate_from_string
from openpyxl.styles import Font
from dateutil import relativedelta

class timesheet:
	def clockin(self):
		time=datetime.datetime.now().strftime("%Y/%m/%d %H:%M")
		
		return time

	def excel(self):
		wb = openpyxl.Workbook()
		sheet =wb.get_sheet_by_name('Sheet')
		sheet= wb.active
		#sheet["a1"]= str(self.register())
		next_row= sheet.max_row+1
		sheet["a%d"%next_row]= self.clockin()
		wb.save("clocking.xlsx")

	def register(self):
		x= input("Name for first Work-Study ->")
		self.name= x
		wb = openpyxl.load_workbook("clocking.xlsx")
		sheet =wb.get_sheet_by_name('Sheet')
		sheet= wb.active
		#self.id = y
		count=0
		while count <=4:
			#if
		sheet["a1"]= str(self.name)
		wb.save("clocking.xlsx")
		return self.name

	#def grep(self,c,d):

def main():
	
	first=timesheet()
	first.clockin()
	
	first.excel()






if __name__=='__main__':
	main()

