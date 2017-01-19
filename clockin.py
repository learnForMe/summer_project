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
		wb = openpyxl.load_workbook("clocking.xlsx")
		sheet =wb.get_sheet_by_name('Sheet')
		sheet= wb.active
		#sheet["a1"]= str(self.register())
		next_row= sheet.max_row+1
		sheet["a%d"%next_row]= self.clockin()
		wb.save("clocking.xlsx")
	'''
	def register(self):
		wb = openpyxl.load_workbook("clocking.xlsx")
		sheet =wb.get_sheet_by_name('Sheet')
		sheet= wb.active
		count=1
		#x= input("Name for first Work-Study ->")
		#self.name= x
		#sheet["%s1"% get_column_letter(count)]= self.name
		if sheet.cell('%s1'% get_column_letter(count)).value==None:
			x= input("Name for first Work-Study ->")
			self.name= x
			sheet["%s1"% get_column_letter(count)]= self.name

		wb.save("clocking.xlsx")
		'''
	def register(self,x,y):
		wb = openpyxl.load_workbook("clocking.xlsx")
		sheet =wb.get_sheet_by_name('Sheet')
		sheet= wb.active
		count=1
		#x= input("Name for first Work-Study ->")
		#self.name= x
		#sheet["%s1"% get_column_letter(count)]= self.name
		while count <= 4:
			if sheet.cell('%s1'% get_column_letter(count)).value==None:
				x= input("Name for first Work-Study ->")
				self.name= x
				sheet["%s1"% get_column_letter(count)]= x
				sheet["%s2"% get_column_letter(count)]= y
			count = count +1

		wb.save("clocking.xlsx")
		
		

	#def grep(self,c,d):

def main():
	
	first=timesheet()
	first.clockin()
	first.register()
	first.excel()
	

if __name__=='__main__':
	main()

