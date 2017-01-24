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
		sheet.column_dimensions['a'].width = 25
		sheet.column_dimensions['b'].width = 25
		sheet.column_dimensions['c'].width = 25
		sheet.column_dimensions['d'].width = 25
		#sheet["a1"]= str(self.register())
		next_row= sheet.max_column
		#print(self.column_to_add('a'))
		#print(self.column_to_add('b'))
		
		if self.ids == sheet.cell("a2").value:
			if "IN" in sheet.cell("a%d" % (self.column_to_add("a")-1)).value:
				sheet["a%d"% self.column_to_add("a")]= "OUT -> " + self.clockin()
			else:
				sheet["a%d"% self.column_to_add("a")]= "IN -> " + self.clockin()
		elif self.ids == sheet.cell("b2").value:
			if "IN" in sheet.cell("b%d" % (self.column_to_add("b")-1)).value:
				sheet["b%d"% self.column_to_add("b")]= "OUT -> " + self.clockin()
			else:
				sheet["b%d"% self.column_to_add("b")]= "IN -> " + self.clockin()
			#sheet["b%d"% self. column_to_add("b")]= self.clockin()
		elif self.ids == sheet.cell("c2").value:
			sheet["c%d"% self. column_to_add("c")]= self.clockin()
		elif self.ids == sheet.cell("d2").value:
			sheet["d%d"% self. column_to_add("d")]= self.clockin()		
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
		
		self.name = x
		self.ids= y
		#while count <= 4:

		if sheet.cell('%s1'% get_column_letter(1)).value==None:
			sheet["%s1"% get_column_letter(1)]= self.name
			sheet["%s2"% get_column_letter(1)]= self.ids
		elif sheet.cell('%s1'% get_column_letter(2)).value==None and self.ids != sheet.cell("a2").value:
			sheet["%s1"% get_column_letter(2)]= self.name
			sheet["%s2"% get_column_letter(2)]= self.ids
		elif sheet.cell('%s1'% get_column_letter(3)).value==None and self.ids != sheet.cell("a2").value and self.ids != sheet.cell("b2").value:
			sheet["%s1"% get_column_letter(3)]= self.name
			sheet["%s2"% get_column_letter(3)]= self.ids
		elif sheet.cell('%s1'% get_column_letter(4)).value==None and self.ids != sheet.cell("a2").value and self.ids != sheet.cell("b2").value and self.ids != sheet.cell("c2").value:
			sheet["%s1"% get_column_letter(4)]= self.name
			sheet["%s2"% get_column_letter(4)]= self.ids
		wb.save("clocking.xlsx")
		
	
	def column_to_add(self,col): 
		wb = openpyxl.load_workbook("clocking.xlsx")
		sheet =wb.get_sheet_by_name('Sheet')
		sheet= wb.active
		sheet_max_row = sheet.max_row
		cell_coord = col + str(sheet_max_row)
		#print (sheet.cell(cell_coord).value)
		while sheet.cell(cell_coord).value == None:
			sheet_max_row -= 1
			cell_coord = col + str(sheet_max_row)
		sheet_max_row += 1
		return int(sheet_max_row)	

	#def grep(self,c,d):
'''
def main():
	
	first=timesheet()
	first.clockin()
	first.register()
	first.excel()
	

if __name__=='__main__':
	main()
'''
