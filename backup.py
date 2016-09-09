import datetime
import shutil, os
from time import strftime

def back_up():
	
	time = strftime("%Y-%m-%d %H:%M:%S")
	shutil.copy('/Users/johnjayveterans/Desktop/summer_project/testing.xlsx', '/Users/johnjayveterans/Desktop/summer_project/back_up')
	shutil.move('/Users/johnjayveterans/Desktop/summer_project/back_up/testing.xlsx', '/Users/johnjayveterans/Desktop/summer_project/back_up/testing%s.xlsx'% time)
	#shutil.copy('/Users/garytsai/Desktop/rfid-reader-http/summer_project/testing.xlsx', '/Users/garytsai/Desktop/rfid-reader-http/summer_project/back_up')
	#shutil.move('/Users/garytsai/Desktop/rfid-reader-http/summer_project/back_up/testing.xlsx', '/Users/garytsai/Desktop/rfid-reader-http/summer_project/back_up/testing%s.xlsx'% time)
	