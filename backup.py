import datetime
import shutil, os
from datetime import timedelta
from time import gmtime, strftime
import time, threading

def back_up():
	interval = strftime("%M:%S")
	time = strftime("%Y-%m-%d %H:%M:%S")
	shutil.copy('/Users/johnjayveterans/Desktop/summer_project/testing.xlsx', '/Users/johnjayveterans/Desktop/summer_project/back_up')
	shutil.move('/Users/johnjayveterans/Desktop/summer_project/back_up/testing.xlsx', '/Users/johnjayveterans/Desktop/summer_project/back_up/testing%s.xlsx'% time)
	#print (interval)
	threading.Timer(60*120, back_up).start()

#back_up()