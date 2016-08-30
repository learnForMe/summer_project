import datetime
import shutil, os
from datetime import timedelta
from time import gmtime, strftime
import time, threading

def back_up():
	interval = strftime("%M:%S")
	time = strftime("%Y-%m-%d %H:%M:%S")
	shutil.copy('/Users/garytsai/Desktop/rfid-reader-http/summer_project/hi.txt', '/Users/garytsai/Desktop/rfid-reader-http/summer_project/back_up')
	shutil.move('/Users/garytsai/Desktop/rfid-reader-http/summer_project/back_up/hi.txt', '/Users/garytsai/Desktop/rfid-reader-http/summer_project/back_up/hi.txt'% time)
	#print (interval)
	threading.Timer(60*60, back_up).start()

#back_up()