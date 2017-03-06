import datetime
import shutil, os
from time import strftime
import time, threading
import calendar
from autoEmail import email
from backup import back_up
from add_column import add_column
from incre_col import increase_col


def exe():
	time_to_exe= "22:00:00"
	backup1="10:00:00"
	backup2="12:00:00"
	backup3="15:00:00"
	backup4="18:00:00"
	backup5="20:00:00"
	#now = strftime("%Y-%m-%d %H:%M:%S")
	while 1:

		end_this_mon=calendar.mdays[datetime.date.today().month]
		today = "{:%d}".format(datetime.datetime.today())
		now = strftime("%H:%M:%S")
		time.sleep(1)
		if now == backup1 or now == backup2 or now == backup3 or now == backup4 or now == backup5:
			back_up()
		if today == end_this_mon and now == time_to_exe:
			email()
			add_column()
			increase_col()
			