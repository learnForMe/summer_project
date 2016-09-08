import datetime
import time
import calendar
import shutil, os
from add_column import add_column
from autoEmail import email
from backup import back_up
def check_date():
	end_this_mon=calendar.mdays[datetime.date.today().month]
	today = "{:%d}".format(datetime.datetime.today())
	now= time.strftime("%H:%M:%S")
	time_to_add= "20:30:00"
	if today == end_this_mon and now >= time_to_add:
		add_column()
		email()



