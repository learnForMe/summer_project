import datetime 
import alert

def art_schedule():
	start = datetime.datetime.strptime("05-08-2016", "%d-%m-%Y")
	end = datetime.datetime.strptime("05-08-2016", "%d-%m-%Y")
	date_generated = [start + datetime.timedelta(days=x) for x in range(0, (end-start).days)]

	for date in date_generated:
		print (alert.christmas)
	
	start = datetime.datetime.strptime("05-08-2016", "%d-%m-%Y")
	end = datetime.datetime.strptime("05-08-2016", "%d-%m-%Y")
	date_generated = [start + datetime.timedelta(days=x) for x in range(0, (end-start).days)]

	for date in date_generated:
		print (alert.stop)



	start = datetime.datetime.strptime("04-08-2016", "%d-%m-%Y")
	end = datetime.datetime.strptime("05-08-2016", "%d-%m-%Y")
	date_generated = [start + datetime.timedelta(days=x) for x in range(0, (end-start).days)]

	for date in date_generated:
		print (alert.johnjay)
