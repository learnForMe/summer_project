
import smtplib
import time
from rsa import passwd 
from header import header
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


def email():
	hashed="43657166f4c72d25ef02dd2b82afb72b58860f1aeda068a45c2a7353962fb57ffa98db5231457318d6ffae8d6bcd56540f2fd871e3053486edd1e305c571af19"
	#passw= passwd(hashed)
	month="{:%B %Y}".format(datetime.date.today())
	fromaddr = "reportandlogs@gmail.com"
	toaddr = ['yue.tsai@jjay.cuny.edu']
	#toaddr = ['rpusateri@jjay.cuny.edu', 'yue.tsai@jjay.cuny.edu', 'holland718@hotmail.com','bryan.williams@jjay.cuny.edu']
	msg = MIMEMultipart()
	msg['From'] = fromaddr
	#msg['To'] = toaddr
	msg['To'] = ",".join(toaddr)
	msg['Subject'] = "Vet Lounge Traffic of %s" % month
	body = "DO NOT reply this email. This is an automatic generated email with traffic data for veterans lounge. Should you have any question, please email yue.tsai@jjay.cuny.edu."
	msg.attach(MIMEText(body, 'plain'))
	filename = "%s.xlsx" %month
	#filename = "August.xlsx" 
	attachment = open("/Users/johnjayveterans/Desktop/summer_project/testing.xlsx", "rb")
	#attachment = open("/Users/garytsai/Desktop/rfid-reader-http/summer_project/testing.xlsx", "rb")
	part = MIMEBase('application', 'octet-stream')
	part.set_payload((attachment).read())
	encoders.encode_base64(part)
	part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
	msg.attach(part)
	server = smtplib.SMTP('smtp.gmail.com', 587)
	server.starttls()
	server.login(fromaddr, "%s" % passwd(hashed))
	text = msg.as_string()
	server.sendmail(fromaddr, toaddr, text)
	server.quit()
