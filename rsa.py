import uuid
import hashlib
from array import *

#Author: Gary Tsai

file = open('/Users/garytsai/Desktop/rfid-reader-http/summer_project/.AllEng.txt', 'r')
s=file.read()
k=s.split()

def passwd(x):
	for line in k:
		hash_object = hashlib.sha512(line)
		hex_dig = hash_object.hexdigest()
		if hex_dig == x:
			#print line
			#print hex_dig
			return line

