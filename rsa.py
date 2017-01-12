import uuid
import hashlib
from array import *

#Author: Gary Tsai

with open('/Users/johnjayveterans/Desktop/pyscard-1.9.4/sudoEng.txt', 'r') as sha:
#with open('/Users/garytsai/Desktop/Fall 2015/Crypto/Portfolio/AllEng.txt', 'r') as sha:
	hashed=sha.read()
	hashed=hashed.split()

def passwd(x):
	for line in hashed:
		hash_object = hashlib.sha512(line.encode('utf-8'))
		hex_dig = hash_object.hexdigest()
		#print (hex_dig, line)
		if hex_dig == x:
			#print (line)
			#print (hex_dig)
			return line

sha.closed

