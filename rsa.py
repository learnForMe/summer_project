import uuid
import hashlib
from array import *

#Author: Gary Tsai

#file = open('/Users/johnjayveterans/Downloads/pyscard-1.9.4/sudoEng.txt', 'r')
file = open('/Users/garytsai/Desktop/Fall 2015/Crypto/Portfolio/AllEng.txt', 'r')
s=file.read()
k=s.split()

def passwd(x):
	for line in k:
		hash_object = hashlib.sha512(line.encode('utf-8'))
		hex_dig = hash_object.hexdigest()
		if hex_dig == x:
			#print line
			#print hex_dig
			return line

