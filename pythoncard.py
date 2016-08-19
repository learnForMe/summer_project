"""Smartcard CardRequest(Modified by Gary Tsai).

__author__ = Gary Tsai and Freddy Velez
For Veterans Lounge of John Jay College

__author__ = "http://www.gemalto.com"

Copyright 2001-2012 gemalto
Author: Jean-Daniel Aussel, mailto:jean-daniel.aussel@gemalto.com

This file is part of pyscard.

pyscard is free software; you can redistribute it and/or modify
it under the terms of the GNU Lesser General Public License as published by
the Free Software Foundation; either version 2.1 of the License, or
(at your option) any later version.

pyscard is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU Lesser General Public License for more details.

You should have received a copy of the GNU Lesser General Public License
along with pyscard; if not, write to the Free Software
Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA
"""

from __future__ import print_function
from smartcard.CardType import ATRCardType
from smartcard.pcsc.PCSCCardRequest import PCSCCardRequest
from smartcard.util import toHexString , toBytes
import openpyxl
from openpyxl import load_workbook
from openpyxl.compat import range
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils import coordinate_from_string
from openpyxl.styles import Font
from monthly_stat_row import add_month
from formular import formular
#from header import header
from art import art_schedule
import alert
import re
import os
import time
import smtplib



def remove(y):
    
    cell_location=re.sub(r'[\D]','',y)
    return cell_location

def search_Student (x):
    count =0
    first_time =1
    for row in sheet.iter_rows():
        for cell in row:
            data = cell.value
            
            if data == str(x):
                cell = str(cell)
                cool=int(remove(cell))
                stop_by= ws.cell('%s%d' % (col3,cool)).value
                #print (stop_by)
                
                ws['%s%d' % (col3,cool)] = stop_by+1
                wb.save('testing.xlsx')
                count+=1
              
    if count< 1:
        print ('\a\a\a\a\a\a')
        os.system ('clear')
        #os.system ('echo "New Student"')
        print ("NEW STUDENT\n")
        #print ("This Student is NEW")
        student_name= input("Enter Name -> ")
        print (student_name,"added to database")
        
        ws['%s%d' % (col2,insert_name)] =str(student_name)
        ws['%s%d' % (col,insert_name)] =str(x)
        ws['%s%d' % (col3,insert_name)] =first_time
        wb.save('testing.xlsx')
    else:
        print ('\a\a')
        again=ws.cell('%s%d' % (col2,cool)).value
        print ("Welcome Back!", again)
        #os.system ('say What is up%s' % again)  # use for prank (aka April Foo)

  


class CardRequest(object):
    """A CardRequest is used for waitForCard() invocations and specifies what
    kind of smart card an application is waited for.
    """

    def __init__(self, newcardonly=False, readers=None, cardType=None,
        cardServiceClass=None, timeout=1):
        """Construct new CardRequest.

        newcardonly:        if True, request a new card
                            default is False, i.e. accepts cards already
                            inserted

        readers:            the list of readers to consider for
                            requesting a card default is to consider all
                            readers

        cardType:           the smartcard.CardType.CardType to wait for;
                            default is smartcard.CardType.AnyCardType,
                            i.e. the request will succeed with any card

        cardServiceClass:   the specific card service class to create
                            and bind to the card default is to create
                            and bind a smartcard.PassThruCardService

        timeout:            the time in seconds we are ready to wait for
                            connecting to the requested card.  default
                            is to wait one second to wait forever, set
                            timeout to None
        """
        self.pcsccardrequest = PCSCCardRequest(newcardonly, readers,
            cardType, cardServiceClass, timeout)

    def getReaders(self):
        """Returns the list or readers on which to wait for cards."""
        return self.pcsccardrequest.getReaders()

    def waitforcard(self):
        """Wait for card insertion and returns a card service."""
        return self.pcsccardrequest.waitforcard()

    def waitforcardevent(self):
        """Wait for card insertion or removal."""
        return self.pcsccardrequest.waitforcardevent()
wb=load_workbook('testing.xlsx', data_only = True)
ws=wb.active
worksheet= wb.get_sheet_names()
sheet = wb.get_sheet_by_name('Sheet')
col=get_column_letter(1)# convert column number to letter and use for first column (ID card data)
col2=get_column_letter(2)# use for second column (Student name)
col3=get_column_letter(3)#use for third column(occurance)
os.system ("python formular.py")
while __name__ == '__main__':
    """Small sample illustrating the use of CardRequest.py."""

    cardtype = "3B 8F 80 01 80 4F 0C A0 00 00 03 06 40 00 00 00 00 00 00 28"
    print('\t''------Tap card to SIGN IN-------')

   # cr = CardRequest(timeout=10, cardType=cardtype)
    cr = CardRequest(timeout=None, newcardonly=True)
    cs = cr.waitforcard()
    cs.connection.connect()
    card = toHexString(cs.connection.getATR())
    SELECT = [0xFF, 0xCA, 0x00, 0x00, 0x00]

    response, sw1, sw2 = cs.connection.transmit( SELECT)
    
   # print ('response: ', response, ' status words: ', "%x %x" % (sw1, sw2))
    texting = toHexString(response).replace(' ','')
    if card == cardtype:
        
        max_col = sheet.max_column
        formular_col= get_column_letter(max_col)
        max_row = sheet.max_row
        insert_name=max_row+1
        
        search_Student(texting)
        #formular()
        #wb.save('testing.xlsx')
        formular()
        add_month()
        cs.connection.disconnect()
    
    else:
        os.system ('clear')
        print (alert.stop)
        os.system ('say Try Again')
        cs.connection.disconnect()
    
   
