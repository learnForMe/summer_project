"""Smartcard CardRequest.

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
import re
import os
import sys
def remove_space(z):
    no_space=re.sub(r'[\s+]',"",z)
    return no_space



def remove(y):
    
    cell_location=re.sub(r'[\D]','',y)
    #cell_location=re.sub(r'>','',cell_location)
    #cell_location=re.sub(r'\D','',cell_location)

    #cell_location=re.sub(r'[< Cell Sheet1]','',cell_location)
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
               # print (cool)
                
                ws['%s%d' % (col3,cool)] = stop_by+1
                wb.save('testing.xlsx')
                count+=1
              
    if count< 1:
        sys.stdout.write('\a\a\a\a\a\a')
        sys.stdout.flush()
        os.system ('clear')
        os.system ('echo "New Student"')
        #print ("This Student is NEW")
        student_name= raw_input("Enter Name -> ")
        print (student_name," added to database")
        
        ws['%s%d' % (col2,insert_name)] =str(student_name)
        ws['%s%d' % (col,insert_name)] =str(x)
        ws['%s%d' % (col3,insert_name)] =first_time
        wb.save('testing.xlsx')
    else:
       
        again=ws.cell('%s%d' % (col2,cool)).value
        print ("Welcome Back!", again)

        return data

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


while __name__ == '__main__':
    """Small sample illustrating the use of CardRequest.py."""

    #cardtype = ATRCardType( toBytes("3b 8f 80 01 80 4f 0c a0 00 00 03 06 40 00 00 00 00 00 00 28"))
    print('\t''------insert card to SIGN IN-------')
   # cr = CardRequest(timeout=10, cardType=cardtype)
    cr = CardRequest(timeout=None, newcardonly=True)
    cs = cr.waitforcard()
    cs.connection.connect()

    SELECT = [0xFF, 0xCA, 0x00, 0x00, 0x00]
    apdu = SELECT

    response, sw1, sw2 = cs.connection.transmit( apdu )
    
   # print ('response: ', response, ' status words: ', "%x %x" % (sw1, sw2))
    tagid = toHexString(response).replace(' ','')
    
    id = tagid
    texting = id

   # print(" UID is",id)


    wb=load_workbook('testing.xlsx',read_only = False, data_only = True)
    sheet = wb.get_sheet_by_name('Sheet')
    ws=wb.active

    col=get_column_letter(1)# convert column number to letter and use for first column (ID card data)
    col2=get_column_letter(2)# use for second column (Student name)
    col3=get_column_letter(3)#use for third column(occurance)

    max_row = sheet.max_row
    insert_name=max_row+1
    
    search_Student(texting)         
    #wb.save('testing.xlsx')
    cs.connection.disconnect()
   
