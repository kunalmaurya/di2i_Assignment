#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Aug 18 01:43:41 2018

@author: bluedanube
"""

import xlrd
import time
from datetime import datetime

time.ctime(int("42948"))

py_date = xlrd.xldate.xldate_as_datetime(42948, book.datemode)

import pandas as pd

def convert_excel_time(excel_time):
    '''
    converts excel float format to pandas datetime object
    round to '1min' with 
    .dt.round('1min') to correct floating point conversion innaccuracy
    '''
    
    return pd.to_datetime('1899-12-30') + pd.to_timedelta(excel_time,'D')

book = open_workbook('Prices.xlsx',on_demand=True)
price=[]
date=[]
    
for name in book.sheet_names():

    sheet = book.sheet_by_name(name)

# Attempt to find a matching row (search the first column for 'john')
    rowIndex = -1
    
    for cell in sheet.row(0): #
        if(cell.value !='Months'):
            print(convert_excel_time(cell.value))
            date.append(convert_excel_time(cell.value))
            
    for i in range(1,len(sheet.row(1))):
        print(sheet.cell(1,i).value)
        price.append(sheet.cell(1,i).value)
        
df_price= pd.DataFrame()                             
df_price['date']= date
df_price['price']=price  

 
