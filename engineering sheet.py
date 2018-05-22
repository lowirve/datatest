# -*- coding: utf-8 -*-
"""
Created on Wed May 16 08:36:41 2018

@author: XuBo
"""

from __future__ import division, print_function

import xlrd, os

import pandas as pd
import numpy as np

name = 'ENG UPDATE 05-18-18'

directory = r'N:\CRY\MfgEng\ENGINEERING SHEETS & NSWORFS'

directory = os.path.join(directory, name+'.xls')

cols = ['Quarter','Engineer', 'Part Number', 'Work Order Number', 'Product Type', 
       'Actual Ship Date', 'Ship Quantity', 'WIP Quantity', 'Coat ', 'Program', 
       'Org','Material','Size','Request Receive Date','Lead Time','Material Source']

data = pd.read_excel(directory, sheet_name=1, skiprows=[1])

data = data[cols]

data.columns = [x.strip() for x in cols]  

data.Quarter=data.Quarter.str[2:]+'Q'+data.Quarter.str[0]

coat_ref = {u'ENG STK': 'STOCK',
 u'ISE': 'ISE',
 u'ISE & Veeco': 'Veeco & ISE',
 u'ISE / VEECO 2': 'Veeco & ISE',
 u'ISE or ACG2': 'ACG or ISE',
 u'ORD': 'ORD',
 u'OSP': 'OSP',
 u'QTF': 'QTF',
 u'STOCK': 'STOCK',
 u'TBD': 'TBD',
 u'THOR': 'TITAN',
 u'THOR / ACG 1': 'TITAN & ACG',
 u'THOR / ACG 2': 'TITAN & ACG',
 u'TITAN': 'TITAN',
 u'Vecco 2': 'Veeco',
 u'Veeco': 'Veeco',
 u'Veeco ': 'Veeco',
 u'Veeco 1': 'Veeco',
 u'Veeco 1 or 2': 'Veeco',
 u'Veeco 2': 'Veeco',
 u'Veeco 2 / ISE': 'Veeco & ISE'}

data.Coat.replace(coat_ref, inplace=True)

type_ref = dict(zip([item for item in data['Product Type'].unique() if item is not np.nan], 
                     ['Gain' if 'Gain' in str(item) else str(item) for item in data['Product Type'].unique() if item is not np.nan]))

data.Coat.replace(type_ref, inplace=True)


#print(data.Quarter)

#xls = pd.ExcelFile(directory)

#print(xls.sheet_names)

#def coat(arr): # This is just a practice. It is definitely not efficient to solve the problem in this way.
#    ini = [item for item in arr if item is not np.nan]
#    
#    mid = [str(item).upper().replace('ENG STK', 'STOCK').split() for item in ini]
#    
#    sol = [ set(ii.rstrip('2').replace('THOR', 'TITAN').replace('VECCO', 'VEECO').replace('VEECO', 'Veeco') 
#            for ii in i if ii in ('THOR','ISE', 'VEECO', 'ACG', 'ACG2', 'VECCO', 
#            'TITAN')) if ('TITAN' in i or 'THOR' in i or 'ISE' in i or 'VEECO' 
#            in i or 'ACG' in i or 'VECCO' in i) else set(i) for i in mid]
#    
#    sol = [' & '.join(i) if ('TITAN' in i or 'ISE' in i or 'Veeco' 
#            in i or 'ACG' in i) else ' '.join(i).lstrip() for i in sol]
#    
#    return dict(zip(ini, sol))
