# -*- coding: utf-8 -*-
"""
Created on Wed Aug  8 14:19:42 2018

@author: C252059
"""

import pytesseract
import PIL
import os
import pandas as pd
import json
from pprint import pprint
import pdftotree 
import subprocess
import signal
import tabula
import pdfminer
tt.core.parse('test.pdf',html_path = 'C:\\Users\\c252059\\Desktop\\test.pdf', visualize=True)

pdfminer.

os.chdir(r'C:\Users\c252059\AppData\Local\Continuum\anaconda3\Lib\site-packages\tabula')

pages = list(range(6,36))
data = tabula.read_pdf('test.pdf',pages='all',multiple_tables=True,spreadsheet=True)

info = []
for table in data:
    for row in range(len(table)):
        if  any(list(map((lambda x:"humalog" in str(x).lower()),table.iloc[row,:])))==True:
            info.append(table.iloc[row,:])
        else:
            pass
            

















data = pd.read_json('test.json')
drugs = list((map(lambda x: str(x).lower(),data['Unnamed: 0'].unique())))
'trulicity' in drugs

a = open('test.json')
data2 = json.load(a)
pprint(data2)

a.close()
#Get all files to look at
a = os.walk('.')
forms = []
for root, dirs, files in os.walk('.'):
    for file in files:
        print(root+'\\'+file)
        forms.append(root+'\\'+file)
forms = [os.path.abspath(x) for x in forms]
#Loop through files to look and and see how many tables each one has
data_labels = dict.fromkeys(forms)        
for form in forms:  
    cmd = 'start "" "'+form+'"'
    a = subprocess.call(cmd, shell=True)
    col = input('How many columns does the pdf table have? ') 
    subprocess.call('TASKKILL /IM AcroRd32.exe')
    data_labels.update({form:col})
    

import pyautogui as pyg

os.chdir('/Users/c252059/Desktop/')
a = pyg.locateCenterOnScreen('RDR.png')
pyg.doubleClick(a)



















