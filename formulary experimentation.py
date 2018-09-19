# -*- coding: utf-8 -*-
"""
Created on Wed Aug  8 14:19:42 2018

@author: C252059
"""

import pytesseract
import PIL
import tabula
import os
import pandas as pd
import json
from pprint import pprint
import pdftotree as tt
import subprocess
import signal
import poppler


tt.parse('test.pdf',html_path = 'C:\\Users\\c252059\\Desktop\\', visualize=True)


os.chdir('/Users/c252059/Documents/AutomationResources/Formularies/')


pdf = tabula.convert_into('test.pdf','test.csv',pages='all',output_format='csv')
print(pdf)

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
    
























