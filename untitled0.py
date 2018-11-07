# -*- coding: utf-8 -*-
"""
Created on Mon Nov  5 15:55:47 2018

@author: C252059
"""

import camelot 
import pandas as pd
import os

os.chdir(r'C:\Users\c252059\Documents\Formulary Automation\Test Documents')

files = os.listdir()

data = camelot.read_pdf(files[2], pages='all', flavor='stream')
data._tables
b = pd.DataFrame()
for frame in data._tables:
    b = a.append(frame.df)
