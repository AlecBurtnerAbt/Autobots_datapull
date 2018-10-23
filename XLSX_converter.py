# -*- coding: utf-8 -*-
"""
Created on Thu Oct 18 13:34:11 2018

@author: C252059
"""

import os
import win32com.client as com
path = r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Test\Claims'

for root, dirs, files in os.walk(path):
    for file in files:
        if file.split('.')[-1]=='xls':
            excel = com.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(root+'\\'+file)
            wb.SaveAs(root+'\\'+file+'x', FileFormat=51)
            wb.Close()
path2 = r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Test\Archive\XLS'       
for root, dirs, files in os.walk(path):
    for file in files:
        if file.split('.')[-1]=='xls':
            shutil.move(root+'\\'+file,path2+'\\'+file)
   

import pandas as pd  
submitted = []       
for root, dirs, files in os.walk(path):
    for file in files:
        submitted.append(file)
submitted = pd.DataFrame(submitted,columns=['Files'])
submitted.to_excel('submitted_files.xlsx',engine='xlsxwriter',index=False)
        