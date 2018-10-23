# -*- coding: utf-8 -*-
"""
Created on Thu Oct 18 13:34:11 2018

@author: C252059
"""

import os
import win32com.client as com


path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder'
os.chdir(path)
master_df = pd.DataFrame()
for root, dirs, files in os.walk(path):
    for file in files:
        temp_df = pd.read_excel(file,skipfooter=3)
        temp_df = temp_df.dropna(axis=0,how='all')
        if len(temp_df)==0:
            continue
        else:
            pass
        temp_df['NDC']= ndc
        temp_df['Program'] = program
        master_df = master_df.append(temp_df)
frames = []
splitters = master_df.Program.unique().tolist()  
for splitter in splitters:
    frame = master_df[master_df['Program']==splitter]
    path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Claims\\Vermont\\'+splitter+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
    file_name = 'VT_'+splitter+'_'+str(qtr)+'Q'+str(yr)+'.csv'
    if os.path.exists(path)==False:
        os.makedirs(path)
    else:
        pass
    os.chdir(path)
    frame.to_excel(file_name, engine='xlsxwriter')