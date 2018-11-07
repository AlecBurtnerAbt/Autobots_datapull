# -*- coding: utf-8 -*-
"""
Created on Tue Oct  9 13:52:57 2018

@author: C252059
"""

import pandas as pd
import os 
import re
import numpy as np

cali_columns =['Claim Control Number','NDC Code','Date of Service','Claim Adjudication Date','Units of Service',
          'Reimbursed Amount','Billed Amount','Adjustment Indicator','RX_ID','Billing Provider Number','Billing Provider Owner Number',
          'Billing Provider Service Location Number','Adjustment Claim Control Number','Recipient Other Coverage Code',
          'Other Health Coverage Indicator','TAR Control Number','Third Party Code','Third Party Amount','Patient Liability Amount',
          'Co-Pay Code','Co-Pay Amount','Days Supply Number','Referring Prescribing Provider Number','Recipient Crossover Status Code',
          'Recipient Health Plan Status Code','Compound Code','Cost Basis Determination Code']


cali_compound_columns =['Claim Control Number','NDC Code','Date of Service','Claim Adjudication Date','Units of Service',
          'Reimbursed Amount','Billed Amount','Adjustment Indicator','RX_ID','Billing Provider Number','Billing Provider Owner Number',
          'Billing Provider Service Location Number','Adjustment Claim Control Number','Recipient Other Coverage Code',
          'Other Health Coverage Indicator','TAR Control Number','Third Party Code','Third Party Amount','Patient Liability Amount',
          'Co-Pay Code','Co-Pay Amount','Days Supply Number','Referring Prescribing Provider Number','Recipient Crossover Status Code',
          'Recipient Health Plan Status Code','Compound Code','Ingredient Cost Basis Determination Code', 'Claim Compound Ingredient Reimbursement Amount']

cajun_columns = ['NDC','Quantity','Payment Date','Date of Service','Prescriber Number','Billed Charge','Paid Amount','Provider Number',
                 'Rx Days Supply','Dispensing Fee','Non-Medicaid Paid Amount','Co-Pay','Refill Code','ICN','ICN Line','Former ICN',
                 'Former ICN Status','Claim Status','Claim Load Date']




class ExcelMaker():
    def __init__(self,state,year,qtr,columns,compound_columns,st_abbrev):
        self.state = state
        self.year = year
        self.qtr = qtr
        self.columns = columns
        self.compound_columns = compound_columns
        self.path = f'O:\\M-R\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Claims\\{state}\\'
        self.excel_path = f'O:\\M-R\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Excels\\{self.state}\\'
        self.raw_text_path = f'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Raw Text\\{self.state}\\'
        self.magic_folder = 'Z:\\'
        self.st_abbrev = st_abbrev
        
    def make_cali_excels(self):
        self.dirs = [dirs for roots, dirs, files in os.walk(self.path)]
        self.programs = [program for program in self.dirs[0]]
        for program in self.programs:
           os.chdir(f'O:\\M-R\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Claims\\{self.state}\\{program}\\{self.year}\\Q{self.qtr}\\')
           if 'CMPD' in program:
               program_df = pd.DataFrame(columns= self.compound_columns)
           else:
               program_df = pd.DataFrame(columns = self.columns)
           for file in os.listdir():
              temp = pd.read_table(file,sep='~',dtype=str,names=program_df.columns,encoding='latin1')
              program_df = program_df.append(temp)
           os.chdir(self.excel_path)
           program_df.to_excel(f'{self.state} {program} for Visualization.xlsx')
           special_name = f'CA_{program}_{self.qtr}Q{self.year}.xlsx'
           shutil.copy(f'{self.state} {program} for Visualization.xlsx',self.magic_folder+'\\'+special_name)
            
    def make_cajun_excels(self):
       os.chdir(self.raw_text_path)
       file_name = f'{self.state} for Visualization.xlsx'
       self.data = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Test\Raw Text\Louisiana\FileFormat.xlsx',usecols='A,C')
       self.data = self.data.dropna(how='all')
       self.data['Position'] = self.data.Position.str.split()
       self.data['start'] = self.data.Position.str[0]
       self.data['end'] = self.data.Position.str[2]
       self.data['start'] = self.data.start.apply(int)
       self.data['end'] = self.data.end.apply(int)
       self.data['start'] = self.data.start.apply((lambda x: x-1))
       self.data['end'] = self.data.end.apply((lambda x: x-1))
       self.data_cuts = [(start,end) for start,end in zip(self.data.start,self.data.end)]
       self.column_names = self.data.Field
       files = [file for file in os.listdir() if '.txt' in file.lower()]
       data = []
       for file in files:    
           with open(file) as f:
               lines = f.readlines()
               
               for line in lines:
                    holder = []
                    for start,end in self.data_cuts:
                        holder.append(line[start:end])
                    data.append(holder)
       cajun_frame = pd.DataFrame(data,columns=self.column_names)

       if os.path.exists(self.excel_path)==False:
           os.makedirs(self.excel_path)
       os.chdir(self.excel_path)
       cajun_frame.to_excel(file_name,index=False)
       
    def mexontana_excels(self):
        self.data = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Test\Raw Text\Montana New Mexico-Conduent Text CLD\DRAMS NCPDP Format2.xls',skiprows=2)
        self.column_names = self.data.iloc[:,0]
        self.data_cuts = [(int(start)-1,int(end)-1) for start,end in zip(self.data.Start,self.data.End)]
        os.chdir(self.raw_text_path)
        files = [file for file in os.listdir() if '.txt' in file]
        data = []
        for file in files:
            with open(file) as F:
                lines = F.readlines()[1:]
                for line in lines:
                    holder = []
                    for start,end in self.data_cuts:
                        holder.append(line[start:end])
                    data.append(holder)
        mexontana_frame = pd.DataFrame(data, columns=self.column_names)
        if os.path.exists(self.excel_path)==False:
            os.makedirs(self.excel_path)
        file_name = f'{self.state} for Visualization.xlsx'
        os.chdir(self.excel_path)
        mexontana_frame.to_excel(file_name,index=False)
        
    def mexontana_for_submission(self):
        self.data = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Test\Raw Text\Montana New Mexico-Conduent Text CLD\DRAMS NCPDP Format2.xls',skiprows=2)
        self.column_names = self.data.iloc[:,0]
        self.data_cuts = [(int(start)-1,int(end)-1) for start,end in zip(self.data.Start,self.data.End)]
        os.chdir(self.raw_text_path)
        files = [file for file in os.listdir() if '.txt' in file]
        for file in files:
            os.chdir(self.raw_text_path)
            program = file.split('_')[1]
            label_code = file.split('_')[3]
            data = []
            with open(file) as F:
                lines = F.readlines()[1:-1]
                for line in lines:
                    holder = []
                    for start, end in self.data_cuts:
                        holder.append(line[start:end])
                    data.append(holder)
            frame = pd.DataFrame(data,columns=self.column_names)
            file_name = f'{self.st_abbrev}_{program}_{self.qtr}Q{self.year}_{label_code}.xlsx'
            path =f'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Converted Raw Text\\Claims\\{self.state}\\{program}\\'
            if os.path.exists(path)==False:
                os.makedirs(path)
            os.chdir(path)
            frame.to_excel(path+file_name,index=False)
        

               
cali = ExcelMaker(state = 'California', year = 2018, qtr = 2, columns = cali_columns, compound_columns = cali_compound_columns, st_abbrev='CA')        
cali.make_cali_excels()    


crawdads = ExcelMaker(state = 'Louisiana',year = 2018, qtr=2, columns = cajun_columns, compound_columns=None, st_abbrev = 'LA')
crawdads.make_cajun_excels()


montana = ExcelMaker(state='Montana Conduet', year = 2018, qtr = 2, columns = None, compound_columns = None, st_abbrev = 'MT')
montana.mexontana_excels()
montana.mexontana_for_submission()

new_mexico = ExcelMaker(state='New Mexico Conduet', year = 2018, qtr = 2, columns = None, compound_columns = None, st_abbrev = 'NM')
new_mexico.mexontana_excels()
new_mexico.mexontana_for_submission()
