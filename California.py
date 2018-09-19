# -*- coding: utf-8 -*-
"""
Created on Fri Jul 20 08:14:27 2018

@author: C252059
"""
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
import time
import os
from win32com.client import Dispatch
import pandas as pd
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from bs4 import BeautifulSoup
import gzip
import shutil
import zipfile
import pandas as pd
import itertools    
from bs4 import BeautifulSoup
import pyautogui as pgi
from selenium.webdriver.common.keys import Keys
import pprint
import gzip
import numpy as np
import xlsxwriter as xl




  
time_stuff = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx', sheet_name = 'Year-Qtr',use_cols='A:B')
yr = time_stuff.iloc[0,0]
qtr = time_stuff.iloc[0,1]
login_credentials = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='California', usecols='F,G',dtype='str')
username = login_credentials.iloc[0,0]
password = login_credentials.iloc[0,1]
chromeOptions = webdriver.ChromeOptions()
prefs = {'download.default_directory':'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder',
         'plugins.always_open_pdf_externally':True,
         'download.prompt_for_download':False}
chromeOptions.add_experimental_option('prefs',prefs)
os.chdir('O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder')
driver = webdriver.Chrome(chrome_options = chromeOptions, executable_path=r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\chromedriver.exe')
driver.get('https://www.medi-cal.ca.gov/')
transaction_tab = driver.find_element_by_xpath('//*[@id="nav_list"]/li[2]/a')
transaction_tab.click()
wait = WebDriverWait(driver,10)
user_name = wait.until(EC.element_to_be_clickable((By.ID,'UserID')))
user_name.send_keys(username)
pass_word = driver.find_element_by_id('UserPW')
pass_word.send_keys(password)
submit_button = driver.find_element_by_id('cmdSubmit')
submit_button.click()

#navigate to the drug rebate invoice page
drug_rebate = driver.find_element_by_xpath('//*[@id="tabpanel_1_sublist"]/li/a')
drug_rebate.click()

#get the three labeler codes.  Will have to update if labeler codes change
lilly_code = lambda : wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="middle_column"]/div[2]/table/tbody/tr/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[2]/td[2]/a')))
dista_code = lambda : wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="middle_column"]/div[2]/table/tbody/tr/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[3]/td[2]/a')))
imclone_code = lambda :wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="middle_column"]/div[2]/table/tbody/tr/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[4]/td[2]/a')))

codes = [lilly_code,dista_code,imclone_code]
'''
This block of code downloads all of the prepared reports.  The reports come in a .gz file
and have to be decompressed, this happens after the download in the next loop.
'''
                    
for code in codes:
    code().click()
    retrieve = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="middle_column"]/div[2]/table/tbody/tr/td[2]/table/tbody/tr[3]/td/a[2]/b')))
    retrieve.click()               
    wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="left_column"]/div[1]/a/img')))     
    soup2 = BeautifulSoup(driver.page_source,'html.parser') 

    bodies = soup2.find_all('tbody')
    body = bodies[2]
    rows = body.find_all('tr')
    data = body.find_all('td')
    data = [x.text for x in data]
    data = np.asarray(data)
    array_length = int(len(data)/3)
    data = data.reshape(-1,3)
    links = [x[0] for x in data if 'Completed' in x[1]]               
    links = ["".join(x.split()) for x in links]              
    for link in range(len(links)):
        xpath = "//a[contains(text(),'"+links[link]+"')]"
        DL_link = driver.find_element_by_xpath(xpath)
        DL_link.click()               
    driver.get(r'https://rais.medi-cal.ca.gov/drug/DrugLablr.asp')
    
'''
This is the loop that goes through the downloaded .gz files, unzips them, renames them
to the file format, makes them a text file, and then deletes the .gz file
'''
os.chdir('/Users/c252059/Downloads/')
files = os.listdir()
for file in files:
    new_name = file[:-6]+'txt'
    try:
        with gzip.open(file,'rt') as ref:
            content = ref.read()
            text_file = open(new_name,'w')
            text_file.write(content)
            text_file.close()
            ref.close()
    except:
        pass
    os.remove(file)
'''    
invoice_list = pd.read_excel(r'C:\Users\c252059\Documents\AutomationResources\Copy of CA Invoice vs System Description 7-19-18.xlsx',usecols=[0,3])
invoice_list['California Invoice Title'] = invoice_list['California Invoice Title'].str.lower()
compound_list = invoice_list[invoice_list['California Invoice Title'].str.contains('compound')]
regular_list = pd.concat([invoice_list,compound_list]).drop_duplicates(keep=False)
regular_columns = ['Claim Control Number','NDC Code','Date of Service (ccyymmdd)',
                   'Claim Adjudication Date (ccyymmdd)','Units of Service','Reimbursed Amount',
                   'Billed Amount','Adjustment Indicator','Prescription Number (Rx)','Billing Provider Number',
                   'Billing Provider Owner Number','Billing Provider Service Location Number',
                   'Adjustment Claim Control Number','Recipient Other Coverage Code','Other Health Coverage Indicator',
                   'TAR Control Number','Third Party Code','Third Party Amount','Patient Liability Amount',
                   'Co-Pay Code','Co-Pay Amount','Days Supply Number','Referring Prescribing Provider Number',
                   'Recipient Crossover Status Code','Recipient Prepaid Health Plan (PHP) Status Code',
                   'Compound Code','Cost Basis Determination Code']
compound_columns = ['Claim Control Number','NDC Code','Date of Service (ccyymmdd)',
                   'Claim Adjudication Date (ccyymmdd)','Units of Service','Reimbursed Amount',
                   'Billed Amount','Adjustment Indicator','Prescription Number (Rx)','Billing Provider Number',
                   'Billing Provider Owner Number','Billing Provider Service Location Number',
                   'Adjustment Claim Control Number','Recipient Other Coverage Code','Other Health Coverage Indicator',
                   'TAR Control Number','Third Party Code','Third Party Amount','Patient Liability Amount',
                   'Co-Pay Code','Co-Pay Amount','Days Supply Number','Referring Prescribing Provider Number',
                   'Recipient Crossover Status Code','Recipient Prepaid Health Plan (PHP) Status Code',
                   'Compound Code','Ingredient Cost Basis Determination Code','Claim Compound Ingredient Reimbrusement Amount']
path = os.getcwd()
files = os.listdir()
files = [x.split('_') for x in files] 
files = np.asarray(files)      
files = pd.DataFrame(files)                         
files['5'] = [files.iloc[i,4].split('.')[0] for i in range(len(files))]
for file in os.listdir():
    for program in files.loc[:,'5']:
        new_file = pd.read_table(file,sep='~',columns)
        if file.split('_')[4].split('.')[0] == program:
            print(file,program)
        else:
            continue
'''
    