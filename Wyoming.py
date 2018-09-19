# -*- coding: utf-8 -*-
"""
Created on Thu Aug 30 09:28:25 2018

@author: C252059
"""

# -*- coding: utf-8 -*-
"""
Created on Thu Aug 30 08:52:03 2018

@author: C252059
"""

# -*- coding: utf-8 -*-
"""
Created on Tue Aug 28 09:09:59 2018

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
from selenium.common.exceptions import NoSuchElementException, TimeoutException, NoAlertPresentException, InvalidElementStateException
from bs4 import BeautifulSoup
import gzip
import shutil
import zipfile
import pandas as pd
import itertools    
from bs4 import BeautifulSoup
from selenium.webdriver.common.keys import Keys
import pprint
import gzip
import numpy as np
from xlrd.biffh import XLRDError
import xlsxwriter as xl
import requests
from requests.auth import HTTPBasicAuth
os.chdir('C:/Users/')
chromeOptions = webdriver.ChromeOptions()
prefs = {'download.default_directory':'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder',
         'plugins.always_open_pdf_externally':True,
         'download.prompt_for_download':False}
chromeOptions.add_experimental_option('prefs',prefs)
driver = webdriver.Chrome(chrome_options = chromeOptions, executable_path=r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\chromedriver.exe')
os.chdir('O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder')
for file in os.listdir():
    os.remove(file)
time_stuff = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx', sheet_name = 'Year-Qtr',use_cols='A:B')
yr = time_stuff.iloc[0,0]
qtr = time_stuff.iloc[0,1]
login_credentials = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Wyoming', usecols='A,B',dtype='str')
username = login_credentials.iloc[0,0]
password = login_credentials.iloc[0,1]
yq=str(yr)+str(qtr)
mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Wyoming', usecols='D,E',dtype='str')
mapper = dict(zip(mapper['State Program'],mapper['Flex Program']))




#Login with provided credentials
driver.get('https://rsp.wygov.changehealthcare.com/RebateServicesPortal/login/home?goto=http://rsp.wygov.changehealthcare.com/RebateServicesPortal')   

#Now login

user = driver.find_element_by_xpath('//input[@id="username"]')
user.send_keys(username)
pass_word = driver.find_element_by_id('password')
pass_word.send_keys(password)
login = driver.find_element_by_id('submit')
login.click()

wait = WebDriverWait(driver,10)
wait2 = WebDriverWait(driver,2)
accept = wait.until(EC.element_to_be_clickable((By.ID,'terms')))
accept.click()

#invoice stuff is below this

invoices = driver.find_element_by_xpath('/html/body/div[3]/div/div[1]/div/div/ul/li[2]/a')
invoices.click()
code_dropdown = lambda: driver.find_element_by_id('labeler')
code_select = lambda: Select(code_dropdown())
codes = [x.text for x in code_select().options][1:]
type_dropdown = lambda: WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.ID,'docType')))
type_select = lambda: Select(type_dropdown())
master_dict = dict.fromkeys(codes)
time_stamp = lambda: driver.find_element_by_xpath('//input[@name="period"]')
values = driver.find_elements_by_xpath('//select[@id="docType"]//option')
values = [x.get_attribute('value') for x in values if int(x.get_attribute('value'))>1]

for code in codes:
    code_select().select_by_visible_text(code)
    print('Selecting '+str(code))
    ndcs = []
    for report,value in zip(list(mapper.keys()),values):
        
        type_select().select_by_visible_text(report)
        time.sleep(1)
        if ' ' in report:
            report = report.replace(' ','_')
        else:
            pass
        print('Selecting '+report)
        try:
            time_stamp().clear()
            time_stamp().send_keys(yq)
        except InvalidElementStateException as ex:
            pass
        submit_button = driver.find_element_by_xpath('//input[@type="submit"]')
        print('Requesting file.')
        submit_button.click()
        wait.until(EC.staleness_of(submit_button))
        success=0
        while success ==0:
            try:
                error = wait2.until(EC.presence_of_element_located((By.XPATH,'//li[contains(text(),"An error")]')))
                print('Website error! Moving back')
                driver.back()
                type_select().select_by_visible_text(report)
                submit_button = driver.find_element_by_xpath('//input[@type="submit"]')
                print('Requesting file.')
                submit_button.click()
                
            except TimeoutException as ex:
                success=1
        print('Files returned.')
        wait.until(EC.element_to_be_clickable((By.XPATH,'//a[@title="Download"]')))
        links = lambda: driver.find_elements_by_xpath('//a[@title="Download"]')
        for i,link in enumerate(links()):
            success_flag =0
            if i ==0:
                file_type = '.txt'
            else:
                file_type = '.pdf'
            file = 'WY-'+code+'-'+yq+'-'+value+file_type
            print('Downloading file '+str(i+1))
            reset_counter=0
            while success_flag ==0:              
                links()[i].click()
                try:
                    alert = driver.switch_to.alert
                    alert.accept()
                except NoAlertPresentException as ex:
                    pass
                try:
                    error = wait2.until(EC.presence_of_element_located((By.XPATH,'//li[contains(text(),"An error")]')))
                    print('Website error! Moving back')
                    driver.back()
                    reset_counter+=1
                    time.sleep(reset_counter*1.5)
                    reset_flag = 1
                except TimeoutException as ex:
                    reset_flag=0
                    pass
                counter = 0
                if reset_flag ==1:
                    pass
                else:
                    while file not in os.listdir() and counter<10:
                        time.sleep(1)
                        counter+=1
                    if file in os.listdir():
                        success_flag=1
                    else:
                        pass
            if file_type == '.txt':
                read_success=0
                while read_success==0:
                    try:
                        with open(file) as ax:
            
                            lines = ax.readlines()
                            xxx = list(set([x[6:17] for x in lines]))
                            [ndcs.append(x) for x in xxx]
                            read_success=1
                    except PermissionError as ex:
                        pass
                        
            path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Invoices\\Wyoming\\'+report+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
            if os.path.exists(path)==False:
                os.makedirs(path)
            else:
                pass
            file_name = 'WY_'+report+'_'+code+'_'+str(qtr)+'Q'+str(yr)+file_type
            shutil.move(file,path+file_name)
    master_dict.update({code:ndcs})

     #############################################CLD below

reports_tab = driver.find_element_by_xpath('//a[text()="Reports"]')       
reports_tab.click()                
report = lambda: driver.find_element_by_xpath('//select[@id="reportList"]')
report_select = lambda: Select(report())
not_ready = []

for labeler in list(master_dict.keys()):
    master_df = pd.DataFrame()
    print('a')
    for rep in [x.text for x in report_select().options][1:]:
        report_select().select_by_visible_text(rep)
        types =lambda: wait.until(EC.element_to_be_clickable((By.XPATH,'//select[@id="docType"]')))
        types_select = lambda: Select(reports())    
        Ts = [x.text for x in types_select().options[1:]]
        values = [x.get_attribute('value') for x in types_select().options if int(x.get_attribute('value'))>1]
        for T,value in zip(Ts,values):
            for ndc in master_dict[labeler]:
                print('b')
                if len(ndc)<2:
                    continue
                else:
                    pass
                
                print('c')
                ndc_in = wait.until(EC.presence_of_element_located((By.XPATH,'//input[@name="ndc"]')))
                ndc_in.send_keys(ndc)
                time.sleep(1)
   
                print('d')

            
                types_select().select_by_visible_text(T)
                print('e')
                time_stamp = driver.find_element_by_xpath('//input[@name="rpuStart"]')
                time_stamp.send_keys(yq)
                submit_button= driver.find_element_by_xpath('//input[@value="Submit"]')
                submit_button.click()
                accept = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@value="Accept"]')))
                accept.click()
                print('f')
                wait.until(EC.staleness_of(accept))
                success = 0
                while success==0:
                    try:
                        print('g')
                        download = wait.until(EC.element_to_be_clickable((By.XPATH,'//tr[contains(text(),"'+ndc+'"]//a[@title="Download"]')))
                        download.click()
                        if T=='JCode':
                            name = 'EXT_JCODE_CLD-'+ndc+'-WY-'+yq+'-'+value+'.xls'
                        else:    
                            name = 'EXT_Claim_Level_Detail_Report-'+ndc+'-WY'+yq+'-'+value+'.xls'
                                
                        while any(map((lambda x: 'EXT_Claim_Level' in x),os.listdir()))==False:
                            time.sleep(1)
                        file = os.listdir()[0]
                        stats = os.stat(file)
                        if stats.st_size >5:
                            print('h')
                            temp = pd.read_excel(file)
                        else:
                            print('i')
                            continue
                        print('j')
                        master_df = master_df.append(temp)
                        os.remove(file)
                        remove_download = driver.find_element_by_xpath('//a[@title="Delete"]')
                        remove_download.click()
                        alert = driver.switch_to.alert
                        alert.accept()
                        success=1
                    except TimeoutException as ex:
                        try:
                            error = driver.find_element_by_xpath('//li[text()="An error has occurred"]')
                            driver.get('https://rsp.wygov.changehealthcare.com/RebateServicesPortal/reports/index')
                            report_select().select_by_visible_text(rep)
                            types_select().select_by_visible_text(T)
                            ndc_in = wait.until(EC.presence_of_element_located((By.XPATH,'//input[@name="ndc"]')))
                            ndc_in.send_keys(ndc)
                            submit_button= driver.find_element_by_xpath('//input[@value="Submit"]')
                            submit_button.click()
                            accept = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@value="Accept"]')))
                            accept.click()
                        except NoSuchElementException as ex:
                            print('j')
                            not_ready.append(ndc)
                            driver.refresh()
                            print('CLD for '+str(ndc)+' is not ready, it will be ready tomorrow.')
                            report_select().select_by_visible_text(rep)
                            success=1
    path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Claims\\Wyoming\\'+lly_prog+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
    file_name = 'WY_'+labeler+'_'+rep+'_'+str(qtr)+'Q'+str(yr)+'.csv'
    if os.path.exists(path)==False:
        os.makedirs(path)
    else:
        pass
    master_df.to_csv(path+file_name,index=False)
    
subject = 'Wyoming CLD'
body = 'The following CLD were not ready and can be downloaded tomorrow:'
body2 = ''.join(['NDC: %s \n' %(x) for x in not_ready])
recipient = 'b2b_cma_llymedicaid@lilly.com'
base = 0x0
obj = Dispatch('Outlook.Application')
newMail = obj.CreateItem(base)
newMail.Subject = subject
newMail.Body = body+'\n'+body2
newMail.To = recipient
newMail.display()
newMail.Send()
        

import mechanicalsoup



