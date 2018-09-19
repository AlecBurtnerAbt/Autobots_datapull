# -*- coding: utf-8 -*-
"""
Created on Wed Sep 12 08:53:27 2018

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
from getReports import getReports

def download_reports():
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
    login_credentials = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Pennsylvania', usecols='A,B',dtype='str')
    username = login_credentials.iloc[0,0]
    password = login_credentials.iloc[0,1]
    yq=str(yr)+str(qtr)
    mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Pennsylvania', usecols='D,E',dtype='str')
    mapper = dict(zip(mapper['State Program'],mapper['Flex Program']))
    
    
    
    
    #Login with provided credentials
    driver.get('https://rsp.pagov.changehealthcare.com/RebateServicesPortal/login/home?goto=http://rsp.pagov.changehealthcare.com/RebateServicesPortal/')   
    
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
    reports_tab = driver.find_element_by_xpath('//a[text()="Reports"]')       
    reports_tab.click()                
    report = lambda: driver.find_element_by_xpath('//select[@id="reportList"]')
    report_select = lambda: Select(report())
    
    types = driver.find_element_by_xpath('//select[@id="docType"]')
    types_select = Select(types)
    programs = [x.text.replace(' ','_') for x in types_select.options]
    values = [x.get_attribute('value') for x in driver.find_elements_by_xpath('//select[@id="docType"]/option')]
    mapper = dict(zip(programs,values))    
    #Helper function to return boolean if report is ready
    def checker(element,xpath):
        try:
            EC.presence_of_element_located(element.find_element_by_xpath(xpath))
            return True
        except NoSuchElementException as ex:
            return False
    #Below is where the script finds the reports, downloads, and moves them
    rows = driver.find_elements_by_xpath('//table[@id="reportsResults"]/tbody/tr')
    rows = [row for row in rows if checker(row,'td//a//span[text()="Download Report"]')==True]
    
    #now that we have rows only for where reports are ready we can move forward
    names = [x.find_element_by_xpath('td[1]').text for x in rows]
    links = [x.find_element_by_xpath('td//a[@href="#"]') for x in rows]
    master_df = pd.DataFrame()
    
    for name, link in zip(names, links):
        #get info for file name
        ndc = name.split(' ')[7]
        state = name.split(' ')[8]
        program = name.split(' ')[10]
        value = mapper[program]
        first_half = '_'.join(name.split(' ')[:5])
        second_half = '-'.join(name.split(' ')[-4:]).replace(program,mapper[program])
        download_name = '-'.join([first_half,second_half])+'.xls'
        #download the file
        flag = 0
        while flag ==0:
            link.click()
            counter = 0
            while download_name not in os.listdir() and counter<21:
                time.sleep(1)
                counter+=1
            if download_name not in os.listdir():
                pass
            else:
                flag = 1
        temp_df = pd.read_excel(download_name,skipfooter=3)
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
        path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Claims\\Pennsylvania\\'+splitter+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
        file_name = 'PA_'+splitter+'_'+str(qtr)+'Q'+str(yr)+'.csv'
        if os.path.exists(path)==False:
            os.makedirs(path)
        else:
            pass
        os.chdir(path)
        frame.to_csv(file_name)
    driver.close()


if __name__=='__main__':
    download_reports()
    
    
    
    
    
    
    