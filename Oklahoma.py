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
from selenium.common.exceptions import NoSuchElementException, TimeoutException, NoAlertPresentException
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
import xlsxwriter as xl
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
login_credentials = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Oklahoma', usecols='A,B',dtype='str')
username = login_credentials.iloc[0,0]
password = login_credentials.iloc[0,1]
#Login with provided credentials
driver.get('https://www.edsdocumenttransfer.com/')   
wait = WebDriverWait(driver,10)
user = driver.find_element_by_xpath('//input[@name="Username"]')
user.send_keys(username)
pw = driver.find_element_by_xpath('//input[@name="Password"]')
pw.send_keys(password)
sign_on = driver.find_element_by_xpath('//button[@class="button signonbutton"]')
sign_on.click()

#Now move to distribution folder
folders = driver.find_element_by_xpath('//select[@id="field_gotofolder"]')
folders_select = Select(folders)
folders_select.select_by_visible_text('/ Distribution / Oklahoma')

labels = lambda: driver.find_elements_by_xpath('//table[@id="folderfilelisttable"]//td[@scope="row"]//span')

for i,label in enumerate(labels()):
    code = labels()[i].text[1:]
    labels()[i].click()
    files = driver.find_elements_by_xpath('//table[@id="folderfilelisttable"]//tr//td[@scope="row"]//span')
    buttons = driver.find_elements_by_xpath('//a[@class="button imgbutton icon_download"]')
    
    for file, button in zip(files,buttons):
        name = file.text
        button.click()
        try:
            dismiss = driver.find_element_by_xpath('//i[@class="ips-icon ips-icon-close"]')
            dismiss.click()
        except NoAlertPresentException as ex:
            pass
        while name not in os.listdir():
            time.sleep(1)
        if name[-3:]=='zip':
            with zipfile.ZipFile(name) as ax:
                ax.extractall()
            os.remove(name)
            for path, folders, files in os.walk(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Landing_Folder'):
                for name in files:
                    shutil.move(os.path.join(path,name),'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder\\'+name)
            for path, folders, files in os.walk(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Landing_Folder'):                
                for folder in folders:
                    shutil.rmtree(folder)
        else:
            pass
        
        file = os.listdir()[0]
        name = ' '.join(file.split('.'))
        if 'Claims Detail Data' in name:
            program = 'CMS'
            file_name = 'OK_'+program+'_'+code+'_'+str(qtr)+'Q'+str(yr)+file[-4:]
            path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Claims\\Oklahoma\\'+program+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
        else:
            if 'Provider' in name:
                os.remove(file)
                continue
            elif 'Utilization' in name:
                file_name = 'OK_Electronic_Invoice_'+code+'_'+str(qtr)+'Q'+str(yr)+file[-4:]
                program = 'CMS'
                path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Invoices\\Oklahoma\\'+program+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
            elif 'Federal' in name:
                program = 'CMS'
                kind = '_'.join(file.split('.')[0].split('_')[1:])
                file_name = 'OK_'+program+'_'+kind+'_'+code+'_'+str(qtr)+'Q'+str(yr)+file[-4:]
                path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Invoices\\Oklahoma\\'+program+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
            elif 'FOIprov' in name:
                os.remove(file)
                continue
            else:
                program = 'CMS Supplemental'
                kind = '_'.join(file.split('.')[0].split('_')[1:])
                file_name = 'OK_'+program+'_'+kind+'_'+code+'_'+str(qtr)+'Q'+str(yr)+file[-4:]
                path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Invoices\\Oklahoma\\'+program+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
        if os.path.exists(path)==False:
            os.makedirs(path)
        else:
            pass
        shutil.move(file,path+file_name)
    return_link = driver.find_element_by_xpath('//span[text()="Oklahoma"]')
    return_link.click()
    
driver.close()
















