# -*- coding: utf-8 -*-
"""
Created on Mon Aug 13 16:06:40 2018

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
login_credentials = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Delaware', usecols='D,E',dtype='str')
username = login_credentials.iloc[0,0]
password = login_credentials.iloc[0,1]
#Get the program map.  This map must be maintainted by the MHS team.
mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Delaware', usecols=[0,1],dtype='str')
mapper = dict(zip(mapper.Delaware,mapper.Lilly))
#build date time group
yq = str(yr)+'q'+str(qtr)
#Get to the page and login
wait = WebDriverWait(driver,10)
wait2 = WebDriverWait(driver,2)
driver.get('https://www.edsdocumenttransfer.com/')
user = wait.until(EC.element_to_be_clickable((By.ID,'form_username')))
user.send_keys('llymedicaid@lilly.com')
password = driver.find_element_by_id('form_password')
password.send_keys('Spring16!')
login = driver.find_element_by_id('submit_button')
login.click()

#Wait until the folder dropdown is available then 
#select the distribution folder
folders = wait.until(EC.element_to_be_clickable((By.ID,'field_gotofolder')))
folders_select = Select(folders)
folders_select.select_by_visible_text('/ Distribution')
#now give it some time to load
time.sleep(2)
sub_folders = lambda: driver.find_elements_by_xpath('//table[@id="folderfilelisttable"]//tr//td//img[@title="Folder"]')
for k,folder in enumerate(sub_folders()):
    sub_folders()[k].click()
    wait.until(EC.element_to_be_clickable((By.XPATH,'//span[text()="Parent Folder"]')))
    try:
        num_pages = driver.find_element_by_xpath('//table//tbody//tr[@class="nullSpacer"]//td//b')
        num_pages = num_pages.text[-1:]
    except NoSuchElementException as ex:
        num_pages = 1
    for i in range(int(num_pages)):
        new_files = lambda: driver.find_elements_by_xpath('//img/following-sibling::span[contains(text(),"%s")]'%(yq))
        if i==0:
            pass
        else:
            nex = driver.find_element_by_xpath('//span[text()="Next"]')
            nex.click()
            wait.until(EC.element_to_be_clickable((By.XPATH,'//span[text()="Parent Folder"]')))
        for j, file in enumerate(new_files()):
            name = new_files()[j].text
            Name = name[6:]
            label_code = Name[:5]
            file_type = Name[5:9]
            program = name[name.find('q')+2:name.find('.')].lower()
            ext = Name[-3:]
            lilly_program = mapper[program]
            new_files()[j].click()
            file_name = lilly_program+'_'+label_code+'_'+str(yr)+'_'+str(qtr)
            if file_type =='clda':
                download = wait2.until(EC.element_to_be_clickable((By.ID,'downloadLink')))
                download.click() 
                ext = '.txt'
                file_name = file_name+ext
                try:
                    close_pop_up = wait.until(EC.element_to_be_clickable((By.XPATH,'//div//ips-verifier//div//div//div//i')))
                    close_pop_up.click()
                except TimeoutException as ex: 
                    pass
                while name not in os.listdir():
                    time.sleep(1)
            elif file_type=='invd' and ext =='dat':
                download = wait2.until(EC.element_to_be_clickable((By.ID,'downloadLink')))
                download.click() 
                ext = '.txt'
                file_name = file_name+ext
                try:
                    close_pop_up = wait.until(EC.element_to_be_clickable((By.XPATH,'//div//ips-verifier//div//div//div//i')))
                    close_pop_up.click()
                except TimeoutException as ex: 
                    pass
                while name not in os.listdir():
                    time.sleep(1)
            else:
                download = wait.until(EC.element_to_be_clickable((By.XPATH,'//a//span[text()="Download"]')))
                download.click()
                ext = '.pdf'
                file_name = file_name+ext
                try:
                    close_pop_up = wait.until(EC.element_to_be_clickable((By.XPATH,'//div//ips-verifier//div//div//div//i')))
                    close_pop_up.click()
                except TimeoutException as ex: 
                    pass                    
                while name not in os.listdir():
                    time.sleep(1)
                
            if file_type == 'clda':
                file_type = 'Claims'
            else:
                file_type='Invoices'
            
            path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\'+file_type+'\\'+'Delaware'+'\\'+lilly_program+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
            
            if os.path.exists(path)==False:
                os.makedirs(path)
            else:
                pass

            shutil.move(name,path+file_name)

            driver.back()
            time.sleep(1.5)
    folders = wait.until(EC.element_to_be_clickable((By.ID,'field_gotofolder')))
    folders_select = Select(folders)
    folders_select.select_by_visible_text('/ Distribution') 
