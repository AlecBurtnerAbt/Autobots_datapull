# -*- coding: utf-8 -*-
"""
Created on Fri Aug 24 09:06:01 2018

@author: c252059
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
#open the driver and get to the site and login
chromeOptions = webdriver.ChromeOptions()
prefs = {'download.default_directory':'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder',
         'plugins.always_open_pdf_externally':True,
         'download.prompt_for_download':False}
chromeOptions.add_experimental_option('prefs',prefs)
driver = webdriver.Chrome(chrome_options = chromeOptions, executable_path=r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\chromedriver.exe')
os.chdir('O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder')
for file in os.listdir():
    os.remove(file)
mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Illinois', usecols='D,E',dtype='str')
mapper = dict(zip(mapper['State Program'],mapper['Flex Program']))
time_stuff = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx', sheet_name = 'Year-Qtr',use_cols='A:B')
yr = time_stuff.iloc[0,0]
qtr = time_stuff.iloc[0,1]
yq=str(yr)+str(qtr)
login_credentials = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Illinois', usecols=[0,1],dtype='str')
username = login_credentials.iloc[0,0]
password = login_credentials.iloc[0,1]
driver.get(r'https://rsp.ilgov.emdeon.com/RebateServicesPortal/login/home?goto=http://rsp.ilgov.emdeon.com/RebateServicesPortal/')
wait = WebDriverWait(driver,10)
wait2 = WebDriverWait(driver,2)
#find username and password and pass the login credentials

user = driver.find_element_by_xpath('//input[@id="username"]')
user.send_keys(username)
pw = driver.find_element_by_xpath('//input[@id="password"]')
pw.send_keys(password)
login = driver.find_element_by_xpath('//input[@value="Login"]')
login.click()

#Now to navigate past the next page

accept = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@value="Accept"]')))
accept.click()

#Now to get to the invoices page
invoices = driver.find_element_by_xpath('//a[text()="Invoices"]')
invoices.click()

#We have labeler, state, and typ

labeler = lambda: driver.find_element_by_xpath('//select[@id="labeler"]')
labeler_select = lambda: Select(labeler())
options = [x.text for x in labeler_select().options]
options = options[1:]

types = lambda: driver.find_element_by_xpath('//select[@name="docType"]')
types_select = lambda: Select(types())

time_stamp = lambda: driver.find_element_by_xpath('//input[@id="period"]')
time_stamp().send_keys(yq)



#now begin looping
for label in options:
    labeler_select().select_by_visible_text(label)
    for report in list(mapper.keys()):
        submit_button = driver.find_element_by_xpath('//input[@value="Submit"]')
        type_select.select_by_visible_text(report)
        time_stamp.send_keys(yq)
        submit_button.click()
        while wait.until(EC.staleness_of(submit_button))==False:
            time.sleep(.2)
        downloads = driver.find_elements_by_xpath('//)
        
        
        

















