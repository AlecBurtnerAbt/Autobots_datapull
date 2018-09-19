# -*- coding: utf-8 -*-
"""
Created on Tue Aug 21 11:46:59 2018

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
    #Open the webdriver, define the wait function, and get through the login page
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
driver.get('https://rxmaxmed.optum.com/rxmaxpiconvm/rxmax/login')
login_credentials = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Optum', usecols='A,B',dtype='str')
username = login_credentials.iloc[0,0]
password = login_credentials.iloc[0,1]
#input user id and password
user = driver.find_element_by_xpath('//input[@name="user_name"]')
user.send_keys(username)
pw = driver.find_element_by_xpath('//input[@name="password"]')
pw.send_keys(password)
login_button = driver.find_element_by_xpath('//input[@value="Login"]')
login_button.click()
yq = str(qtr)+'Q'+str(yr)

#Now have to execute a switch to


#Now that we're inside the portal we have to switch to the right frame
#and then click the "History" button to get to the most recent reports
driver.switch_to.window(driver.window_handles[1])
download_reports = driver.find_element_by_xpath('//a[@id="topDownload_Reports"]')
ActionChains(driver).move_to_element(download_reports).pause(1).click().perform()
download_reports2 = driver.find_element_by_xpath('//a[@id="Download_Reports"]')
download_reports2.click()

history_button = driver.find_element_by_xpath('//input[@value="History"]')
history_button.click()

#We now can see the most recent files



pages = lambda: driver.find_elements_by_xpath('//table//tr[@class="pageNavProperties"]//td/a')
#For each page define the rows, links, dates, and data
for i,page in enumerate(pages()):
    print('Working on page '+str(i+1))
    rows = lambda: driver.find_elements_by_xpath('//tr[count(child::td)>3]')
    data = []
    [data.append(''.join(x.text.replace('\n',' ').split(' ')[:2])) for x in rows()]
    links = driver.find_elements_by_xpath('//tr[count(child::td)>3]//a[contains(@href,"selectRecordForDownload")]')
    dates = driver.find_elements_by_xpath('//tr[count(child::td)>3]//td[2]')
    dates = [x.text.strip() for x in dates]
    #If there is a row that has the current quarter in it, continue
    if any(map((lambda x: yq in x),[x.text for x in rows()]))==True:
    
        for link, date in zip(links, dates):
            print('Downloading '+link.text+' for '+date)
            file = link.text
            if date == yq: 
                success = 0
                while success == 0:
                    print('Clicking link...')
                    link.click()
                    time.sleep(3)
                    if link.text in os.listdir():
                        success=1
                        print('Success!')
                    else:
                        print('Retrying...')
                        pass
                    with zipfile.ZipFile(link.text) as ax:
                        ax.extractall()
                    os.remove(link.text)
                    label_code = file.split('-')[0]
                    program = file.split('-')[1].split('_')[3]
                    report_num = file.split('-')[1].split('_')[2]
                    for file in os.listdir():
                        file_name = label_code+'_'+program+'_'+report_num+'_'+yq+file[-4:]
                        path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Invoices\\Nevada\\'+program+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
                        if os.path.exists(path)==False:
                            os.makedirs(path)
                        else:
                            pass
                        shutil.move(file,path+file_name)


            else:
                pass
            
        #If there isn't a row that has the current quarter, stop    
    else:
        print('No more current files on page '+str(i+1))
        break
    print('Done with page '+str(i+1)+' moving onto page '+str(i+2))
    pages()[i].click()
    time.sleep(3)
    
    












