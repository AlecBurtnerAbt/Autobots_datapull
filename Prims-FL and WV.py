# -*- coding: utf-8 -*-
"""
Created on Mon Jul 30 08:45:46 2018

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
login_credentials = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Prims', usecols='A,B',dtype='str')
username = login_credentials.iloc[0,0]
password = login_credentials.iloc[0,1]
#Open the webdriver, define the wait function, and get through the login page
chromeOptions = webdriver.ChromeOptions()
prefs = {'download.default_directory':'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder',
         'plugins.always_open_pdf_externally':True,
         'download.prompt_for_download':False}
chromeOptions.add_experimental_option('prefs',prefs)
os.chdir('O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder')
driver = webdriver.Chrome(chrome_options = chromeOptions, executable_path=r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\chromedriver.exe')
driver.implicitly_wait(30)
wait = WebDriverWait(driver,15)
driver.get('https://www.primsconnect.molinahealthcare.com/_layouts/fba/primslogin.aspx?ReturnUrl=%2f_layouts%2fAuthenticate.aspx%3fSource%3d%252FSitePages%252FHome%252Easpx&Source=%2FSitePages%2FHome%2Easpx')
i_accept = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_RadCheckBoxAccept"]/span[1]')))
i_accept.click()
user_name = driver.find_element_by_xpath('//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_txtUserName"]')
user_name.send_keys(username)
pass_word = driver.find_element_by_xpath('//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_txtPassword"]')
pass_word.send_keys(password)
login = driver.find_element_by_xpath('//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_btnLogin_input"]')
login.click()    
'''
First have to download the invoice files and parse them to get NDCs to request
Very similar to California
'''
#Now inside the webpage, begin selection process
submit_request = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_radLnkSubmitRequest_input"]')))
submit_request.click()    

#Now in the request page, navigate to invoice tab

invoice_request_page = lambda: wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_rtsRequest"]/div/ul/li[2]/a/span/span/span')))
invoice_request_page().click()        

#Now we have to start iterating through labeler codes, states, and programs
soup = BeautifulSoup(driver.page_source,'html.parser')
lists = soup.find_all('ul',attrs={'class':'rcbList'})
states_list = [x.text for x in lists[0]]
codes = [x.text for x in lists[2]]
for State in states_list:        
    states = lambda: wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_StateDropDown"]/table/tbody/tr')))
    ActionChains(driver).move_to_element(states()).click().send_keys(State).send_keys(Keys.ENTER).perform()
    time.sleep(5)
    available_quarter = lambda: driver.find_element_by_id('ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_AvailableQuarterLabelValue')
    soup = BeautifulSoup(driver.page_source,'html.parser')
    program_soup = soup.find_all('ul',attrs={'class':'rcbList'})
    program_list = [x.text for x in program_soup[1].contents]
    program_select = lambda: wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,'#ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_ProgramDropDown > table > tbody > tr > td.rcbInputCell.rcbInputCellLeft')))        
    if available_quarter().text =='Q0-0000':
        print(State+' is not ready to have invoics requested, moving to next state.')
        continue
    else:
        print('Begining requests for '+State)

    for Program in program_list:
        soup = BeautifulSoup(driver.page_source,'html.parser')
        verify_program = soup.find('input',attrs={'title':'Select Program'})
        while verify_program.get('value') != Program:
            print('a')
            xpath = '//li[text()="'+Program+'"]'
            current_program = lambda: wait.until(EC.visibility_of_element_located((By.XPATH,xpath)))
            print('b')
            ActionChains(driver).move_to_element(program_select()).click().perform()
            time.sleep(3)
            print('c')
            current_program().click()
            time.sleep(3)
            print('d')
            soup = BeautifulSoup(driver.page_source,'html.parser')
            verify_program = soup.find('input',attrs={'title':'Select Program'})
        print('Reqesting invoice for '+State+', '+Program)
        for Code in codes:
            print('Requesting invoice for '+Code)
            xpath2 = '//span[text()="'+Code+'"]'
            code = lambda: wait.until(EC.element_to_be_clickable((By.XPATH,xpath2)))
            time.sleep(3)
            print('j')
            ActionChains(driver).move_to_element(code()).double_click().perform()
            time.sleep(3)
            print('k')
            submit_button = lambda: wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_EInvoiceSubmitButton_input"]')))
            print('l')
            time.sleep(3)
            ActionChains(driver).move_to_element(submit_button()).double_click().perform()
            time.sleep(3)
            success  = lambda: wait.until(EC.presence_of_element_located((By.CLASS_NAME,'PC_SuccessMessage')))
            while success().is_displayed()==False:
                time.sleep(3)
            print(success().text)
            del(success)
        print('refreshing')
        try:
            driver.refresh()
            alert = driver.switch_to.alert
            alert.accept()
        except:
            pass
  


    
    
  xxx['West Virginia'][0]['N'].index(ndc)  
   v = soup.find('li',attrs={'class':'rcbItem'}) 
    
    
    #ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_TYearQuarterDropDown_DropDown > div > ul > li:nth-child(41)
    //*[@id="ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_TYearQuarterDropDown_DropDown"]/div/ul/li[41]
    
    