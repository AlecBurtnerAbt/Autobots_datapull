# -*- coding: utf-8 -*-
"""
Created on Tue Sep 11 15:16:19 2018

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
import multiprocessing as mp

def getReports(num,chunk):
    print('Working on chunk: '+str(num))
    os.chdir('C:/Users/')
    chromeOptions = webdriver.ChromeOptions()
    prefs = {'download.default_directory':'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder',
         'plugins.always_open_pdf_externally':True,
         'download.prompt_for_download':False}
    chromeOptions.add_experimental_option('prefs',prefs)
    chromeOptions.add_argument('--headless')
    chromeOptions.add_argument('--disable-gpu')
    driver = webdriver.Chrome(chrome_options = chromeOptions, executable_path=r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\chromedriver.exe')
    os.chdir('O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder')
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
    #Now starting iterating through the chunk
    for label, program, ndc in chunk:
        report = driver.find_element_by_xpath('//select[@name="stateReportId"]')
        select_report = Select(report)        
        select_report.select_by_index(1)
        
        ndc_in = driver.find_element_by_xpath('//input[@name="ndc"]')
        ndc_in.send_keys(ndc)
        
        docType = driver.find_element_by_xpath('//select[@name="docType"]')
        select_docType = Select(docType)
        select_docType.select_by_visible_text(program.replace('_',' '))
        
        rpu = driver.find_element_by_xpath('//input[@name="rpuStart"]')
        rpu.send_keys(yq)
        
        submit_button= driver.find_element_by_xpath('//input[@value="Submit"]')
        submit_button.click()
        accept = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@value="Accept"]')))
        accept.click()
        wait.until(EC.staleness_of(accept))