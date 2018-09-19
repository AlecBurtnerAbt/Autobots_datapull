# -*- coding: utf-8 -*-
"""
Created on Wed Jul 18 13:44:28 2018

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
import multiprocessing as mp

def get_invoices():
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
    time_stuff = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx', sheet_name = 'Year-Qtr',use_cols='A:B')
    yr = time_stuff.iloc[0,0]
    qtr = time_stuff.iloc[0,1]
    yq=str(yr)+str(qtr)
    login_credentials = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Vermont', usecols=[0,1],dtype='str')
    username = login_credentials.iloc[0,0]
    password = login_credentials.iloc[0,1]
    driver.get(r'https://www.vermontrsp.com/RebateServicesPortal/login/home?goto=http://www.vermontrsp.com/RebateServicesPortal/')
    user = driver.find_element_by_id('username')
    user.send_keys(username)
    pass_word = driver.find_element_by_id('password')
    pass_word.send_keys(password)
    login = driver.find_element_by_id('submit')
    login.click()
    
    wait = WebDriverWait(driver,10)
    accept = wait.until(EC.element_to_be_clickable((By.ID,'terms')))
    accept.click()
    
    #invoice stuff is below this
    
    invoices = driver.find_element_by_xpath('//a[text()="Invoices"]')
    invoices.click()
    code_dropdown = lambda: driver.find_element_by_id('labeler')
    code_select = lambda: Select(code_dropdown())
    codes = code_select().options
    type_dropdown = lambda: WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.ID,'docType')))
    type_select = lambda: Select(type_dropdown())
    types = type_select().options
    report_ndcs = []
    codes = [item.text for item in codes[1:]]
    types = [item.text for item in types[1:]]
    types_2 = []
    for i,typ in enumerate(types):
        if len(typ.split(' '))==1:
            types_2.append(typ)
        elif len(typ.split('(')[-1].split(' '))==1:
            _ = typ.split(' ')[-1].replace('(','').replace(')','').replace(' ','_')
            types_2.append(_)
        else:
            _='_'.join(typ.split('(')[1].split(' ')).replace(')','')
            types_2.append(_)
    not_downloaded = []
    master_dict = {}
    #loop through the codes from the dropdown to get the invoices
    for code in codes:
        interim = {}
        for typ in types:       
            
            present = 0
            zounter=0
            while present == 0 and zounter <11:
                time.sleep(zounter*1.2)
                code_select().select_by_visible_text(code)
                type_select().select_by_visible_text(typ)
                print('selected labeler code: '+code)
                print('selected report type: '+typ)
                submit_button = driver.find_element_by_id('invSubmit')
                submit_button.click()
                wait.until(EC.staleness_of(submit_button))
                check = driver.find_element_by_xpath('//table[@id="invoiceResults"]//tr//td[4]')
                if check.text=='Please Wait...':
                    zounter+=1
                    driver.refresh()
                    pass
                else:
                    present=1
            if zounter >9:
                print('Could not generate reports for '+code+','+typ)
            else:
                pass
            print('generating reports to download')
            CMS_button = lambda: WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH,'//form[@name="downloadCms"]//a[@title="Download"]')))
            PDF_button = lambda: driver.find_element_by_xpath('//form[@name="downloadPdf"]//a[@title="Download"]')
            buttons = [CMS_button, PDF_button]
            n = types.index(typ)
            path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Invoices\\Vermont\\'+types_2[n]+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
            if os.path.exists(path)==False:
                os.makedirs(path)
            else:
                pass
            for i, button in enumerate(buttons):
                success=0
                kounter = 0                
                while success==0 and kounter <10:
                    if i==0:
                        time.sleep(kounter*1.5)
                        print('CMS')
                        button().click()   
                        print('downloading')      
                        WebDriverWait(driver,10).until(EC.alert_is_present())
                        alert = driver.switch_to.alert
                        alert.accept()
                        file_name = 'VT-'+code+'-'+yq+'-'+types_2[n]+'.txt'
                        counter=0
                        while file_name not in os.listdir() and counter<11:
                            time.sleep(1)
                            counter+=1
                        if file_name not in os.listdir():
                            driver.back()
                            kounter+=1
                            continue
                        else:
                            file = open(file_name,'r')
                            lines = file.readlines()
                            file.close()
                            ndcs = list(set([line[6:17] for line in lines]))
                            ndcs = [ndc for ndc in ndcs if len(ndc)>1]
                            interim.update({typ:ndcs})
                            shutil.move(file_name,path+file_name)
                            success=1
                                
                    else:
                        time.sleep(kounter*1.5)
                        print('PDF')
                        button().click()
                        WebDriverWait(driver,10).until(EC.alert_is_present())
                        alert = driver.switch_to.alert
                        alert.accept()
                        counter=0
                        file_name = 'VT-'+code+'-'+yq+'-'+types_2[n]+'.pdf'
                        while file_name not in os.listdir() and counter <11:
                            time.sleep(1)
                            counter+=1
                        if file_name not in os.listdir():
                            driver.back()
                            kounter+=1
                            continue
                        else:
                            shutil.move(file_name,path+file_name)
                            success=1
                
                if kounter > 9:
                    print('Tried 10 times, could not download '+typ+' for label code '+code)
                    not_downloaded.append((typ,code))
                    driver.back()
                else:
                    pass
        master_dict.update({code:interim})
            
            
    return yq, username, password, master_dict

def make_chunks(dictionary):
    #Break the information for each report down into 
    reports = []
    for key in master_dict.keys():
        for key2 in master_dict[key].keys():
            for value in master_dict[key][key2]:
                    report = (key,key2,value)
                    reports.append(report)
    n = round(len(reports)/(mp.cpu_count()-1))
    chunks = [reports[x:x+n] for x in range(0,len(reports),n)]
    return chunks

##reports stuff is below this
def getReports(num,chunk,types):
    print('Working on chunk: '+str(num))
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
    time_stuff = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx', sheet_name = 'Year-Qtr',use_cols='A:B')
    yr = time_stuff.iloc[0,0]
    qtr = time_stuff.iloc[0,1]
    yq=str(yr)+str(qtr)
    login_credentials = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Vermont', usecols=[0,1],dtype='str')
    username = login_credentials.iloc[0,0]
    password = login_credentials.iloc[0,1]
    driver.get(r'https://www.vermontrsp.com/RebateServicesPortal/login/home?goto=http://www.vermontrsp.com/RebateServicesPortal/')
    user = driver.find_element_by_id('username')
    user.send_keys(username)
    pass_word = driver.find_element_by_id('password')
    pass_word.send_keys(password)
    login = driver.find_element_by_id('submit')
    login.click()
    
    wait = WebDriverWait(driver,10)
    accept = wait.until(EC.element_to_be_clickable((By.ID,'terms')))
    accept.click()
    invoices = driver.find_element_by_xpath('//a[text()="Invoices"]')
    invoices.click()
    type_dropdown = lambda: WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.ID,'docType')))
    type_select = lambda: Select(type_dropdown())
    types = type_select().options
    types = [item.text for item in types[1:]]
    types_2=[]
    for i,typ in enumerate(types):
        if len(typ.split(' '))==1:
            types_2.append(typ)
        elif len(typ.split('(')[-1].split(' '))==1:
            _ = typ.split(' ')[-1].replace('(','').replace(')','').replace(' ','_')
            types_2.append(_)
        else:
            _='_'.join(typ.split('(')[1].split(' ')).replace(')','')
            types_2.append(_)
    
    reports_tab = wait.until(EC.element_to_be_clickable((By.XPATH,'//a[@href="/RebateServicesPortal/reports/index"]')))       
    reports_tab.click()                
    report = lambda: wait.until(EC.element_to_be_clickable((By.XPATH,'//select[@id="reportList"]')))
    report_select = lambda: Select(report())
    #Now starting iterating through the chunk
    for label, program, ndc in chunk:
        success = 0
        while success==0:
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
            soup = BeautifulSoup(driver.page_source,'html.parser')
            Reports = [x.text.strip() for x in soup.find_all('td')]
            if any(map((lambda x: (ndc+' VT '+yq+' '+types_2[types.index(program)]) in x),Reports)):
                success=1
            else:
                pass
    driver.close()
    
def download_reports():
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
    time_stuff = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx', sheet_name = 'Year-Qtr',use_cols='A:B')
    yr = time_stuff.iloc[0,0]
    qtr = time_stuff.iloc[0,1]
    yq=str(yr)+str(qtr)
    login_credentials = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Vermont', usecols=[0,1],dtype='str')
    username = login_credentials.iloc[0,0]
    password = login_credentials.iloc[0,1]
    driver.get(r'https://www.vermontrsp.com/RebateServicesPortal/login/home?goto=http://www.vermontrsp.com/RebateServicesPortal/')
    user = driver.find_element_by_id('username')
    user.send_keys(username)
    pass_word = driver.find_element_by_id('password')
    pass_word.send_keys(password)
    login = driver.find_element_by_id('submit')
    login.click()
    
    wait = WebDriverWait(driver,10)
    accept = wait.until(EC.element_to_be_clickable((By.ID,'terms')))
    accept.click()
    reports_tab = driver.find_element_by_xpath('//a[text()="Reports"]')       
    reports_tab.click()                
    report = lambda: driver.find_element_by_xpath('//select[@id="reportList"]')
    report_select = lambda: Select(report())
    
    types = driver.find_element_by_xpath('//select[@id="docType"]')
    types_select = Select(types)
    programs = [x.text for x in types_select.options][1:]
    types_2=[]
    for i,typ in enumerate(programs):
        if len(typ.split(' '))==1:
            types_2.append(typ)
        elif len(typ.split('(')[-1].split(' '))==1:
            _ = typ.split(' ')[-1].replace('(','').replace(')','')
            types_2.append(_)
        else:
            _='_'.join(typ.split('(')[1].split(' ')).replace(')','')
            types_2.append(_)
    values = [x.get_attribute('value') for x in driver.find_elements_by_xpath('//select[@id="docType"]/option')][1:]
    mapper = dict(zip(types_2,values))    
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
        path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Claims\\Vermont\\'+splitter+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
        file_name = 'PA_'+splitter+'_'+str(qtr)+'Q'+str(yr)+'.csv'
        if os.path.exists(path)==False:
            os.makedirs(path)
        else:
            pass
        os.chdir(path)
        frame.to_csv(file_name)
    driver.close()
        
if __name__=='__main__':
    #yq, username, password, master_dict = get_invoices()  

    #chunks = make_chunks(master_dict)


    processes = [mp.Process(target=getReports,args=(i,chunk,types)) for i,chunk in enumerate(chunks)]
    for p in processes:
        p.start()       
    for p in processes:
        p.join()   
        
    download_reports()
'''
deletes = lambda: driver.find_elements_by_xpath('//a[@title="Delete"]')
for i in range(len(deletes())):
    deletes()[-1].click()
    alert = driver.switch_to.alert
    alert.accept()
'''









