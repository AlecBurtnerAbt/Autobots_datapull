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
from mail_maker import send_message
def generate_browser():
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
    return driver
    
def get_invoices(driver):

    time_stuff = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx', sheet_name = 'Year-Qtr',use_cols='A:B')
    yr = time_stuff.iloc[0,0]
    qtr = time_stuff.iloc[0,1]
    yq=str(yr)+str(qtr)
    login_credentials = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Vermont', usecols=[0,1],dtype='str')
    username = login_credentials.iloc[0,0]
    password = login_credentials.iloc[0,1]
    to_address = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Notification Address', usecols='A',dtype='str',names=['Email'],header=None).iloc[0,0]

    driver.get(r'https://www.vermontrsp.com/RebateServicesPortal/login/home?goto=http://www.vermontrsp.com/RebateServicesPortal/')
    user = driver.find_element_by_id('username')
    user.send_keys(username)
    pass_word = driver.find_element_by_id('password')
    pass_word.send_keys(password)
    login = driver.find_element_by_id('submit')
    login.click()
    mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Vermont', usecols='D,E',dtype='str')
    mapper = dict(zip(mapper['State Program'],mapper['Flex Program']))
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
    dme_index = types.index([x for x in types if 'DME' in x][0])
    types[dme_index] = 'DME'
    medicare_wrap_index = types.index([x for x in types if 'Wrap' in x][0])
    types[medicare_wrap_index] = 'Medicare_Wrap'
    types_2 = []
    invoices_obtained = []
    invoices_to_get = list(itertools.product(codes,types))  
    reference_list = []
    values = [x.get_attribute('value') for x in type_select().options][1:]
    mapper2 = dict(zip(types,values))
    #helper function to check dates in results table
    def check():
        invoice_period = driver.find_element_by_xpath('//table[@id="invoiceResults"]/tbody/tr/td[4]')
        if invoice_period.text == yq:
            available =1
        else:
            available = 0
        return available
    for label, report in invoices_to_get:
        counter = 0
        available =0
        while available ==0:
            counter +=1
            code_select().select_by_visible_text(label)
            time.sleep(1)
            type_select().select_by_index(types.index(report)+1)
            time.sleep(1)
            submit = driver.find_element_by_xpath('//input[@id="invSubmit"]')
            submit.click()
            wait.until(EC.staleness_of(submit))
            available = check()
            if available ==0:
                driver.refresh()
            else:
                pass
            time.sleep(counter*1.5)
        buttons = driver.find_elements_by_xpath('//a[@class="btn"][contains(@onclick,"download")]')
        for button in buttons:
            button.click()
            alert = driver.switch_to.alert
            alert.accept()
            if buttons.index(button)==0:
                file_type='.txt'
            else:
                file_type = '.pdf'
            file_name = f'VT-{label}-{yq}-{report}{file_type}'
            while file_name not in os.listdir():
                time.sleep(1)
            #Now open the file and return the NDCs associated to the label code and program
            if file_type =='.txt':
                read_flag = 0
                while read_flag==0:
                    try:
                        with open(file_name) as f:
                            lines = f.readlines()
                            ndcs = list(set([line[6:17] for line in lines]))
                            ndcs = [ndc for ndc in ndcs if len(ndc)>1]
                            reference_list.append((label,report,ndcs))
                        read_flag=1
                    except PermissionError as ex:
                        pass
            else:
                pass
            try:
                report_value = mapper2[report]
                flex_name = mapper[report_value]
            except KeyError as err:
                flex_name = report
            new_name = f'VT_{flex_name}_{qtr}Q{yr}_{label}{file_type}'
            path =  f'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Invoices\\Vermont\\{flex_name}\\{yr}\\Q{qtr}\\'
            if os.path.exists(path)==False:
                os.makedirs(path)
            shutil.move(file_name,path+new_name)
    from collections import defaultdict
    master_dict = defaultdict(dict)        
    for label, report, ndcs in reference_list:
        if len(ndcs)>0:
            master_dict[label][report]=ndcs        
    invoices_obtained = [f'{label}-{report}' for label,report,ndcs in reference_list if len(ndcs)>0]
    body = 'The following invoices were obtained\n'+'\n'.join(invoices_obtained)
    subject = 'Vermont Invoices'
    send_message(subject,body,to_address)
    driver.stop_client()
    driver.close()
    return yq, username, password, master_dict,types,reference_list
def make_chunks(master_dict):
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
    in_flag = 0
    wait = WebDriverWait(driver,10)
    while in_flag == 0:
        login.click()
        try:
            canary = wait.until(EC.element_to_be_clickable((By.ID,'terms')))        
            in_flag = 1
        except TimeoutException as ex:
            driver.get(r'https://www.vermontrsp.com/RebateServicesPortal/login/home?goto=http://www.vermontrsp.com/RebateServicesPortal/')
            user = driver.find_element_by_id('username')
            user.send_keys(username)
            pass_word = driver.find_element_by_id('password')
            pass_word.send_keys(password)
            login = driver.find_element_by_id('submit')
    
    
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
        if program == 'Medicare_Wrap':
            program = 'VPharm/SPAP (Medicare Wrap)'
        elif program == 'DME':
            program = 'State Only Diabetic (DME)'
        while success==0:
            try:
                report = driver.find_element_by_xpath('//select[@name="stateReportId"]')
                select_report = Select(report)        
                select_report.select_by_index(1)
                
                ndc_in = driver.find_element_by_xpath('//input[@name="ndc"]')
                ndc_in.send_keys(ndc)
                wait.until(EC.element_to_be_clickable((By.XPATH,'//select[@name="docType"]')))
                docType = driver.find_element_by_xpath('//select[@name="docType"]')
                select_docType = Select(docType)
                select_docType.select_by_visible_text(program.replace('_',' '))
                wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@name="rpuStart"]')))
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
                    driver.refresh()
                    pass
            except TimeoutException as ex:
                driver.refresh()
                pass
    driver.stop_client()
    driver.close()
  
    
def download_reports():
    driver = generate_browser()
    os.chdir('O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder')
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
    links = [x.find_element_by_xpath('td//a[@href="#"]/i') for x in rows]
    master_df = pd.DataFrame()
    files = []
    for name, link in zip(names, links):
        success_flag = 0
        while success_flag==0:
            #get info for file name
            ndc = name.split(' ')[7]
            state = name.split(' ')[8]
            program = name.split(' ')[10]
            value = mapper[program]
            first_half = '_'.join(name.split(' ')[:5])
            second_half = '-'.join(name.split(' ')[-4:]).replace(program,mapper[program])
            download_name = '-'.join([first_half,second_half])+'.xls'
            if download_name in os.listdir():
                continue
            else:
                pass
            files.append(download_name)
            #download the file
            
            counter = 0
            try:
                link.click()
            except WebDriverException as ex:
                driver.refresh()
                continue
            while download_name not in os.listdir() and counter<10:
                time.sleep(1)
                counter+=1
            if counter >9:
                pass
            else:
                success_flag=1
    
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
        file_name = 'VT_'+splitter+'_'+str(qtr)+'Q'+str(yr)+'.xlsx'
        if os.path.exists(path)==False:
            os.makedirs(path)
        else:
            pass
        os.chdir(path)
        frame.to_excel(file_name, engine='xlsxwriter', index=False)
    closers = lambda: driver.find_elements_by_xpath('//a[@title="Delete"]/i')
    for close in closers():
        closers()[0].click()
        alert = driver.switch_to.alert
        alert.accept()
    driver.stop_client()
    driver.close()
    for file in os.listdir():
        os.remove(file)

def main():
    driver = generate_browser()
    yq, username, password, master_dict,types,reference_list = get_invoices(driver)  
    chunks = make_chunks(master_dict)
    processes = [mp.Process(target=getReports,args=(i,chunk,types)) for i,chunk in enumerate(chunks)]
    for p in processes:
        p.start()
    for p in processes:
        p.join()   


    download_reports()
    for p in processes:
        p.terminate()
    
if __name__=='__main__':
    main()










