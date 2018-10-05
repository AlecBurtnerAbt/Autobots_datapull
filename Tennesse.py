# -*- coding: utf-8 -*-
"""
Created on Tue Aug 28 17:45:00 2018

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
from mail_maker import send_message
def pull():
    os.chdir('C:/Users/')
    
    #make sure the directory is the downloads folder!
    
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
    login_credentials = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Magellan', usecols='A,B',dtype='str')
    username = login_credentials.iloc[0,0]
    password = login_credentials.iloc[0,1]
    #Login with provided credentials
    driver.get('https://einvoiceop.magellanmedicaid.com/rebate')   
    user_name = driver.find_element_by_xpath('//*[@id="username"]')
    user_name.send_keys(username)
    pass_word = driver.find_element_by_xpath('//*[@id="password"]')
    pass_word.send_keys(password)
    to_address = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Notification Address', usecols='A',dtype='str',names=['Email'],header=None).iloc[0,0]

    wait = WebDriverWait(driver,10)
    wait2 = WebDriverWait(driver,3)
    login_button = driver.find_element_by_xpath('//*[@id="content"]/div/div[2]/fieldset/ol[2]/li/input[3]')
    login_button.click()
    '''
    Now moving onto invoices
    '''
    #These lines of code get all available options
    invoices_tab = wait.until(EC.element_to_be_clickable((By.ID,'mainForm:invoices')))
    invoices_tab.click()
    business_line = lambda: driver.find_element_by_id('mainForm:srchBusinessLine')
    business_line_select = lambda: Select(business_line())
    business_line_types = [x.text for x in business_line_select().options]
    year_qtr = lambda: driver.find_element_by_id('mainForm:srchYearQtr')
    issue_list = []
    cld_to_get = []
    already_have = []
    state = "Tennessee"
    retrieved = []
    for root, dirs, files in os.walk(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Test\Invoices'):
        already_have.append(root)
    
    already_have = [x.split('\\')[-3] for x in already_have if len(x.split('\\'))>9 and x.split('\\')[-1]=='Q'+str(qtr)]
    
    #Now starting to loop through the options and downloading the files
    #start of business line loop
    for biz in business_line_types:
        print("Working on "+biz+" files")
        business_line_select().select_by_visible_text(biz)
        time.sleep(2)
        year_qtr().clear()
        year_qtr().send_keys(str(yr)+str(qtr))
        time.sleep(1)
        search = driver.find_element_by_xpath('//input[@value="Search"]')
        search.click()
        time.sleep(1)
        invoices = lambda: driver.find_elements_by_xpath('//table//tbody//tr//td//input')
        names = [x.get_attribute('name') for x in invoices()]
        if len(names)==0:
            continue
        else:
            pass
        invoice_labels = lambda: driver.find_elements_by_xpath('//table//tbody//tr//td[string-length()>0][1]')
        invoice_labels = [x.text for x in invoice_labels()]
        #Loop through the available invoices for the program
        for inv_name, label in zip(names, invoice_labels):
            invoice = lambda: driver.find_element_by_name(inv_name)
            invoice().click()
            program = label.split('-')[0][-3:]
            _ = [label.split('-')[1],program]
            cld_to_get.append(_)
            print('Downloading '+label)
            time.sleep(1)
            invoice_options = lambda: driver.find_element_by_id('mainForm:selectedFormatType')
            invoice_options_select = lambda: Select(invoice_options())
            invoice_options_options = [x.text for x in invoice_options_select().options]
            #Get both the PDF and the CMS file for the invoice
            for i,option in enumerate([invoice_options_options[0],invoice_options_options[-1]]):
                invoice_options_select().select_by_visible_text(option)                 
                continue_button = wait.until(EC.element_to_be_clickable((By.ID,'mainForm:continueButton')))
                continue_button.click()
                time.sleep(5)
                if i ==0:
                    if 'Invoice Report .pdf' in os.listdir():
                        success_flag = 1
                    else: 
                        pass
                    print('Downloading PDF format.')
                    try:
                        zzz = wait2.until(EC.element_to_be_clickable((By.XPATH,'//a[@href="mailto:rebate@magellanhealth.com"]')))
                        issue_text = program+' '+label+' PDF was not downloaded due to website error, looping unitl downloaded'
                        print(issue_text)
                        issue = [program,label]
                        print('a')
                        driver.back()
                        success_flag = 0
                        count = 0
                        print('b')
                        while (success_flag ==0 and count <10):
                            if 'Invoice Report .pdf' in os.listdir():
                                success_flag = 1
                            else:
                                pass
                            driver.refresh()
                            print('c')
                            wait.until(EC.element_to_be_clickable((By.NAME,inv_name)))
                            invoice().click()
                            invoice_options_select().select_by_visible_text(option)
                            if 'Invoice Report .pdf' in os.listdir():
                                success_flag = 1
                            else:
                                pass
                            print('d')
                            continue_button = driver.find_element_by_id('mainForm:continueButton')
                            continue_button.click()
                            print('e')
                            time.sleep(5)
                            kount=0
                            while ('Invoice Report .pdf' not in os.listdir() and kount <10):
                                print('f')
                                time.sleep(1)
                                kount+=1  
                            count +=1
                            print('g')
                            try:
                                print('h')
                                zzz = wait2.until(EC.element_to_be_clickable((By.XPATH,'//a[@href="mailto:rebate@magellanhealth.com"]')))
                                driver.back()
                                try:
                                    print('i')
                                    wait2.until(EC.visibility_of_element_located((By.ID,'suggestions-list')))
                                    driver.refresh()
                                except TimeoutException as ec:
                                    pass
                            except TimeoutException as ex:
                                print('j')
                                if "Invoice Report .pdf" in os.listdir():
                                    success_flag = 1
                                else:
                                    pass
                        if count >10:
                            print('Tried to get PDF invoice for ' + program+' moving onto next')
                            issues_list.append(issue)
                        else:
                            print('Download success after '+str(count)+' tries!')
                            pass                                       
                    except TimeoutException as ex:
                        pass
                    else:
                        pass
                    path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\'+'Invoices\\'+state+'\\'+program+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
                    file_name = program+'_'+'_'.join(label.split('-')[1:])+'.pdf'
                    if os.path.exists('path')==False:
                        os.makedirs(path, exist_ok=True)
                    else:
                        pass
                    shutil.move("Invoice Report .pdf",path+file_name)
                    retrieved.append(label)
                    time.sleep(1)                   
                else:
                    print('Downloading CMS format.')
                    try:
                        zzz = wait2.until(EC.element_to_be_clickable((By.XPATH,'//a[@href="mailto:rebate@magellanhealth.com"]')))
                        issue = program+' '+label+' CMS was not downloaded due to website error, please try again later.'
                        print(issue)
                        driver.get('https://einvoice.magellanmedicaid.com/rebate/spring/main?execution=e2s1')
                        invoices_tab = driver.find_element_by_id('mainForm:invoices')
                        invoices_tab.click()
                        year_qtr().clear()
                        year_qtr().send_keys(str(yr)+str(qtr))
                        business_line_select().select_by_visible_text(biz)
                        program_name_select().select_by_visible_text(program)
                        search = driver.find_element_by_xpath('//*[@id="srchInvoiceDiv"]/ol[2]/li/input')
                        search.click()
                        time.sleep(2)
                        invoice().click()
                        issue_list.append(issue)
                        time.sleep(5)
                        continue
                    except TimeoutException as ex:
                        pass
                    while 'einvoice.txt' not in os.listdir():
                        time.sleep(1)
                    path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\'+'Invoices\\'+state+'\\'+program+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
                    file_name = program+'_'+'_'.join(label.split('-')[1:])+'.txt' 
                    if os.path.exists('path')==False:
                        os.makedirs(path, exist_ok=True)
                    else:
                        pass
                    shutil.move("einvoice.txt",path+file_name)
                    retrieved.append(label)
                    time.sleep(1)
            invoice().click()
    pulled = ''.join(['{} \n'.format(x) for x in invoice_labels])
    body = 'The following invoices were pulled from the Tennessee portal \n'+pulled+'\n Beep Boop I am a bot.  This \
        message was generated by the Tennessee data pull bot.  For any issues please see the support team.'
     ########################################################CLD Below this line###########################################   
    
    send_message('TN Invoices',body=body,to=to_address) 
    claims_details = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="mainForm:claims"]')))
    claims_details.click()
    
    
    yq = str(yr)+str(qtr)
    
    
    """
    Sets dropdown default to null
    """
    labeler = lambda: driver.find_element_by_id('mainForm:labelerCode')
    labeler_select = lambda: Select(labeler())
    codes = [x.text for x in labeler_select().options if len(x.text)>1]
    year_qtr = lambda: driver.find_element_by_id('mainForm:srchYearQtr')
    year_qtr().clear()
    year_qtr().send_keys(yq)
    program_name = lambda: driver.find_element_by_id('mainForm:srchProgramName')
    program_name_select = lambda: Select(program_name()) 
    programs = [x.text for x in program_name_select().options if len(x.text)>1]
    
    for code in codes:
        labeler_select().select_by_visible_text(code)
        for program in programs:
            email_flag=0
            program_name_select().select_by_visible_text(program)
            submit = wait.until(EC.element_to_be_clickable((By.ID,'mainForm:btnContinue')))
            submit.click()
            time.sleep(3)
            print('submit clicked')
            #sometimes the site wants to email you when the data is ready, 
            #so switch to that notificaiton and accept if required
            try:
                print('a')
                alert = driver.switch_to.alert
                alert.accept()
                email_flag=1
            except:      
                print('b')
                pass
            #If for some reason the CLD doesn't exist detect the error message
            #add the CLD to the issues list to be sent to the user and move on
            try:
                print('c')
                driver.find_element_by_class_name('errorMsg')
                print('No data for this program')
                driver.refresh()
                continue
                print('d')
            except NoSuchElementException as ex:
                print('e')
                pass
            if email_flag ==0:
                success_flag = 0
                while success_flag ==0:
                    try:
                        wait2.until(EC.element_to_be_clickable((By.XPATH,'//p[contains(text(),"We apologize for the inconvenience and appreciate your patience.")]')))
                        driver.back()
                        submit = wait.until(EC.element_to_be_clickable((By.ID,'mainForm:btnContinue')))
                        submit.click()
                        print('f')
                    except TimeoutException as ex:
                        print('g')
                        success_flag=1
                        pass
                while 'claimdetails.xls' not in os.listdir():
                    time.sleep(1)
                path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\'+'Claims\\Tennessee\\'+program+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
                if os.path.exists(path)==False:
                    os.makedirs(path)              
                else:
                    pass
                new_name ='TN_{}_{}Q{}_{}.xls'.format(program,qtr,yr,code)
                shutil.move('claimdetails.xls',path+new_name)

def main():
    pull()
    
if __name__=='__main__':
    main()




































