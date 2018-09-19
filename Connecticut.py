# -*- coding: utf-8 -*-
"""
Created on Thu Jul 19 08:47:24 2018

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
time_stuff = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx', sheet_name = 'Year-Qtr',use_cols='A:B')
yr = time_stuff.iloc[0,0]
qtr = time_stuff.iloc[0,1]
login_credentials = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Connecticut', usecols=[0,1],dtype='str')
login_credentials = login_credentials[login_credentials['Username']!='nan']
user = list(login_credentials.Username)
password = list(login_credentials.Password)
chromeOptions = webdriver.ChromeOptions()
prefs = {'download.default_directory':'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder'}
chromeOptions.add_experimental_option("prefs",prefs)

driver = webdriver.Chrome(chrome_options = chromeOptions, executable_path=r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\chromedriver.exe')
os.chdir('O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder')
for file in os.listdir():
    os.remove(file)
yq = str(yr)+str(qtr)
mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Connecticut', usecols=[2,3],dtype='str')
mapper = dict(zip(mapper['CT Code'],mapper['Lilly Code']))
wait = WebDriverWait(driver,19)
for USER, PW in zip(user,password):
    driver.get('https://www.ctdssmap.com/CTPortal/Provider/Secure%20Site/tabId/56/Default.aspx')
    user_name = driver.find_element_by_xpath('//*[@id="dnn_ctr383_LoginPage_SearchPage_dataPanel_ctl01_ctl11_userName_mb_userName"]')
    user_name.send_keys(USER)
    pass_word = driver.find_element_by_xpath('//*[@id="dnn_ctr383_LoginPage_SearchPage_dataPanel_ctl01_ctl12_password_mb_password"]')
    pass_word.send_keys(PW)
    login_button = driver.find_element_by_xpath('//*[@id="dnn_ctr383_LoginPage_SearchPage_dataPanel_ctl01_ctl13_LoginButton"]')
    login_button.click()
    '''
    The next section of code is built to detect if the password is expired or not.
    If it is an email will be sent to a designated person notifying them that the password
    is no longer good.
    '''
    try:
        new_password = WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="dnn_ctr383_LoginPage_SearchPage_changePassword_ctl02_ctl02_newPassword_mb_newPassword"]')))
        subject = 'Connecticut Healthcare Report Portal Password is Expired'
        body = 'While attempting to pull reports from the Connecticut Dept of Social Services the bot was notified the password is expired.  Please go to site, change the password, and update the parameter file.'
        recipient = 'burtner_abt_alec@lilly.com'
        base = 0x0
        obj = win32com.client.Dispatch('Outlook.Application')
        newMail = obj.CreateItem(base)
        newMail.Subject = subject
        newMail.Body = body
        newMail.To = recipient
        newMail.display()
        newMail.Send()
    except:
        pass

        provider_tab = wait.until(EC.element_to_be_clickable((By.XPATH,'//a[@title="Provider"]')))
        ActionChains(driver).move_to_element(provider_tab).perform()    
        secure_site = driver.find_element_by_xpath('//a[@title="Secure Site"]')
        secure_site.click()     
        user_name = driver.find_element_by_xpath('//*[@id="dnn_ctr383_LoginPage_SearchPage_dataPanel_ctl01_ctl11_userName_mb_userName"]')
        user_name.send_keys(USER)
        pass_word = driver.find_element_by_xpath('//*[@id="dnn_ctr383_LoginPage_SearchPage_dataPanel_ctl01_ctl12_password_mb_password"]')
        pass_word.send_keys(PW)
        login_button = driver.find_element_by_xpath('//*[@id="dnn_ctr383_LoginPage_SearchPage_dataPanel_ctl01_ctl13_LoginButton"]')
        login_button.click()
        
        #now find the trade files tab and click it, select the 
        #drug rebate file transfer option from the drop down
        #and then click the search button
        trade_file_tab = driver.find_element_by_xpath('//a[@title="Trade Files"]')
        ActionChains(driver).move_to_element(trade_file_tab).perform()
        downloads = driver.find_element_by_xpath('//li//a[text()="Download"]')    
        downloads.click()
        transaction_type = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="dnn_ctr416_DownloadSearchPage_SearchPage_CriteriaPanel_ctl01_ctl00_TransactionType"]')))
        transaction_type_select = Select(transaction_type)     
        transaction_type_select.select_by_visible_text('Drug Rebate File Transfer')     
        search = driver.find_element_by_xpath('//*[@id="dnn_ctr416_DownloadSearchPage_SearchPage_CriteriaPanel_ctl01_ctl00_SearchButton"]')     
        search.click()     
        wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="dnn_ctr416_DownloadSearchPage_SearchPage_CriteriaPanel_ctl01_ctl01_ClearButton"]')))
        claims_to_get = dict.fromkeys(mapper.values())
        #Continue flag is 1 while there are files on the page that
        #have the current YYYYQ in their title, when there are no longer
        #current files to download the loop will break
        continue_flag = 1
        while continue_flag ==1:
            #Get the links on each page
            links =  driver.find_elements_by_xpath('//table[@class="iC_DataListContainer"]//tbody//tr[position()>2][contains(@class,"iC_Data")]//td[3]')
            names = []
            xpaths = []
            for i in range(len(links)):
                n = str(i+3)
                xpath = '//table[@class="iC_DataListContainer"]//tbody/tr[%s]/td[3]'%(n)
                xpaths.append(xpath)
                name = driver.find_element_by_xpath('//table[@class="iC_DataListContainer"]//tbody/tr[%s]/td[2]'%(n))
                name = name.text
                names.append(name)
    
            names =  driver.find_elements_by_xpath('//table[@class="iC_DataListContainer"]//tbody//tr[position()>2][contains(@class,"iC_Data")]//td[2]')
            names = [x.text for x in names]
            #loop through the links
            for j, x in enumerate(xpaths):
                item = driver.find_element_by_xpath(x)
                xxx = item.text
                if yq not in xxx:
                    continue
                else:
                    pass
                program = xxx[6:8].lower()
                program = mapper[program]
                label_code = xxx[-9:-4]
                ext = names[j][-4:]
                if ext == '.dat':
                    ext = '.txt'
                else:
                    pass
                item.click()
                #pause for the file to download
                while names[j] not in os.listdir():
                    time.sleep(1)
                #build the path
                path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\'+'Invoices\\Connecticut\\'+program+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
                if os.path.exists(path)==False:
                    os.makedirs(path)
                else:
                    pass
                ndcs = []
                if ext == '.txt':
                    with open(names[j]) as a:
                        lines = a.readlines()
                        ndcs = [x[6:17] for x in lines]
                    ndcs = list(set(ndcs))
                    claims_to_get.update({program:ndcs})
                else:
                    pass
                file_name = program+'_'+label_code+'_'+str(yr)+'_'+'Q'+str(qtr)+ext
                shutil.move(names[j], path+file_name)
            #At this point all files on the current page are downloaded
            #and moved to the LAN drive, so now it moves to the next page
            #and checks the files.  If any file has the current YYYYQ in it 
            #it will be downloaded
            next_page = driver.find_element_by_xpath('//a[@class="Next"]')
            next_page.click()
            soup = BeautifulSoup(driver.page_source,'html.parser')
            files = soup.find_all('td') 
            if any(map((lambda x: yq in x),[x.text for x in files]))==True:
                continue_flag = 1
            else:
                continue_flag =0
             
    #Now moving onto CLD
    cld_page = driver.find_element_by_xpath('//div//ul//li//a[@title="Claim Level Detail"][@href="/CTPortal/Trade%20Files/Claim%20Level%20Detail/tabId/85/Default.aspx"]')
    trade_files = driver.find_element_by_xpath('//a[@title="Trade Files"]')
    ActionChains(driver).move_to_element(trade_files).move_to_element(cld_page).click().perform()      
    label_code = USER[-5:]
    #Currently only care about two programs
    programs_we_care_about = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Connecticut', usecols='E',dtype='str')
    programs_we_care_about = list(programs_we_care_about['CLD Programs'])

    for key in programs_we_care_about:
        master_frame = pd.DataFrame()
        for ndc in claims_to_get[key]:
            success_flag = 0
            ndc_box = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="dnn_ctr418_ClaimLevelDetailPage_SearchPage_CriteriaPanel_ctl01_ctl00_NDC_mb_NDC"]')))
            ndc_box.clear()
            ndc_box.send_keys(ndc)
            year_quarter_box = driver.find_element_by_xpath('//*[@id="dnn_ctr418_ClaimLevelDetailPage_SearchPage_CriteriaPanel_ctl01_ctl01_InvoiceCycle_mb_InvoiceCycle"]')
            year_quarter_box.clear()
            year_quarter_box.send_keys(yq)
            invoice_type_select = driver.find_element_by_xpath('//*[@id="dnn_ctr418_ClaimLevelDetailPage_SearchPage_CriteriaPanel_ctl01_ctl02_InvoiceType"]')
            invoice_type_select_select = Select(invoice_type_select)
            invoice_type_select_select.select_by_visible_text(key)
            while success_flag ==0 :
                try:
                    search_button = driver.find_element_by_xpath('//*[@id="dnn_ctr418_ClaimLevelDetailPage_SearchPage_CriteriaPanel_ctl01_ctl02_SearchButton"]')
                    search_button.click()
                    download_button =  wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="dnn_ctr418_ClaimLevelDetailPage_SearchPage_DownloadButton"]')))
                    download_button.click()
                    while 'ClaimLevelDetail.csv' not in os.listdir():
                        time.sleep(1)           
                    flag = 0
                    while flag == 0:
                        try:
                            a = open('ClaimLevelDetail.csv')
                            flag=1
                            a.close()
                        except PermissionError as ex:
                            flag = 0
                            pass
                    temp = pd.read_csv('ClaimLevelDetail.csv',usecols=list(range(16)),skiprows=8,engine='python')
                    meta_data = pd.read_csv('ClaimLevelDetail.csv',usecols=[0,1],nrows=8,header=None,names=['Field','Value'],engine='python')
                    temp['NDC'] = ''.join(meta_data.Value[2].split('-'))
                    temp['ICN'] = temp.ICN.str.replace('=','').str.replace('"','')
                    temp['Inv_Qtr'] = meta_data['Value'][1]
                    master_frame = master_frame.append(temp)
                    os.remove('ClaimLevelDetail.csv')
                    #if os.path.exists(path)==False:
                        #os.makedirs(path)
                    #else:
                        #pass
                    #shutil.move('ClaimLevelDetail.csv',path+file_name)
                    success_flag=1
                except TimeoutException as ex:
                    time.sleep(2)
                    cld_tab = driver.find_element_by_xpath('//*[@id="dnn_ctl00_SubSystemMenu"]/li[3]/a')
                    cld_tab.click()
                    ndc_box = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="dnn_ctr418_ClaimLevelDetailPage_SearchPage_CriteriaPanel_ctl01_ctl00_NDC_mb_NDC"]')))
                    ndc_box.clear()
                    ndc_box.send_keys(ndc)
                    year_quarter_box = driver.find_element_by_xpath('//*[@id="dnn_ctr418_ClaimLevelDetailPage_SearchPage_CriteriaPanel_ctl01_ctl01_InvoiceCycle_mb_InvoiceCycle"]')
                    year_quarter_box.clear()
                    year_quarter_box.send_keys(yq)
                    invoice_type_select = driver.find_element_by_xpath('//*[@id="dnn_ctr418_ClaimLevelDetailPage_SearchPage_CriteriaPanel_ctl01_ctl02_InvoiceType"]')
                    invoice_type_select_select = Select(invoice_type_select)
                    invoice_type_select_select.select_by_visible_text(key)
        path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\'+'Claims\\Connecticut\\'+key+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
        if os.path.exists(path)==False:
            os.makedirs(path)
        else:
            pass
        file_name = key+'_'+label_code+'_'+str(yr)+'_Q'+str(qtr)+'.csv'
        master_frame.to_csv('master_table.csv', index=False)
        shutil.move('master_table.csv',path+file_name)

