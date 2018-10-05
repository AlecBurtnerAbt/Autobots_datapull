# -*- coding: utf-8 -*-
"""
Created on Fri Jul 20 08:14:27 2018

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

def download_reports():
    os.chdir('C:/Users/')
      
    time_stuff = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx', sheet_name = 'Year-Qtr',use_cols='A:B')
    yr = time_stuff.iloc[0,0]
    qtr = time_stuff.iloc[0,1]
    login_credentials = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='California', usecols='F,G',dtype='str')
    username = login_credentials.iloc[0,0]
    password = login_credentials.iloc[0,1]
    chromeOptions = webdriver.ChromeOptions()
    prefs = {'download.default_directory':'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder',
             'plugins.always_open_pdf_externally':True,
             'download.prompt_for_download':False}
    chromeOptions.add_experimental_option('prefs',prefs)
    program_mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='California', usecols=[1,2,3],dtype='str')
    mapper = dict(zip(program_mapper['Code on CA Invoice'],program_mapper['Contract ID in MRB']))
    driver = webdriver.Chrome(chrome_options = chromeOptions, executable_path=r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\chromedriver.exe')
    os.chdir('O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder')
    for file in os.listdir():
        os.remove(file)

    yq = str(yr)+str(qtr)
    yq2 = str(qtr)+'Q'+str(yr)
    #navigate to the drug rebate invoice page

    
    #get the three labeler codes.  Will have to update if labeler codes change
    lilly_code = lambda : wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="middle_column"]/div[2]/table/tbody/tr/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[2]/td[2]/a')))
    dista_code = lambda : wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="middle_column"]/div[2]/table/tbody/tr/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[3]/td[2]/a')))
    imclone_code = lambda :wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="middle_column"]/div[2]/table/tbody/tr/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[4]/td[2]/a')))
    
    codes = [lilly_code,dista_code,imclone_code]
    '''
    This block of code downloads all of the prepared reports.  The reports come in a .gz file
    and have to be decompressed, this happens after the download in the next loop.
    '''
    for user, password in zip(login_credentials.Username[:2],login_credentials.Password[:2]):      
        driver.get('https://www.medi-cal.ca.gov/')
        transaction_tab = driver.find_element_by_xpath('//a[text()="Transactions"]')
        transaction_tab.click()
        wait = WebDriverWait(driver,10)
        user_name = wait.until(EC.element_to_be_clickable((By.ID,'UserID')))
        user_name.send_keys(user)
        pass_word = driver.find_element_by_id('UserPW')
        pass_word.send_keys(password)
        submit_button = driver.find_element_by_id('cmdSubmit')
        submit_button.click()             
        drug_rebate = driver.find_element_by_xpath('//*[@id="tabpanel_1_sublist"]/li/a')
        drug_rebate.click()
        
        for code in codes:
            code().click()
            retrieve = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="middle_column"]/div[2]/table/tbody/tr/td[2]/table/tbody/tr[3]/td/a[2]/b')))
            retrieve.click()               
            wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="left_column"]/div[1]/a/img')))     
            soup2 = BeautifulSoup(driver.page_source,'html.parser') 
        
            bodies = soup2.find_all('tbody')
            body = bodies[2]
            rows = body.find_all('tr')
            data = body.find_all('td')
            data = [x.text for x in data]
            data = np.asarray(data)
            array_length = int(len(data)/3)
            data = data.reshape(-1,3)
            links = [x[0] for x in data if 'Completed' in x[1] and str(yr)+str(qtr)==x[0].split('_')[-2]]               
            links = ["".join(x.split()) for x in links]              
            for link in range(len(links)):
                xpath = "//a[contains(text(),'"+links[link]+"')]"
                DL_link = driver.find_element_by_xpath(xpath)
                DL_link.click()            
                while links[link] not in os.listdir():
                    time.sleep(1)
                    
            driver.get(r'https://rais.medi-cal.ca.gov/drug/DrugLablr.asp')
        exit_link = driver.find_element_by_xpath('//a[text()="Exit"]')
        exit_link.click()
        
            
        '''
        This is the loop that goes through the downloaded .gz files, unzips them, renames them
        to the file format, makes them a text file, and then deletes the .gz file
        '''
    files = os.listdir()
    num_reports = len(files)
    for file in files:
        prog_code = file.split('_')[-1].split('.')[0]
        prog = mapper[prog_code]
        label_code = file.split('_')[2]
        request_number = file.split('_')[1]
        new_name =  'CA_{}_{}Q{}_{}_{}.txt'.format(prog,qtr,yr,label_code,request_number)
        unzipped_name = file[:-3]
        try:
            with gzip.open(file,'rt') as ref:
                content = ref.read()
                text_file = open(unzipped_name,'w')
                text_file.write(content)
                text_file.close()
                ref.close()
        except:
            pass
        path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Claims\\California\\'+prog+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
        if os.path.exists(path)==False:
            os.makedirs(path)
        else:
            pass
        shutil.move(unzipped_name,path+new_name)
        os.remove(file)
    driver.close()
    return num_reports
def main():
    num_reports = download_reports()
    subject = 'California Step Two'
    body = 'All CLD data requested has been downloaded, there were {} files downloaded.'.format(num_reports)
    to_address = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Notification Address', usecols='A',dtype='str',names=['Email'],header=None).iloc[0,0]
    send_message(subject=subject,body=body,to=to_address)
if __name__ == '__main__':
    main()    