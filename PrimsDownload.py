# -*- coding: utf-8 -*-
"""
Created on Tue Aug 21 09:46:33 2018

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
from selenium.common.exceptions import NoSuchElementException, TimeoutException, NoAlertPresentException
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

def prims_download():
    os.chdir('C:/Users/')
    statesII = {
        'AK': 'Alaska',
        'AL': 'Alabama',
        'AR': 'Arkansas',
        'AS': 'American Samoa',
        'AZ': 'Arizona',
        'CA': 'California',
        'CO': 'Colorado',
        'CT': 'Connecticut',
        'DC': 'District of Columbia',
        'DE': 'Delaware',
        'FL': 'Florida',
        'GA': 'Georgia',
        'GU': 'Guam',
        'HI': 'Hawaii',
        'IA': 'Iowa',
        'ID': 'Idaho',
        'IL': 'Illinois',
        'IN': 'Indiana',
        'KS': 'Kansas',
        'KY': 'Kentucky',
        'LA': 'Louisiana',
        'MA': 'Massachusetts',
        'MD': 'Maryland',
        'ME': 'Maine',
        'MI': 'Michigan',
        'MN': 'Minnesota',
        'MO': 'Missouri',
        'MP': 'Northern Mariana Islands',
        'MS': 'Mississippi',
        'MT': 'Montana',
        'NA': 'National',
        'NC': 'North Carolina',
        'ND': 'North Dakota',
        'NE': 'Nebraska',
        'NH': 'New Hampshire',
        'NJ': 'New Jersey',
        'NM': 'New Mexico',
        'NV': 'Nevada',
        'NY': 'New York',
        'OH': 'Ohio',
        'OK': 'Oklahoma',
        'OR': 'Oregon',
        'PA': 'Pennsylvania',
        'PR': 'Puerto Rico',
        'RI': 'Rhode Island',
        'SC': 'South Carolina',
        'SD': 'South Dakota',
        'TN': 'Tennessee',
        'TX': 'Texas',
        'UT': 'Utah',
        'VA': 'Virginia',
        'VI': 'Virgin Islands',
        'VT': 'Vermont',
        'WA': 'Washington',
        'WI': 'Wisconsin',
        'WV': 'West Virginia',
        'WY': 'Wyoming'
                        }
        #Open the webdriver, define the wait function, and get through the login page

    chromeOptions = webdriver.ChromeOptions()
    prefs = {'download.default_directory':'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder'}
    chromeOptions.add_experimental_option("prefs",prefs)
    driver = webdriver.Chrome(chrome_options = chromeOptions)
    os.chdir('O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder')
    for file in os.listdir():
        os.remove(file)
    driver.implicitly_wait(30)
    wait = WebDriverWait(driver,15)
    time_stuff = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx', sheet_name = 'Year-Qtr',use_cols='A:B')
    yr = time_stuff.iloc[0,0]
    qtr = time_stuff.iloc[0,1]
    login_credentials = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Prims', usecols='A,B',dtype='str')
    username = login_credentials.iloc[0,0]
    password = login_credentials.iloc[0,1]
    driver.get('https://www.primsconnect.molinahealthcare.com/_layouts/fba/primslogin.aspx?ReturnUrl=%2f_layouts%2fAuthenticate.aspx%3fSource%3d%252FSitePages%252FHome%252Easpx&Source=%2FSitePages%2FHome%2Easpx')
    driver = driver
    i_accept = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_RadCheckBoxAccept"]/span[1]')))
    i_accept.click()
    flex_mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Programs', usecols='B,C,D,E',dtype='str',names=['state','flex_id','state_id','state_name'])
    user_name = driver.find_element_by_xpath('//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_txtUserName"]')
    user_name.send_keys(username)
    pass_word = driver.find_element_by_xpath('//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_txtPassword"]')
    pass_word.send_keys(password)
    login = driver.find_element_by_xpath('//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_btnLogin_input"]')
    login.click()          
    #Now inside the webpage, begin selection process
    submit_request = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_radLnkSubmitRequest_input"]')))
    submit_request.click()    
    
    yq2 = '{}Q{}'.format(qtr,yr)
    yq3 = '{}-Q{}'.format(yr,qtr)
    #Make the program to state dictionaries
    soup = BeautifulSoup(driver.page_source,'html.parser')
    lists = soup.find_all('ul',attrs={'class':'rcbList'})    
    states = [x.text for x in lists[0]]    
    state_programs = {}
    
    
    
    #have to select the state to get the state programs to populate
    '''
    The below block of code is creating the state: programcode:program name dictionary
    to create the filenames for after download
    '''
    
    
    for state in states:
        drop_down = driver.find_element_by_xpath('//*[@id="ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_StateDropDown_Input"]')
        if drop_down.get_attribute('value')==state:
            pass
        else:
            xpath = '//div[contains(@id,"ctl00_StateDropDown_DropDown")]//li[text()="{}"]'.format(state)
            state_to_select = driver.find_element_by_xpath(xpath)
            ActionChains(driver).move_to_element(drop_down).click().pause(1).move_to_element(state_to_select).click().perform()
            wait.until(EC.staleness_of(drop_down))
        soup = BeautifulSoup(driver.page_source,'html.parser')
        lists2 = soup.find_all('ul',attrs={'class':'rcbList'})
        programs = [x.text.split('-') for x in lists2[1]]
        codes = [x[0].strip() for x in programs]
        name = ['-'.join(x[1:]) for x in programs]
        programs = dict(zip(codes,name))
        state_programs.update({state:programs})        
    driver.back()
    try:
        alert = driver.switch_to.alert
        alert.accept()
    except NoAlertPresentException as ex:
        pass
    '''
    The below block of code crawls through available download pages and 
    downloads the data, renames it, and moves it to the appropriate directory
    '''
    
    #Chane the number of reports per page
    number_per_page = driver.find_element_by_xpath('//*[@id="ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_radGridRequestSummary_ctl00_ctl03_ctl01_ChangePageSizeTextBox"]')
    number_per_page.clear()    
    number_per_page.send_keys('10000')    
    inter=driver.find_element_by_xpath('//*[@id="ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_radGridRequestSummary_ctl00_ctl03_ctl01_ChangePageSizeLinkButton"]')
    inter.click()
    #Get the pages downloads are on
    pages = lambda: driver.find_elements_by_xpath('//div[@class="rgWrap rgNumPart"]/a')     
    for i,page in enumerate(pages()):
        p = pages()[i]
        pages()[i].click()
        soup = BeautifulSoup(driver.page_source,'html.parser')
        table = soup.find('table',attrs={'class':'rgMasterTable'})
        body = table.find_all('tbody')[1]            
        data = [x.text.strip() for x in body.find_all('td')]
        data = np.asarray(data)
        data = data.reshape(-1,9)
        data = data[data[:,-1]=='Download']
        data = pd.DataFrame(data,columns=['report_id','manufacturer','state','date_requested','type','status','file_name','date_complete','download_link'])
        data['state_id'] = data['file_name'].str.rsplit('_',0).str[-1]
        data['state_id'] = data['state_id'].str.rsplit('.').str[0]
        data= pd.merge(data,flex_mapper,how='left',on=['state','state_id'])
        data = data.fillna('no_flex_id')
        data['labeler'] = data['file_name'].str.split('_').str[2]
        data = data[data['type']=='Invoice']
        data = data.drop_duplicates(subset=['file_name']).reset_index(drop=True)
        
        ndc_list = []
        for i in range(len(data)):
            state = data.loc[i,'state']
            program = data.loc[i,'flex_id']
            file_type = data.loc[i,'file_name'][-4:].lower()
            labeler= data.loc[i,'labeler']
            if program =='no_flex_id':
                program = state_programs[statesII[state]][data.loc[i,'state_id']].strip()
            path ='O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Invoices\\{}\\{}\\{}\\Q{}\\'.format(statesII[state],program,yr,qtr)
            file_name = '_'.join([state,program,yq2,labeler])+file_type
            if os.path.exists(path)==False:
                os.makedirs(path)
            xpath = '//tr/td[text()="{}"]/following-sibling::td/span[contains(@id,"_lnkDownload")]'.format(data.loc[i,'file_name'])
            link = driver.find_element_by_xpath(xpath)
            link.click()
            while data.loc[i,'file_name'] not in os.listdir():
                time.sleep(1)
            if file_type =='.txt':
                read_flag =0
                while read_flag ==0:
                    try:
                        with open(data.loc[i,'file_name']) as f:
                            lines = f.readlines()
                            menu_item = '  -  '.join([data.loc[i,'state_id'],state_programs[statesII[state]][data.loc[i,'state_id']].strip()])
                            ndcs = (state,menu_item,list(set([x[6:17] for x in lines])))
                            ndc_list.append(ndcs)
                            read_flag=1
                    except PermissionError as ex:
                        pass
            shutil.move(data.loc[i,'file_name'],path+file_name)
    
    #Have the state : NDC tuples, move onto getting CLD
    submit_request = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_radLnkSubmitRequest_input"]')))
    submit_request.click()    
    
    for report in ndc_list:
        
        state_drop_down = driver.find_element_by_xpath('//input[contains(@name,"$ctl00$StateDropDown")]')
        if state_drop_down.get_attribute('value')==statesII[report[0]]:
            pass
        else:
            state_to_select = driver.find_element_by_xpath('//div[contains(@id,"a_ctl00_StateDropDown_DropDown")]//li[contains(text(),"{}")]'.format(statesII[report[0]]))
            ActionChains(driver).move_to_element(state_drop_down).click().pause(1).move_to_element(state_to_select).click().perform()
            wait.until(EC.staleness_of(state_drop_down))
            state_drop_down = driver.find_element_by_xpath('//input[contains(@name,"$ctl00$StateDropDown")]')
            while state_drop_down.get_attribute('value') !=statesII[report[0]]:
                time.sleep(1)
        selected_flag = 0
        while selected_flag ==0:
            try:
                program_drop_down = lambda: driver.find_element_by_xpath('//input[contains(@name,"a$ctl00$ProgramDropDown")]')
                program_drop_down().click()
                time.sleep(1)
                xpath = '//div[@id="ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_ProgramDropDown_DropDown"]//li[text()="{}"]'.format(report[1])
                program_to_select = driver.find_element_by_xpath(xpath)
                if program_drop_down().get_attribute('value')==report[1]:
                    selected_flag=1
                else:
                    program_to_select.click()
                    wait.until(EC.staleness_of(program_to_select))
                    if program_drop_down().get_attribute('value')==report[1]:
                        selected_flag=1

            except NoSuchElementException as ex:
                pass
        dates_acquired = 0
        from_q = driver.find_element_by_xpath('//input[contains(@name,"$ctl00$FYearQuarterDropDown")]')
        from_q.click()
        time.sleep(1)
        while dates_acquired==0:
            dates =[x.text for x in driver.find_elements_by_xpath('//div[@id="ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_FYearQuarterDropDown_DropDown"]//li[contains(text(),"Q")]')]
            if len(dates[0])==0:
                pass
            else:
                dates_acquired=1
        if any(yq3 in x for x in dates)==False:
            continue
            #pass
        else:
            from_q = driver.find_element_by_xpath('//input[contains(@name,"$ctl00$FYearQuarterDropDown")]')
            to_q = driver.find_element_by_xpath('//input[contains(@name,"$ctl00$TYearQuarterDropDown")]')
            current_qtr = driver.find_element_by_xpath('//div[contains(@id,"a_ctl00_FYearQuarterDropDown_DropDown")]//li[text()="{}"]'.format(yq3))
            current_qtr2 = driver.find_element_by_xpath('//div[contains(@id,"a_ctl00_TYearQuarterDropDown_DropDown")]//li[text()="{}"]'.format(yq3))
            ActionChains(driver).move_to_element(from_q).click().pause(1).click(current_qtr).move_to_element(to_q).click().pause(1).move_to_element(current_qtr2).click().perform()
            
        ndcs = ','.join(report[2])
        ndc_box = driver.find_element_by_xpath('//input[contains(@name,"$ctl00$NDCTextBox")]')
        ndc_box.send_keys(ndcs)
        submit = driver.find_element_by_xpath('//input[@type="submit"][@value="Submit"]')
        submit.click()
        wait.until(EC.staleness_of(submit))
        ndc_box = driver.find_element_by_xpath('//input[contains(@name,"$ctl00$NDCTextBox")]')
        ndc_box.clear()
    driver.close()

def main():

    prims_download()
    #send_message(subject=,body=,to='burtner_abt_alec@lilly.com')
    
if __name__=='__main__':
    main()
