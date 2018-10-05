# -*- coding: utf-8 -*-
"""
Created on Mon Oct  1 12:28:42 2018

@author: C252059
"""

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
import multiprocessing as mp



def prims_download():
    def login_proc(driver):
        i_accept = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_RadCheckBoxAccept"]/span[1]')))
        i_accept.click()
        flex_mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Programs', usecols='B,C,D,E',dtype='str',names=['state','flex_id','state_id','state_name'])
        user_name = driver.find_element_by_xpath('//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_txtUserName"]')
        user_name.send_keys(username)
        pass_word = driver.find_element_by_xpath('//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_txtPassword"]')
        pass_word.send_keys(password)
        login = driver.find_element_by_xpath('//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_btnLogin_input"]')
        login.click()
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
    os.chdir('C:/Users/')
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
    login_proc(driver)
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
    #The below block of code is creating the state: programcode:program name dictionary
    #to create the filenames for after download    
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
    #The below block of code crawls through available download pages and 
    #downloads the data, renames it, and moves it to the appropriate directory   
    xpaths = []
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
        data = data.fillna('no_flex_id')
        data['labeler'] = data['file_name'].str.split('_').str[2]
        data_copy = data
        data = data[data['type']=='Claims']
        data = data.drop_duplicates(subset=['file_name']).reset_index(drop=True)
        
      
        for i in range(len(data)):
            state = data.loc[i,'state']
            state_full = statesII[state]
            program_identifier = data.loc[i,'file_name'].split('_')[7]
            program = state_programs[state_full][program_identifier]
            file_type = data.loc[i,'file_name'][-4:].lower()
            labeler= data.loc[i,'labeler']
            file_name = '_'.join([state,program,yq2,labeler])+file_type
            xpath = '//tr/td[text()="{}"]/following-sibling::td/span[contains(@id,"_lnkDownload")]'.format(data.loc[i,'file_name'])
            xpaths.append((xpath,data.loc[i,'file_name']))
    driver.close()
    return xpaths, state_programs, username, password, statesII, yr, qtr


def make_chunks(list_of_files):
    #Break the information for each report down into 
    import math
    n = math.ceil(len(list_of_files)/3)
    chunks = [list_of_files[x:x+n] for x in range(0,len(list_of_files),n)]
    return chunks
         
def get_cld_reports(xpaths,username,password):
    def login_proc():
        i_accept = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_RadCheckBoxAccept"]/span[1]')))
        i_accept.click()
        flex_mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Programs', usecols='B,C,D,E',dtype='str',names=['state','flex_id','state_id','state_name'])
        user_name = driver.find_element_by_xpath('//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_txtUserName"]')
        user_name.send_keys(username)
        pass_word = driver.find_element_by_xpath('//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_txtPassword"]')
        pass_word.send_keys(password)
        login = driver.find_element_by_xpath('//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_btnLogin_input"]')
        login.click()
    os.chdir('C:/Users/')
    chromeOptions = webdriver.ChromeOptions()
    prefs = {'download.default_directory':'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder'}
    chromeOptions.add_experimental_option("prefs",prefs)
    driver = webdriver.Chrome(chrome_options = chromeOptions)
    driver.implicitly_wait(30)
    wait = WebDriverWait(driver,15)
    os.chdir('O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder')  
    driver.get('https://www.primsconnect.molinahealthcare.com/_layouts/fba/primslogin.aspx?ReturnUrl=%2f_layouts%2fAuthenticate.aspx%3fSource%3d%252FSitePages%252FHome%252Easpx&Source=%2FSitePages%2FHome%2Easpx')
    in_flag = 0
    counter = 0
    while in_flag ==0:
        login_proc()
        try: 
            driver.find_element_by_xpath('//input[@value="Submit Request"]')
            in_flag=1
            break
        except TimeoutException as ex:
            pass                
        driver.get('https://www.primsconnect.molinahealthcare.com/_layouts/fba/primslogin.aspx?ReturnUrl=%2f_layouts%2fAuthenticate.aspx%3fSource%3d%252FSitePages%252FHome%252Easpx&Source=%2FSitePages%2FHome%2Easpx')
        counter +=1
        time.sleep(1*counter)
    def expand():
        number_per_page = driver.find_element_by_xpath('//*[@id="ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_radGridRequestSummary_ctl00_ctl03_ctl01_ChangePageSizeTextBox"]')
        number_per_page.clear()    
        number_per_page.send_keys('10000')    
        inter=driver.find_element_by_xpath('//*[@id="ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_radGridRequestSummary_ctl00_ctl03_ctl01_ChangePageSizeLinkButton"]')
        inter.click()
    expand()
    for xpath,file in xpaths:
        success=0

        while success==0:
            try:
                link = driver.find_element_by_xpath(xpath)
                counter = 0
                link.click()
                while file not in os.listdir() and counter<10:
                    time.sleep(1.5)
                    counter+=1
            except counter >9 or TimeoutException as ex:
                pass
            if file in os.listdir():
                success=1
                break
            else:
                try:
                    driver.find_element_by_xpath('//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_btnLogin_input"]')
                    login_proc()
                    expand()
                except NoSuchElementException as exc:
                    pass
    driver.close()
    

def make_files(state_programs,yr,qtr,statesII):
    names = list(set(['{}_{}'.format(x.split('_')[1],x.split('_')[7]) for x in os.listdir()]))    
    frames = {k:pd.DataFrame() for k in names}
    files = [x.split('_') for x in os.listdir()]    
    files = sorted(files,key=lambda x: (x[1],x[7]))
    files = ['_'.join(x) for x in files]
    cld_obtained = []
    to_address = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Notification Address', usecols='A',dtype='str',names=['Email'],header=None).iloc[0,0]

    for file in files:
        state = file.split('_')[1]
        program_code = file.split('_')[7]
        key = '{}_{}'.format(state,program_code)
        if state=='FL':
            skip = 1
        else:
            skip = 3
        temp = pd.read_excel(file,skiprows=skip,skipfooter=1)
        frames[key] = frames[key].append(temp)
        #os.remove(file)
  
    for key in frames.keys():
        state = key.split('_')[0]
        program = state_programs[statesII[state]][key.split('_')[1]]
        path ='O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Claims\\{}\\{}\\{}\\Q{}\\'.format(statesII[state],program,yr,qtr)
        if os.path.exists(path):
            pass
        else:
            os.makedirs(path)
        file_name = '{}_{}_{}Q{}.csv'.format(state,program,qtr,yr)
        cld_obtained.append(file_name)
        os.chdir(path)
        frames[key].to_csv(file_name,index=False)

       
    #Send message to CMA team notifying them which invoices were  downloaded
    from mail_maker import send_message
    body = "The following CLD files were obtained\n"+'\n'.join(cld_obtained)
    subject = "Florida and West Viriginia Invoices"
    send_message(subject,body,to_address)



def main():
    xpaths, state_programs, username, password, statesII, yr, qtr = prims_download()
    #chunks = make_chunks(xpaths)
    get_cld_reports(xpaths,username,password)
    '''
    Multiprocessing does not work for this site, there
    is some kind of secuirty measure which will cause authentication to fail
    if there is more than one open window downloading files from the site
    
    processes = [mp.Process(target=get_cld_reports,kwargs={'xpaths':chunk,'username':username,'password':password}) for chunk, username,password in zip(chunks,[username for x in range(len(chunks))],[password for x in range(len(chunks))])]
    for p in processes:
        p.start()       
    for p in processes:
        p.join()  
    '''
    make_files(state_programs,yr,qtr,statesII)
    
if __name__=='__main__':
    main()
    
    
    