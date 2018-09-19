# -*- coding: utf-8 -*-
"""
Created on Mon Jul 16 10:08:13 2018

@author: C252059
"""
# Load all applicable modules
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



states = {
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
'WY': 'Wyoming',
'Absolute' : 'South Carolina',
'BlueChoice' :'South Carolina',
'First' :'South Carolina',
'Unison' :'Ohio'
}
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
login_credentials = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='DrugRebate.com', usecols='A,B',dtype='str')
username = login_credentials.iloc[0,0]
password = login_credentials.iloc[0,1]
#This creates the webdriver instance, navigates to the login page, and logs in
wait = WebDriverWait(driver,10)

driver.get('https://www.drugrebate.com/RebateWeb/login.do')
driver.maximize_window()
user_name = driver.find_element_by_xpath('//*[@id="username"]')
user_name.send_keys(username)
pass_word = driver.find_element_by_xpath('//*[@id="password"]')
pass_word.send_keys(password)
login_button= driver.find_element_by_xpath('//*[@id="loginBtn"]')
login_button.click()
##Inside page, got to reports

menu = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="SearchInvoices"]/span/a')))
menu.click()


#Leave labeler code blank and it will pull all of them
quarter = wait.until(EC.element_to_be_clickable((By.XPATH,'//select[@id = "startQq"]')))
quarter.send_keys(str(qtr))    
year = driver.find_element_by_xpath('//input[@id = "qtrYear"]')    
year.send_keys(str(yr))

submit_button = driver.find_element_by_xpath('//input[@value="Submit"]')
submit_button.click()    

#The above block of code got all of the invoics for all labels.
#Now we have to create the reference dictionary to name files after they
#are downloaded and then download the files.

soup = BeautifulSoup(driver.page_source,'html.parser')
data = soup.find_all('td')
columns = soup.find_all('th')
columns = [x.text for x in columns]
data = [x.text for x in data]
data = np.asarray(data)
data = data.reshape(-1,9)
data_asframe = pd.DataFrame(data,columns=columns)
grouped = data_asframe.groupby(['Payer','Program','Labeler'])

file_dict = {k:v for k,v in zip(data[:,1],data[:,2:])}

#This built the file naming dictionary, keys are teh invoice number

all_invoices = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@name="downloadWhat"][@value="all"]')))
all_invoices.click()    
download_button = driver.find_element_by_xpath('//input[@value="Download Invoices"]')
download_button.click()

try:
    alert = driver.switch_to.alert
    alert.accept()
except NoSuchElementException as ex:
    pass
while any(map((lambda x: 'invoice' in x),os.listdir()))==False:
    time.sleep(1)
grouped.groups
missing = {}
formats = list(set(data_asframe.Format))
for group in grouped.groups:
    print(len(grouped.get_group(group)))
    print(grouped.get_group(group)['Format'])
    missing_formats = []
    for item in formats:
        if item not in list(grouped.get_group(group)['Format']):
            missing_formats.append(item)
        else:
            pass
        if len(missing_formats)==0:
            pass
        else:
            missing.update({group:missing_formats})
programs = list(set(data_asframe.Program))
subject = 'Missing Data Formats for Conduet Wesbite (DrugRebate.com)'
body = 'The following programs had files pulled\n'+'\n'.join(programs)+'\n\n'+\
'The below programs had missing data formats\n'
body2 = pprint.pformat(missing, width = 120, indent=4)
recipient = 'burtner_abt_alec@lilly.com'
base = 0x0
obj = Dispatch('Outlook.Application')
newMail = obj.CreateItem(base)
newMail.Subject = subject
newMail.Body = body+'\n'+body2+'\nPlease contact the appropriate state agency for the missing documents.\n'\
'Beep Boop, I am a robot.  This output was generated by the Conduet utility script.'
newMail.To = recipient
newMail.display()
newMail.Send()
            
    

'''
The below snippet goes to the downloads and unzips all the downloads from the above loop.
After unzipping the file the zip file gets deleted. 
'''
flag = 0
while flag==0:
    file = os.listdir()[0]
    if file[-3:] != 'zip':
        pass
    else:
        flag=1
zips = os.listdir()
zips = [file for file in zips if 'zip' in file]

for file in zips:
    flag = 0
    while flag ==0:
        try:
            with zipfile.ZipFile(file,'r') as zip_ref:
                zip_ref.extractall()
            os.remove(file)
            flag=1
        except PermissionError as ex:
            pass
#now that the files have been unzipped they will be renamed,
# and moved to the LAN folder
files = os.listdir()
for file in files:
    key = file.split('.')[0]
    file_data = file_dict[key]
    if file.split('.')[1][:2] in states.keys():
        state = states[file.split('.')[1][:2]]
    elif file_data[1].split(' ')[0] in states.values():
        state = file_data[1].split(' ')[0]
    else:
        state = 'New Mexico'
    if file_data[4] =='NCPDP Claims Own':
        file_type = 'Claims'
    else:
        file_type = 'Invoices'
    ext = file[-4:]
    path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\'+file_type+'\\'+state+'\\'+file_data[2]+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
    file_name = file_data[2]+'_'+file_data[3]+'_'+file_data[0]+ext
    if os.path.exists(path)==False:
        os.makedirs(path)
    else:
        pass
    shutil.move(file,path+file_name)

    
    
    
    
    
    
    