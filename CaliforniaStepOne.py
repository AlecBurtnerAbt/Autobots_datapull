# -*- coding: utf-8 -*-
"""
Created on Tue Jul 24 08:42:21 2018

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

def cali_one():
    
    os.chdir('C:/Users/')
    
    
    
    #Login step, credentials and time and quarter
    #and brought in from the parameters excel sheet
    time_stuff = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx', sheet_name = 'Year-Qtr',use_cols='A:B')
    yr = str(time_stuff.iloc[0,0])
    qtr = str((time_stuff.iloc[0,1]))
    login_credentials = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='California', usecols='F,G',dtype='str')
    username = login_credentials.iloc[0,0]
    password = login_credentials.iloc[0,1]
    chromeOptions = webdriver.ChromeOptions()
    prefs = {'download.default_directory':'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder',
             'plugins.always_open_pdf_externally':True,
             'download.prompt_for_download':False}
    chromeOptions.add_experimental_option('prefs',prefs)
    driver = webdriver.Chrome(chrome_options = chromeOptions, executable_path=r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\chromedriver.exe')
    os.chdir('O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder')
    for file in os.listdir():
        os.remove(file)
    program_mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='California', usecols=[1,2,3],dtype='str')
    driver.get('https://www.medi-cal.ca.gov/')
    transaction_tab = driver.find_element_by_xpath('//*[@id="nav_list"]/li[2]/a')
    transaction_tab.click()
    wait = WebDriverWait(driver,10)
    user_name = wait.until(EC.element_to_be_clickable((By.ID,'UserID')))
    user_name.send_keys(username)
    pass_word = driver.find_element_by_id('UserPW')
    pass_word.send_keys(password)
    submit_button = driver.find_element_by_id('cmdSubmit')
    submit_button.click()
    user_name2 = login_credentials.iloc[1,0]
    pass_word2 = login_credentials.iloc[1,1]
       
    
    #navigate to the drug rebate invoice page
    drug_rebate = driver.find_element_by_xpath('//*[@id="tabpanel_1_sublist"]/li/a')
    drug_rebate.click()
    ref_dict = {}
    
    #get the three labeler codes.  Will have to update if labeler codes change
    lilly_code = lambda : wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="middle_column"]/div[2]/table/tbody/tr/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[2]/td[2]/a')))
    dista_code = lambda : wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="middle_column"]/div[2]/table/tbody/tr/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[3]/td[2]/a')))
    imclone_code = lambda :wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="middle_column"]/div[2]/table/tbody/tr/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[4]/td[2]/a')))
    
    codes = [lilly_code,dista_code,imclone_code]
    #Below are the prefixes which will be used later for looping through programs/ndc permutations
    lilly_prefix = '00002'
    dist_prefix = '00777'
    imclone_prefix = '66733'
    prefixes = [lilly_prefix,dist_prefix,imclone_prefix]
       
    '''
    This loops through all labeler codes, gets the invoices, unzips them, opens them
    looks through them for NDCs, and creates matched dictionaries of programs and NDCs
    so that you only request the NDCs which associate with programs.
    '''    
    master_dict = {}
    for code,prefix in zip(codes,prefixes):
        driver.get('https://rais.medi-cal.ca.gov/drug/DrugLablr.asp')
        code().click()
        #get the invoices
        xpath = "//a[contains(@href,'DrugRetr')]"
        invoice = wait.until(EC.element_to_be_clickable((By.XPATH,xpath)))
        invoice.click() 
        path = os.getcwd()
        soup = BeautifulSoup(driver.page_source,'html.parser')
        td = soup.find_all('td')
        td = [x.text.replace(' ','').replace('\n','').replace('\t','') for x in td if 'Yr' in x.text]
        d = td[-1]
        d = d.split(';')
        inv_year = d[0].split('=')[1]
        inv_qtr = d[1].split('=')[1]
        file_name = 'ALL_L'+prefix+'_Q'+inv_qtr+'_Y'+inv_year+'.zip'
        all_report = driver.find_element_by_partial_link_text('ALL')
        while file_name not in os.listdir(): 
            all_report.click()
            time.sleep(5)
        '''while any('.crdownload' in x for x in os.listdir())==True:
            time.sleep(3)
        while any('.tmp' in x for x in os.listdir())==True:
            time.sleep(3)'''
    
            
        list_of_files = os.listdir()   
        latest_file = sorted(list_of_files, key=os.path.getctime)[-1]
        latest_file = os.path.abspath(latest_file)
        flag = 1
        
        
        while flag ==1:    
            try:
                with zipfile.ZipFile(latest_file,'r') as zip_ref:
                    zip_ref.extractall()
                os.remove(latest_file)
                flag = 0
            except PermissionError:
                flag = 1
                print('Permission Error')
                time.sleep(1)
            
        list_of_dat_files = list(filter(lambda x: '.dat' in x and 'ALL' not in x,os.listdir()))
        list_of_files = [os.path.splitext(x)[0] for x in list_of_dat_files]
        list_of_text_files = [x+'.txt' for x in list_of_files]
        for x,y in zip(list_of_dat_files,list_of_text_files):
            os.rename(x,y)
        ndcs={}
        for file in list_of_text_files:
            ca_program_code = file.split('_')[0]
            info = open(file,'rt')
            program_ndcs = [line[7:18] for line in info.readlines()]
            program_ndcs = list(set(program_ndcs))
            ndcs.update({ca_program_code:program_ndcs})
            info.close()
        
        names = program_mapper.iloc[:,0].tolist()
        lilly_names = program_mapper.iloc[:,1].tolist()
        programs = program_mapper.iloc[:,2].tolist()        
        mapper = dict(zip(programs,names))
        mapper2 = dict(zip(programs,lilly_names))
        mapper3 = dict(zip(lilly_names,names))
        NDC_List = {}
        for key, value in ndcs.items():
            program = mapper[key]
            xxx = ndcs[key]
            NDC_List.update({program:xxx})
        for value in NDC_List.values():
            for item in value:
                item.lstrip(prefix)
            master_dict.update({code:ndcs})
            ref_dict.update({prefix:NDC_List})
    '''
    This chunk of code below takes the files that have been unzipped and transforms them
    according to the VBA script provided by California.  The script I used was current on
    7/30/2018
    '''        
    for file in os.listdir():
        if file =='debug.log':
            continue
        else:
            pass
        code = file.split('_')[0]
        name = mapper[code]
        if 'MCO' in name:
            xl_id = 'MCOU'
        else:
            xl_id = 'FFSU'
        data = pd.read_csv(file,sep='~',names=list(range(0,17)))
        data[17],data[18] = 0,0
        data[3] = data[3].fillna('0000000000')
        data = data.fillna(0)
        for row in range(len(data)):
            if data.iloc[row,1]==qtr and data.iloc[row,0]==yr:
                data.iloc[row,16]= data.iloc[row,5]
                data.iloc[row,17]= data.iloc[row,8]
                data.iloc[row,18]= 0
            else:
                data.iloc[row,16] = data.iloc[row,6]
                data.iloc[row,17] = round(data.iloc[row,11]+data.iloc[row,12],2)
                data.iloc[row,18] = 1
        data = data.astype(str)
        data[2] = data[2].str.pad(width=11,side='left',fillchar='0')
        data[3] = [x[:10] for x in data[3]]
        data[7] = [x.split('.')[0].rjust(5,'0')+'.'+x.split('.')[1][:6].ljust(6,'0') for x in data[7]]
        for i in range(len(data)):
            if float(data.iloc[i,16])>=0:
                data.iloc[i,16] = data.iloc[i,16].split('.')[0].rjust(11,'0')+'.'+data.iloc[i,16].split('.')[1][:3].ljust(3,'0')
            else:
                data.iloc[i,16] = '-'+data.iloc[i,16].split('.')[0].replace('-','').rjust(11,'0')+'.'+data.iloc[i,16].split('.')[1][:3].ljust(3,'0')
        data[17] = [x.split('.')[0].rjust(9,'0')+'.'+x.split('.')[1][:2].ljust(2,'0') for x in data[17]]        
        data[9] = [x.rjust(8,'0') for x in data[9]]
        data[14] = [x.split('.')[0].rjust(10,'0')+'.'+x.split('.')[1][:2].ljust(2,'0') for x in data[14]]
        data[13] = [x.split('.')[0].rjust(10,'0')+'.'+x.split('.')[1][:2].ljust(2,'0') for x in data[13]]
        data[15] = [x.split('.')[0].rjust(11,'0')+'.'+x.split('.')[1][:2].ljust(2,'0') for x in data[15]]
        data['Z'] = xl_id + 'CA' + data[2]+data[1]+data[0]+data[3]+data[7]+data[16]+data[17]+data[9]+data[14]+data[13]+data[15]+data[18]
        formatted= data[['Z']]
        formatted.to_csv(file,index=False,header=False)    
        
     
        
        
    
    for file in os.listdir():
        if file == 'debug.log':
            continue
        else:
            pass
        program_code = file.split('_')[0]
        program_name = mapper[program_code]
        labeler_code = file.split('_')[1]
        file_name = mapper[program_code]+'_'+'_'.join(file.split('_')[1:])
        path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\'+'Invoices\\California\\'+program_name+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
        if os.path.exists('O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\'+'Invoices\\California\\'+program_name+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\')==False:
            os.makedirs('O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\'+'Invoices\\California\\'+program_name+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\')
            shutil.move(file,path+file_name)
        else:
            shutil.move(file,path+file_name)
    
    '''
    Now moving onto retrieving PDF copies from California
    '''
    
    driver.get('https://rais.medi-cal.ca.gov/drug/DrugLablr.asp')
    for labeler, prefix in zip(codes,prefixes):
        labeler().click()
        copy_of_invoice = lambda: wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="middle_column"]/div[2]/table/tbody/tr/td[2]/table/tbody/tr[1]/td/a/b')))            
        copy_of_invoice().click()            
        for program in master_dict[labeler]:
            drop_down = lambda: wait.until(EC.element_to_be_clickable((By.ID,'Program')))
            drop_down_select = lambda: Select(drop_down())
            drop_down_select().select_by_value(program)
            program_name = mapper2[program]
            quarter = driver.find_element_by_id('Qtr')
            quarter.send_keys(qtr)
            year = driver.find_element_by_id('Year')
            year.send_keys(yr)
            submit_button = driver.find_element_by_xpath('//*[@id="frmDrugRecs"]/table[2]/tbody/tr[6]/td/input[1]')            
            submit_button.click()
            ok = wait.until(EC.element_to_be_clickable((By.ID,'btnOK')))
            ok.click()            
            while (any(map((lambda x: '_wLbl_' in x),os.listdir()))==False or any(map((lambda x: '.tmp' in x),os.listdir()))==True):
                time.sleep(1)
            files = os.listdir()
            latest_file = max(os.listdir(),key=os.path.getctime)
            file_name = mapper[program]+'_'+yr+'_'+qtr+'_'+prefix+'.pdf'
            path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\'+'Invoices\\California\\'+mapper[program]+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
            shutil.move(latest_file,path+file_name)
            driver.get(url)
        driver.get('https://rais.medi-cal.ca.gov/drug/DrugLablr.asp')     
        
    '''
    Now to move the downloaded files into their appropriate lan drive folders
    '''
       
        
    '''Now sending the email notifying the MHS team of which invoices have been downloaded
    '''
    subject = 'California Invoices'
    aaa = list(map((lambda x: x.keys()),ref_dict.values()))
    aaa = list(map((lambda x: list(x)),aaa))
    bbb = dict(zip(ref_dict.keys(),aaa))
    body = 'The following invoices have been downloaded:'
    body2 = ['For %s label code: %s' %(k,v) for k,v in bbb.items()]
    recipient = 'b2b_cma_llymedicaid@lilly.com'
    base = 0x0
    obj = Dispatch('Outlook.Application')
    newMail = obj.CreateItem(base)
    newMail.Subject = subject
    newMail.Body = body+'\n'+str(pprint.pformat(body2))
    newMail.To = recipient
    newMail.display()
    newMail.Send()
            
            
            
    driver.get('https://rais.medi-cal.ca.gov/drug/DrugLablr.asp')
    '''
    Top level of the loop goes through each labeler, second loop goes through each program,
    third loop goes through NDCs for each labeler
    '''    
    user2_todo = {}
    '''
    Go in and find all the things to use in the coming loops.
    Create blank dictionaries which will be upated in the loops
    '''
    lilly_code().click()
    claims_request = lambda: wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="middle_column"]/div[2]/table/tbody/tr/td[2]/table/tbody/tr[3]/td/a[1]/b')))
    claims_request().click()
    program_list = lambda: wait.until(EC.element_to_be_clickable((By.ID,'Program')))
    program_select = lambda: Select(program_list())
    ndc_list = lambda: wait.until(EC.element_to_be_clickable((By.ID,'NDC')))
    list_ndc = lambda: wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="NDCSelect"]')))
    list_ndc_select = lambda: Select(list_ndc())
    add = lambda: wait.until(EC.element_to_be_clickable((By.NAME,'Add')))
    submit = lambda: wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="frmDrugRecs"]/table[2]/tbody/tr[7]/td/input[1]')))
    quarter = lambda: wait.until(EC.element_to_be_clickable((By.ID,'Qtr')))
    year = lambda: wait.until(EC.element_to_be_clickable((By.ID,'Year')))
    programs_with_claims_data = [c.text.replace(' ','') for c in program_select().options]
    issue_programs = {}
    too_much = {}
    #this loop scrubs out all programs from the master dictionary that don't have claims data
    a = ref_dict.copy()
    
    
    #This level of the loop goes through the label code
    #functions and the prefixes
    for label_code, prefix in zip(master_dict.keys(),prefixes):
        print('Getting all info for label code '+prefix )
        #this level of the loop goes through by programs associated to label codes
        for program in master_dict[label_code].keys():  
            print('Getting info for '+program)
            ndc_chunks = [master_dict[label_code][program][i:i+20] for i in range(0,len(master_dict[label_code][program]),20)]
            if len(ndc_chunks)>10:
                alt_list= ndc_chunks[10:]
                ndc_chunks = ndc_chunks[:10]
                too_much.update({program:alt_list})
                user2_todo.update({label_code:too_much})
            else:
                pass
            print(str(len(ndc_chunks)) +' blocks of NDCs to request')
            for chunk in ndc_chunks:
                try:
                    driver.get('https://rais.medi-cal.ca.gov/drug/DrugLablr.asp')
                    label_code().click()
                    claims_request().click()
                    program_select().select_by_value(program)
                except NoSuchElementException:
                    print('This program does not have an invoice for this label code')
                    issue_programs.update({prefix:program})
                    break
                year().send_keys(yr)
                quarter().send_keys(qtr)
                for ndc in chunk:
                    prefix = ndc[:5]
                    ndc = ndc[5:]
                    ndc_list().send_keys(ndc)
                    add().click()
                    ndc_list().clear()
                try:
                    submit().click()
                    try:
                        alert = driver.switch_to.alert
                        alert.accept()
                    except:
                        pass
                    okButton = wait.until(EC.element_to_be_clickable((By.ID,'btnOK')))
                    okButton.click()
                    print('All codes good to go!')
                    print('Returning to transactions')
                    try:
                        return_to_transactions = driver.find_element_by_xpath('//*[@id="frmRet"]/input')
                        return_to_transactions.click()
                        claims_request().click()
                    except NoSuchElementException:
                        print('Too many requests for'+program)
                        print('for program code '+program)
                        _ = {program:current_ndcs}
                        user2_todo.update({label_code:_})
                        break
                except TimeoutException as ex:
                    print('Some codes are invalid!')
                    time.sleep(1)
                    soup = BeautifulSoup(driver.page_source,'html.parser')
                    table = soup.find('table')
                    blue_content = table.find_all('span',attrs={'class':'blueContent'})
                    blue_content = [line.text for line in blue_content]
                    errors = [error for error in blue_content if 'ERROR' in error]
                    non_valid_ndcs = [error[7:13] for error in errors]
                    if len(non_valid_ndcs)==len(chunk):
                        print('All codes have been downloaded or are invalid!')
                        driver.get('https://rais.medi-cal.ca.gov/drug/DrugSelect.asp?sel=pdinv&lbl=00002')
                        continue
                    print(non_valid_ndcs)
                    print('Are invalid NDCs')
                    print('Going back to remove NDCs')
                    driver.back()
                    recall_button = wait.until(EC.element_to_be_clickable((By.NAME,'Recall')))
                    recall_button.click()
                    print('Removing NDCs')
                    print(non_valid_ndcs)
                    current_ndcs = [curr.text[5:] for curr in list_ndc_select().options]
    
                    while any(x in current_ndcs for x in non_valid_ndcs)==True:
                        for x in non_valid_ndcs:
                            while x in [z.text[5:] for z in list_ndc_select().options]:
                                print(x+' Is invalid, removing')
                                list_ndc_select().deselect_all()
                                list_ndc_select().select_by_value(prefix+x)
                                remove_button = lambda: wait.until(EC.element_to_be_clickable((By.NAME,'Remove')))
                                remove_button().click()
     #------------------------>TEST!    #time.sleep(1)<------------------------------------------------STILL NOT TESTED BE SURE TO TEST!
                                print('Removed '+x)
                                current_ndcs = [curr.text[5:] for curr in list_ndc_select().options]
    
                        submit().click()
                        ok_button = wait.until(EC.element_to_be_clickable((By.ID,'btnOK')))    
                        ok_button.click()    
                        try:
                            return_to_transactions = driver.find_element_by_xpath('//*[@id="frmRet"]/input')
                            return_to_transactions.click()
                        except NoSuchElementException:
                            print('Too many requests!')
                            print('The following NDCs could not be downloaded')
                            for item in current_ndcs:
                                print(item)
                            print('for program code '+program)
                            _ = {program:current_ndcs}
                            user2_todo.update({label_code:_})
     
    
    
    exit_link = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="nav_list"]/li[2]/ul/li[2]/a')))
    exit_link.click()
    transactions_link = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="nav_list"]/li[2]/a')))
    transactions_link.click()     
    user_name = wait.until(EC.element_to_be_clickable((By.ID,'UserID')))
    pass_word = driver.find_element_by_id('UserPW')                           
    user_name.send_keys(user_name2)
    pass_word.send_keys(pass_word2)                        
    submit_button = driver.find_element_by_id('cmdSubmit')
    submit_button.click()            
    driver.get('https://rais.medi-cal.ca.gov/drug/DrugLablr.asp')  
    user2_issue_programs = {}
    for label_code in user2_todo.keys():
        print('Getting all info for leftover NDCS')
        for program in user2_todo[label_code].keys():  
            print('Getting info for '+program)
            print('There are ' + str(len(user2_todo[label_code][program]))+' chunks to request')
            for chunk in user2_todo[label_code][program]:
                print('Getting next chunk')
                try:
                    driver.get('https://rais.medi-cal.ca.gov/drug/DrugLablr.asp')
                    label_code().click()
                    claims_request().click()
                    program_select().select_by_value(program)
                except NoSuchElementException:
                    print('This program does not have an invoice for this label code')
                    user2_issue_programs.update({label_code:program})
                    break
                year().send_keys(yr)
                quarter().send_keys(qtr)
                for ndc in chunk:
                    prefix = ndc[:5]
                    ndc = ndc[5:]
                    ndc_list().send_keys(ndc)
                    add().click()
                    ndc_list().clear()
                try:
                    submit().click()
                    try:
                        alert = driver.switch_to.alert
                        alert.accept()
                    except:
                        pass
                    okButton = wait.until(EC.element_to_be_clickable((By.ID,'btnOK')))
                    okButton.click()
                    print('All codes good to go!')
                    print('Returning to transactions')
                    try:
                        return_to_transactions = driver.find_element_by_xpath('//*[@id="frmRet"]/input')
                        return_to_transactions.click()
                        claims_request().click()
                    except NoSuchElementException:
                        print('Too many requests for'+program)
                        break
                except TimeoutException as ex:
                    print('Some codes are invalid!')
                    time.sleep(1)
                    soup = BeautifulSoup(driver.page_source,'html.parser')
                    table = soup.find('table')
                    blue_content = table.find_all('span',attrs={'class':'blueContent'})
                    blue_content = [line.text for line in blue_content]
                    errors = [error for error in blue_content if 'ERROR' in error]
                    non_valid_ndcs = [error[7:13] for error in errors]
                    if len(non_valid_ndcs)==len(chunk):
                        print('All codes have been downloaded or are invalid!')
                        driver.get('https://rais.medi-cal.ca.gov/drug/DrugSelect.asp?sel=pdinv&lbl=00002')
                        break
                    print(non_valid_ndcs)
                    print('Are invalid NDCs')
                    print('Going back to remove NDCs')
                    driver.back()
                    recall_button = wait.until(EC.element_to_be_clickable((By.NAME,'Recall')))
                    recall_button.click()
                    print('Removing NDCs')
                    print(non_valid_ndcs)
                    current_ndcs = [curr.text[5:] for curr in list_ndc_select().options]
    
                    while any(x in current_ndcs for x in non_valid_ndcs)==True:
                        for x in non_valid_ndcs:
                            while x in [z.text[5:] for z in list_ndc_select().options]:
                                print(x+' Is invalid, removing')
                                list_ndc_select().deselect_all()
                                list_ndc_select().select_by_value(prefix+x)
                                remove_button = lambda: wait.until(EC.element_to_be_clickable((By.NAME,'Remove')))
                                remove_button().click()
     #------------------------>TEST!    #time.sleep(1)<------------------------------------------------STILL NOT TESTED BE SURE TO TEST!
                                print('Removed '+x)
                                current_ndcs = [curr.text[5:] for curr in list_ndc_select().options]
    
                        submit().click()
                        ok_button = wait.until(EC.element_to_be_clickable((By.ID,'btnOK')))    
                        ok_button.click()    
                        return_to_transactions = driver.find_element_by_xpath('//*[@id="frmRet"]/input')
                        return_to_transactions.click()
def main():
    cali_one()

if __name__=='__main__':
    main()                                                   
        

    

    
    
    
    
    