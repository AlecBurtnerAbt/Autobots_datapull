# -*- coding: utf-8 -*-
"""
Created on Mon Jul 16 14:20:05 2018

@author: C252059
"""

from selenium import webdriver

def IdahoMedicaid(yr,qtr):
    yq = str(yr)+str(qtr)
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    import time
    import pyautogui as pgi
    from bs4 import BeautifulSoup 
    driver = webdriver.Chrome()
    wait = WebDriverWait(driver,10)
    driver.get('https://einvoicing.magellanmedicaid.com/rebate')
    user_name = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="input_1"]')))
    user_name.send_keys('llymedicaid')
    pass_word = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="input_2"]')))
    pass_word.send_keys('ELlymdcd!')
    login_button = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="auth_form"]/fieldset/ol[2]/li/input')))
    login_button.click()

    '''
    This part of the code grabs the invoices
    '''
    other_agencies = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="mainForm:manufacturerTable:1:sortByNameLink"]')))
    other_agencies.click()
    invoices = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="mainForm:invoices"]')))
    invoices.click()
    select_lilly = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="selectedManufacturer"]')))
    select_lilly.click()
    continue_1 = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="mainForm:btnContinue"]')))
    continue_1.click()
    year_quarter = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="mainForm:selYearDate"]')))
    select = webdriver.support.ui.Select(year_quarter)
    select.select_by_visible_text(yq)
    continue_2 = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="mainForm:btnContinue"]')))
    continue_2.click()
    
    #Have to get all of the states available and get their invoices
    wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="mainForm:btnContinue"]')))
20181_Nebraska_MCO Supplemental Program_00002
        

                                                        







IdahoMedicaid(2017,2)


