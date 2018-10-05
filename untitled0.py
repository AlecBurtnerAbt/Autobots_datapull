# -*- coding: utf-8 -*-
"""
Created on Wed Oct  3 09:23:23 2018

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


class Driver():
        
    def __init__(self):
        self._name = 'Chrome Driver'
        self.prefs = {'download.default_directory':'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder',
             'plugins.always_open_pdf_externally':True,
             'download.prompt_for_download':False}
    def generate_browser(self):
        chromeOptions = webdriver.ChromeOptions()
        chromeOptions.add_experimental_option("prefs",self.prefs)
        browser = webdriver.Chrome(chrome_options = self.chromeOptions)
        return browser

   
