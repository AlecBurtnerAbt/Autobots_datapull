# -*- coding: utf-8 -*-
"""
Created on Tue Sep 11 10:28:29 2018

@author: C252059
"""

import mechanicalsoup as mech
from requests.auth import HTTPProxyAuth
from pypac import PACSession, get_pac
import requests


session = requests.Session()
session.auth = HTTPProxyAuth('c252059','AnugsHound123!@3')
session.proxies = proxies
session.get('http://google.com')

session.close()





pac = get_pac('us_proxy_indy.xh1.lilly.com:9000')

session = PACSession(auth=HTTPProxyAuth('c252059','AngusHound123!@#'))
session.get('http://google.com')
response = session.get('https://rsp.pagov.changehealthcare.com/RebateServicesPortal/reports/index',data=payload)
response2 = session.post('https://rsp.pagov.changehealthcare.com/RebateServicesPortal/reports/index',data=payload)


browser = mech.StatefulBrowser()
browser.get('https://rsp.pagov.changehealthcare.com/RebateServicesPortal/dashboard/index',
            data=login_data,proxies=proxies,auth=HTTPProxyAuth('c252059','AngusHound123!@#'))
login_data = {
        'j_username':username,
        'j_password':password
        }

accept_data = {'terms':'Accept'}
payload = {}
        'stateReportID':'EXT Claim Level Detail Report',
        'docType':'OBRA',
        'rpuStart':'20182',
        'ndc':'00002143380'
        }
proxies = {
        'http':'us_proxy_indy.xh1.lilly.com:9000',
        'https':'us_proxy_indy.xh1.lilly.com:9000'
        }
os.chdir('C:/Users/c252059/Desktop/')
with open('page1.txt','a') as ax:
    ax.write(str(response2.text))
    
    
    
sesh = PACSession(proxies=proxies,auth=HTTPProxyAuth('c252059','AngusHound123!@#'))
post =  sesh.post('https://rsp.pagov.changehealthcare.com/RebateServicesPortal/application/termsLoggingIn',data=login_data)   
auth = sesh.post('https://rsp.pagov.changehealthcare.com/RebateServicesPortal/dashboard/index',data=accept_data)
git = sish.post('https://rsp.pagov.changehealthcare.com/RebateServicesPortal/reports/index',data=payload)                           