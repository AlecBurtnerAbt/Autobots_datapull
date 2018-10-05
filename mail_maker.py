# -*- coding: utf-8 -*-
"""
Created on Mon Sep 17 08:42:03 2018

@author: C252059
"""


from win32com.client import Dispatch


def send_message(subject,body,to):
    subject = subject
    body = body
    recipient = to
    base = 0x0
    obj = Dispatch('Outlook.Application')
    newMail = obj.CreateItem(base)
    newMail.Subject = subject
    newMail.Body = body+'\nBeep Boop, I am a robot.  For issues please contact the Data Pull Support Team'
    newMail.To = recipient
    newMail.display()
    newMail.Send()
    
def main():
    send_message(subject,body,to)
    
if __name__=='__main__':
    main()