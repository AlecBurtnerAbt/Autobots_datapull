# -*- coding: utf-8 -*-
"""
Created on Fri Sep 21 11:12:21 2018

@author: C252059
"""

import socket

HOST = '127.0.0.1'

PORT = 65432

with socket.socket(socket.AF_INET,socket.SOCK_STREAM) as s:
    s.connect((HOST,PORT))
    s.sendall(b'Hello World!')
    data = s.recv(1024)
    
print('Receive'+repr(data))