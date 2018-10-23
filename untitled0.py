# -*- coding: utf-8 -*-
"""
Created on Thu Oct 18 13:34:11 2018

@author: C252059
"""

import csv

import sqlite3

with sqlite3.connect("new.db") as connection:
    c = connection.cursor()

    # open the csv file and assign it to a variable
    employees = csv.reader(open(r"C:\Users\c252059\Documents\RealPython\real-python-test\sql\employees.csv", "rU"))

    # create a new table called employees
    c.execute("CREATE TABLE employees(firstname, lastname)")

    # insert data into table
    c.executemany("INSERT INTO employees(firstname, lastname) values (?, ?)", employees)