#! /usr/bin/env python
# coding: utf-8

import xlwt
from openpyxl import load_workbook
print 'početak'
data = load_workbook('/home/matija/PycharmProjects/excell_script/Analiza realizacije razrada 01.01.2009 - 31.12.2017..xlsx')
print 'čitam'
sheet = data.get_sheet_names()[0]
worksheet = data.get_sheet_by_name(sheet)
