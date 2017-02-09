#!/usr/bin/env python3
# -*- coding:utf-8 -*-

from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import colors
from openpyxl.styles import Fill,fills
from openpyxl.styles import Border
from openpyxl.styles import Side
from openpyxl.formatting.rule import ColorScaleRule
import time
import datetime

#打开工作簿 open the workbook
wb = load_workbook('IT Daily Report.xlsx')

#定义边框
BD = Border(left=Side(style='medium',color='FF000000'),right=Side(style='medium',color='FF000000'),bottom=Side(style='medium',color='FF000000'))

#读取全部表名 read all sheet names
sheets = wb.get_sheet_names()
print (sheets)

#获取当前日期,并循环将连续五天的日期写入days这个list
date_times = 1
days = []
today = datetime.datetime.today()
while date_times <= 5:
    processing_day = str(today.strftime('%mM%dD'))
    processing2_day = processing_day.replace("M","月")
    usable_day = processing2_day.replace("D","日")
    days.append(usable_day)
    deltadays = datetime.timedelta(days=1)
    today = today + deltadays
    date_times = date_times + 1
else:
	print (days)


#循环
#循环更改表名
Index = 0
while Index < 5:
   st = sheets[Index]
   actsheet = wb.get_sheet_by_name(st)
   actsheet['A20'].border = BD
   actsheet.title = days[Index]
   print (actsheet.title)
   Index = Index + 1
else:
	#另存为文件，并在文件名中增加保存日期
	date2nd = datetime.datetime.today()
	savedate = str(date2nd.strftime('20%y%m%d'))
	wb.save('IT Daily Report %s.xlsx' %savedate)