from openpyxl import Workbook
from openpyxl import load_workbook
import datetime
from openpyxl.styles import NamedStyle
import requests
import alpha_vantage
from xml.dom import minidom
from tkinter import messagebox

mydoc = minidom.parse("config.xml")
token = mydoc.getElementsByTagName('token')
function = mydoc.getElementsByTagName('function')
outputsize = mydoc.getElementsByTagName('outputsize')
interval = mydoc.getElementsByTagName('interval')


if function[0].firstChild.data == 'TIME_SERIES_DAILY':
    datastring = 'Time Series (Daily)'
    intervalstr = ""
else:
    datastring = 'Time Series (' + str(interval[0].firstChild.data) + ')'
    intervalstr = interval[0].firstChild.data, 

symbols = []


file = open("stocklist.txt", "r") 
for line in file: 
    symbols.append(line.rstrip())
file.close()


date_style = NamedStyle(name='datetime', number_format='MM/DD/YYYY HH:MM:SS AM/PM')
date_style_d = NamedStyle(name='datetime_d', number_format='MM/DD/YYYY')
API_URL = "https://www.alphavantage.co/query" 
#symbols = ["BAC","MS","JPM"]
wb = Workbook()

#use alphavantage to download data and save in excel file
s = 1
for symbol in symbols:
        data = { 
        "function": function[0].firstChild.data, 
        "symbol": symbol,
        "outputsize": outputsize[0].firstChild.data,
        "interval" : intervalstr,       
        "datatype": "json", 
        "apikey": token[0].firstChild.data } 
        response = requests.get(API_URL, data) 
        data = response.json()
        a = (data[datastring])
        keys = (a.keys())

        if s == 1:
              ws0 = wb.active
              ws0.title = symbol
              ws0 = wb[symbol]
        else:
              ws0 = wb.create_sheet(symbol)
              ws0 = wb[symbol]
        r = 1
        for key in keys:
              for y in range(1,8):
                  if y == 1:
                      ws0.cell(row = r, column = y).value = symbol
                  elif y == 2:
                      if function[0].firstChild.data == 'TIME_SERIES_INTRADAY':
                           date_time_obj = datetime.datetime.strptime(key, '%Y-%m-%d %H:%M:%S')
                           ws0.cell(row = r, column = y).style = date_style
                      else:
                           date_time_obj = datetime.datetime.strptime(key, '%Y-%m-%d')
                           ws0.cell(row = r, column = y).style = date_style_d
                      
                      
                      ws0.cell(row = r, column = y).value = date_time_obj
                  elif y == 3:
                      ws0.cell(row = r, column = y).value = float((a[key]['5. volume']))
                  elif y == 4:
                      ws0.cell(row = r, column = y).value = float((a[key]['1. open']))
                  elif y == 5:
                      ws0.cell(row = r, column = y).value = float((a[key]['2. high']))
                  elif y == 6:
                      ws0.cell(row = r, column = y).value = float((a[key]['3. low']))
                  elif y == 7:
                      ws0.cell(row = r, column = y).value = float((a[key]['4. close']))
              r = r + 1                   
        s = s + 1

if function[0].firstChild.data == 'TIME_SERIES_INTRADAY':
    intvstr = str(intervalstr).replace("(","").replace(")","").replace(",","").replace("'","") + "_"
else:
    intvstr = ""

now = datetime.datetime.now()
wb.save(filename = function[0].firstChild.data + "_" + intvstr + now.strftime("%m/%d/%Y, %H:%M:%S").replace(":", "_").replace("/", "_").replace(",", "_") + ".xlsx")
        