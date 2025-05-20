def income_statement(ws, columnIndex, rowIndex, response, year):
    # Net Sales
    try:
        for item in response[2]['values']:
            if item['period'] == year:
                ws.cell(row=rowIndex, column = columnIndex+1).value = item['value']         
    except:
        ws.cell(row=rowIndex, column = columnIndex+1).value = 'error'
    # COGS
    try:
        for item in response[3]['values']:
            if item['period'] == year:
                ws.cell(row=rowIndex, column = columnIndex+2).value = item['value']
    except:
        ws.cell(row=rowIndex, column = columnIndex+2).value = 'error' 
    # Gross Profit
    try:
        ws.cell(row=rowIndex, column=columnIndex+3).value = '='+get_column_letter(columnIndex+1)+str(rowIndex)+'-'+get_column_letter(columnIndex+2)+str(rowIndex)       
    except:
        ws.cell(row=rowIndex, column=columnIndex+3).value = 'error'                
    # Selling Expenses
    try:
        for item in response[9]['values']:
            if item['period'] == year:
                ws.cell(row=rowIndex, column=columnIndex+4).value = item['value']
    except:
        ws.cell(row=rowIndex, column=columnIndex+4).value = 'error'   
    # General and Administration Expenses
    try:
        for item in response[10]['values']:
            if item['period'] == year:
                ws.cell(row=rowIndex, column=columnIndex+5).value = item['value']
    except:
        ws.cell(row=rowIndex, column=columnIndex+5).value = 'error' 
    # EBIT
    try:
        ws.cell(row=rowIndex, column=columnIndex+6).value = '='+get_column_letter(columnIndex+3)+str(rowIndex)+'-'+get_column_letter(columnIndex+4)+str(rowIndex)+'-'+get_column_letter(columnIndex+5)+str(rowIndex) 
    except:
        ws.cell(row=rowIndex, column=columnIndex+6).value = 'error'
    # Total Costs
    try:
        ws.cell(row=rowIndex, column=columnIndex+7).value = '='+get_column_letter(columnIndex+1)+str(rowIndex)+'-'+get_column_letter(columnIndex+6)+str(rowIndex) 
    except:
        ws.cell(row=rowIndex, column=columnIndex+7).value = 'error'
    # MOTC
    try:
        ws.cell(row=rowIndex, column=columnIndex+8).value = '='+get_column_letter(columnIndex+6)+str(rowIndex)+'/'+get_column_letter(columnIndex+7)+str(rowIndex)
        ws.cell(row=rowIndex, column=columnIndex+8).style = style_percentage
    except:
        ws.cell(row=rowIndex, column=columnIndex+8).value = 'error'
    # ROS
    try:
        ws.cell(row=rowIndex, column=columnIndex+9).value = '='+get_column_letter(columnIndex+6)+str(rowIndex)+'/'+get_column_letter(columnIndex+1)+str(rowIndex)
        ws.cell(row=rowIndex, column=columnIndex+9).style = style_percentage
    except:
        ws.cell(row=rowIndex, column=columnIndex+9).value = 'error'

def balance_sheet(ws, columnIndex, rowIndex, response, year):
    # Inventories
    try:
        for item in response[17]['values']:
            if item['period'] == year:
                ws.cell(row=rowIndex, column = columnIndex+1).value = item['value']         
    except:
        ws.cell(row=rowIndex, column = columnIndex+1).value = 'error'
    # Tangible Fixed Assets
    try:
        for item in response[35]['values']:
            if item['period'] == year:
                ws.cell(row=rowIndex, column = columnIndex+2).value = item['value']
    except:
        ws.cell(row=rowIndex, column = columnIndex+2).value = 'error'                
    # Intangible Fixed Assets
    try:
        for item in response[41]['values']:
            if item['period'] == year:
                ws.cell(row=rowIndex, column=columnIndex+3).value = item['value']
    except:
        ws.cell(row=rowIndex, column=columnIndex+3).value = 'error'   
    # Total Assets
    try:
        for item in response[61]['values']:
            if item['period'] == year:
                ws.cell(row=rowIndex, column=columnIndex+4).value = item['value']
    except:
        ws.cell(row=rowIndex, column=columnIndex+4).value = 'error'     


def stock_by_industry(ws, response, count, start_year, period):
    rowIndex = 2
    title_general = ['Listed name', 'ICB code', 'Vietnamese Name', 'English Name', 'Listed in', 'Business Description','Biggest Shareholder', 'Website']
    title_fs = ['Net Sales', 'Cost of Goods Sold', 'Gross Profit', 'Selling Expenses', 'G&A Expenses', 'EBIT', 'Total Costs', 'MOTC', 'ROS', 'Inventories', 'Tangible Fixed Assets', 'Intangible Fixed Assets', 'Total Assets']
    for columnIndex, title in enumerate(title_general, start=1):
        ws.cell(row=2, column=columnIndex).value = title
    for i in range(int(period)-1):
        year = str(int(start_year)-i)
        ws.cell(row=1, column=len(title_general)+len(title_fs)*i+1).value = year
        for columnIndex, title in enumerate(title_fs, start=1):
            ws.cell(row=2, column=len(title_general)+len(title_fs)*i+columnIndex).value = title
    rowIndex +=1
    for instrument in response:
        count[instrument['type']]['i'] += 1
        if instrument['type'] == 'stock': 
            print(str(rowIndex-2), end = '\r')
            link_stockCode = 'https://api.fireant.vn/symbols/' + instrument['symbol'] + '/profile'
            res_stockCode = requests.get(link_stockCode, headers = headers)
            response_stockCode = json.loads(res_stockCode.text)
            if response_stockCode['isListed'] == False:
                continue

            ws.cell(row=rowIndex, column = 1).value = instrument['symbol']
            try:
                ws.cell(row=rowIndex, column = 2).value = response_stockCode['icbCode']
            except:
                ws.cell(row=rowIndex, column = 2).value = '-'
            try:
                ws.cell(row=rowIndex, column = 3).value = response_stockCode['companyName']
            except:
                ws.cell(row=rowIndex, column = 3).value = '-'
            try:
                ws.cell(row=rowIndex, column = 4).value = response_stockCode['internationalName']
            except:
                ws.cell(row=rowIndex, column = 4).value = '-'
            try:
                ws.cell(row=rowIndex, column = 5).value = response_stockCode['exchange']
            except:
                ws.cell(row=rowIndex, column = 5).value = '-'
            ws.cell(row=rowIndex, column = 6).value = response_stockCode['overview']
            
            ws.cell(row=rowIndex, column = 8).value = response_stockCode['webAddress']

            # link_shareholder = 'https://api.fireant.vn/symbols/' + instrument['symbol'] + '/holders'
            # res_shareholder = requests.get(link_shareholder, headers = headers)
            # response_shareholder = json.loads(res_shareholder.text)
            # try:
            #     ownership = '{:.2%}'.format(response_shareholder[0]['ownership'])
            # except:
            #     ownership = 'error'
            # try:
            #     ws_summary.cell(row=rowIndex, column = 7).value = response_shareholder[0]['name'] + ' - ' + str(ownership)
            # except:
            #     ws_summary.cell(row=rowIndex, column = 7).value = 'error'

            link_ic = "https://restv2.fireant.vn/symbols/"+ instrument['symbol'] +"/full-financial-reports?type=2&year="+ str(current_year) +"&quarter=0&limit=7"
            res_ic = requests.get(link_ic, headers = headers)
            response_ic = json.loads(res_ic.text)

            link_bs = "https://restv2.fireant.vn/symbols/"+ instrument['symbol'] +"/full-financial-reports?type=1&year="+ str(current_year) +"&quarter=0&limit=7"
            res_bs = requests.get(link_bs, headers = headers)
            response_bs = json.loads(res_bs.text)


            for i in range(int(period)-1):
                year = str(int(start_year)-i)
                columnIndex = len(title_general) + len(title_fs)*i
                income_statement(ws, columnIndex, rowIndex, response_ic, year)
                columnIndex += 9
                balance_sheet(ws, columnIndex, rowIndex, response_bs, year)
            
            rowIndex +=1

def icb_list(ws, response):
    rowIndex = 1
    for item in response:
        ws.cell(row=rowIndex, column = 1).value = str(item['industryCode'])
        ws.cell(row=rowIndex, column = 2).value = item['name']
        rowIndex += 1

print('# Start')
from datetime import datetime
current_datetime = datetime.now()
current_date = current_datetime.day
current_month = current_datetime.month
current_year = current_datetime.year

import time
startTime = time.time()

import os
current_directory = os.getcwd()
path = r'D:\Python'

# User Input
start_year = str(input('Latest study year: '))
period = str(input('Number of year study: '))

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import NamedStyle
style_percentage = NamedStyle(name='percentage', number_format='0.00%')
wb = openpyxl.load_workbook(path + '/Excel/StockBiz.xlsx')
ws_summary = wb['Summary']
print('# Open Excel')

import requests
import json
headers = {"Content-Type": "Application/json", "Authorization": "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6IkdYdExONzViZlZQakdvNERWdjV4QkRITHpnSSIsImtpZCI6IkdYdExONzViZlZQakdvNERWdjV4QkRITHpnSSJ9.eyJpc3MiOiJodHRwczovL2FjY291bnRzLmZpcmVhbnQudm4iLCJhdWQiOiJodHRwczovL2FjY291bnRzLmZpcmVhbnQudm4vcmVzb3VyY2VzIiwiZXhwIjoyMDAzNTYwOTU3LCJuYmYiOjE3MDM1NjA5NTcsImNsaWVudF9pZCI6InN0b2NrYml6LndlYiIsInNjb3BlIjpbIm9wZW5pZCIsInByb2ZpbGUiLCJyb2xlcyIsImVtYWlsIiwiYWNjb3VudHMtcmVhZCIsImFjY291bnRzLXdyaXRlIiwib3JkZXJzLXJlYWQiLCJvcmRlcnMtd3JpdGUiLCJjb21wYW5pZXMtcmVhZCIsImluZGl2aWR1YWxzLXJlYWQiLCJmaW5hbmNlLXJlYWQiLCJwb3N0cy13cml0ZSIsInBvc3RzLXJlYWQiLCJzeW1ib2xzLXJlYWQiLCJ1c2VyLWRhdGEtcmVhZCIsInVzZXItZGF0YS13cml0ZSIsInVzZXJzLXJlYWQiLCJzZWFyY2giLCJhY2FkZW15LXJlYWQiLCJhY2FkZW15LXdyaXRlIiwiYmxvZy1yZWFkIiwiaW52ZXN0b3BlZGlhLXJlYWQiXSwic3ViIjoiMjczZjM2YjQtYmYxZS00ODVkLWFkY2ItMDU5OGExMTMxODI5IiwiYXV0aF90aW1lIjoxNzAzNTYwOTU3LCJpZHAiOiJHb29nbGUiLCJuYW1lIjoidm8ubGUueHVhbi50dW5nQHB3Yy5jb20iLCJzZWN1cml0eV9zdGFtcCI6ImFhMjFmNzc3LWY5NmYtNDEwMS05Mzk5LTgxZmUxYWNlNDZiMiIsImp0aSI6IjI1NjE2Y2ZkODRlNGFlYjU5ZGM2ZGFiY2Q0OWYyYzNlIiwiYW1yIjpbImV4dGVybmFsIl19.cnTscfQVW86c7cc27xOX51ukl8Ih-yJO_iqE5E1U35WD09IXID6hmo6liYSnY0Qlb-mQkYI_wtBHfa1GBo2K_5O84bldqVLaowx_FCZNSevAVF8jeJsMukND2SrfntANX2goC0_bVaki4knR7HQ77BNiq0nFvCndkEMKiLOFV2CzFy4dnqjdM4Ub2i1QN1G9yOxMS1q8zLmvwQECs9OWBFkfqxmCPgg4sjoigTtehhwp66i02NvzAvZjELdlLVExXUSvhhPWeaavp2yVC4mLivfdF7Me9GhjPBmjVUP-M-fk8AnHZ7BQaOY55Fdm2XUChf_g_gD0rJCg-uEZTMd8fQ"}

res = requests.get('https://api.fireant.vn/instruments', headers = headers)
response_instrument = json.loads(res.text)
print('# Collect stock codes')
count = {
    'stock': {
        'i': 0,
        'name': 'stock (cổ phiếu)'
    },
    'index': {
        'i': 0,
        'name': 'index (chỉ số)'
    },
    'futures': {
        'i': 0,
        'name': 'futures (hợp đồng tương lai)'
    },
    'warrant': {
        'i': 0,
        'name': 'warrant (chứng quyền)'
    },
    'fund': {
        'i': 0,
        'name': 'fund (quỹ)'
    },
    'bond': {
        'i': 0,
        'name': 'bond (trái phiếu)'
    },
    'commodity': {
        'i': 0,
        'name': 'commodity (hàng hóa)'
    }
}
stock_by_industry(ws_summary, response_instrument, count, start_year, period)

wb.create_sheet("ICB Summary")
ws_ICB = wb['ICB Summary']
res = requests.get('https://api.fireant.vn/industries', headers=headers)
response = json.loads(res.text)
icb_list(ws_ICB, response)
        
print('\nCompleted')
file_name = "StockBiz - " + str(current_date) + "." + str(current_month) + "." + str(current_year)
wb.save(file_name + ".xlsx")

processing_time = int(time.time() - startTime)
print(f"Total companies under Instrument link: {len(response_instrument)}")
for item in count:
    print(f"Number of companies classified under {count[item]['name']}: {str(count[item]['i'])}")

if processing_time < 60:
    print("# Running Time: " + str(processing_time) + " s")
elif processing_time < 3599:
    print("# Running Time: " + str(processing_time//60) + "m" + str(processing_time%60) + "s")
else:
    print("# Running Time: " + str(processing_time//3600) + "h" + str(processing_time//3600//60) + "m" + str(processing_time//3600%60) + "s")

print("Press any key to continue...")
input()