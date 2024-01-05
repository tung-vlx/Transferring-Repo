import os
path = os.getcwd()

import time
startTime = time.time()
from PIL import Image

import selenium
from selenium import webdriver
chromeOpts = webdriver.EdgeOptions()
chromeOpts.add_argument("window-size=960,1080")
chrome = webdriver.Edge(options = chromeOpts)
from selenium.webdriver.common.by import By

def crop_image(image_name, elementStart_x, elementStart_y, elementEnd_x, elementEnd_y):
    img = Image.open(image_name)
    img = img.crop((elementStart_x, elementStart_y, elementEnd_x, elementEnd_y))
    img.save(image_name)

def insert_image_to_Excel(ws,image_name, cell_index):
    img = openpyxl.drawing.image.Image(image_name)
    img.anchor = cell_index
    ws.add_image(img)

def insert_image(ws, stock_code, row_index):
    image_name = 'Excel/resource/' + stock_code + '-1.png'
    cell_index = 'A' + str(row_index)
    insert_image_to_Excel(ws, image_name, cell_index)

    image_name = 'Excel/resource/' + stock_code + '-2.png'
    cell_index = 'O' + str(row_index)
    insert_image_to_Excel(ws, image_name, cell_index)

def stock_code_screenshot(stock_code, ws, row_index):
    chrome.get('https://stockbiz.vn/ma-chung-khoan/' + stock_code)
    chrome.execute_script('window.scrollTo(0,80)')
    chrome.execute_script("document.body.style.zoom='85%'")
    time.sleep(2)
    try:
        element_more = chrome.find_elements(By.CLASS_NAME, 'py-4')[1]
        element_more = element_more.find_element(By.TAG_NAME, 'a')
        chrome.execute_script ("arguments[0].click();", element_more)
    except:
        print('', end='\r')
    image_name = 'Excel/resource/' + stock_code + '-1.png'
    chrome.save_screenshot(image_name)

    element_shareholder = chrome.find_element(By.XPATH, '//*[@id="__next"]/div[3]/div/main/div/div[2]/div[1]/ul/li[2]/div')
    chrome.execute_script ("arguments[0].click();", element_shareholder)
    element_shareholder = chrome.find_element(By.CLASS_NAME, 'my-6')
    chrome.execute_script('window.scrollTo(0,2000)')
    time.sleep(0.5)
    chrome.execute_script('window.scrollTo(0,-2000)')
    time.sleep(0.5)
    chrome.execute_script('window.scrollTo(0,' + str((element_shareholder.location['y']-element_shareholder.size['height'])*0.85) + ')')
    time.sleep(0.5)
    image_name = 'Excel/resource/' + stock_code + '-2.png'
    chrome.save_screenshot(image_name)
    insert_image(ws, stock_code, row_index)


import os
import shutil
try:
    os.mkdir(path + "/Excel/resource")
    print("--------------------")
    print("Create temp folder")
except:
    shutil.rmtree(path + "/Excel/resource")
    os.mkdir(path + "/Excel/resource")
    print("--------------------")
    print("Delete existing temp folder")
    print("Create new temp folder")

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment, Color
wb = openpyxl.load_workbook(path + '/Excel/input.xlsx')
ws_main = wb.worksheets[0]

start_row = 0
print("--------------------")
print("Finding Local Search Table")
for row in range(1, ws_main.max_row + 1):
    cell_value = ws_main.cell(row=row, column=2).value
    if cell_value and "vietnamese" in str(cell_value).lower():
        start_row = row
        break
start_row += 1
print("--------------------")
print("Scanning Local Search Table")
for row in range(start_row, ws_main.max_row + 1):
    if ws_main.cell(row=row, column=4).value == None:
        print("--------------------")
        print("Detecting ", str(ws_main.cell(row=row, column=1).value))
        sheet_name = 'Code ' + str(ws_main.cell(row=row, column=1).value[0:4])
        print("--------------------")
        print("Create sheet ", sheet_name)
        wb.create_sheet(sheet_name)
        first_line_index = row + 1
        ws_ICB = wb[sheet_name]
        continue
    if row - first_line_index == 0:
        ws_ICB_row_index = 1
    else:
        ws_ICB_row_index = (row - first_line_index) * 70
    print("++++++++++")
    print("#Processing ", ws_main.cell(row=row, column=4).value)
    ws_ICB.cell(row=ws_ICB_row_index, column=1).value = ws_main.cell(row=row, column=4).value
    ws_ICB.cell(row=ws_ICB_row_index, column=1).font = Font(name='Georgia', b=True, sz=11 , color='000000')
    ws_ICB.cell(row=ws_ICB_row_index, column=1).fill = PatternFill('solid', fgColor='FFFF00')
    print("#Screenshot ", ws_main.cell(row=row, column=4).value)
    print("++++++++++")
    stock_code_screenshot(ws_main.cell(row=row, column=4).value, ws_ICB, ws_ICB_row_index+1)
print("--------------------")
print("Closing Browser")
chrome.quit()
print("--------------------")
print("Saving Excel File")
try:
    wb.save('Local Search & Screenshot.xlsx')
except :
    while_bln = True
    i=1
    while while_bln == True:
        try:
            wb.save('Local Search & Screenshot-' + str(i) + '.xlsx')
            break
        except:
            i += 1
print("--------------------")    
print("End")
print("--------------------")    
print("Factos:")
processing_time = int(time.time() - startTime)
if processing_time < 60:
    print("# Running Time: " + str(processing_time) + " s")
elif processing_time < 3599:
    print("# Running Time: " + str(processing_time//60) + "m" + str(processing_time%60) + "s")
else:
    print("# Running Time: " + str(processing_time//3600) + "h" + str(processing_time//3600//60) + "m" + str(processing_time//3600%60) + "s")
print("# Made by yoyoitsme")
print("--------------------")    
print("Press any key to continue...")
input()
