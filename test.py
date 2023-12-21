from PIL import Image
import pytesseract
import cv2
from imutils import contours
import numpy as np

import openpyxl
from openpyxl_image_loader import SheetImageLoader
import os
import shutil
import subprocess

path = r'C:\Users\vtung010\Downloads\Stockbiz Scraping\Python\Tesseract-OCR\tesseract.exe'
pytesseract.pytesseract.tesseract_cmd = path


class TableCleanse:
    def __init__(self, img_path):
        self.path_img = img_path
        # create temp folder for processed img
        try:
            os.mkdir(self.path_img + '/img_TableCleanse')
        except:
            shutil.rmtree(self.path_img + '/img_TableCleanse')
            os.mkdir(self.path_img + '/img_TableCleanse')

    # execute
    def execute(self):
        # read Image
        self.img = cv2.imread(self.path_img + '/temp.png')
        self.store_img("0_original.png", self.img)

        # -------------------------------------------------
        # if covered by table
        # self.format_default()
        # -------------------------------------------------    

        # add 10% padding
        self.add_padding('0.1')
        self.store_img("5_addPadding.png", self.add_padding)
        return self.add_padding

    def format_default(self):
        # gray
        self.gray = cv2.cvtColor(self.img, cv2.COLOR_BGR2GRAY)
        self.store_img("1_gray.png", self.gray)
        # threshold
        self.thresh = cv2.threshold(self.gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
        self.store_img("2_threshold.png", self.thresh)
        # invert
        self.invert = cv2.bitwise_not(self.thresh)
        self.store_img("3_invert.png", self.invert)
        # dilate
        self.dilate = cv2.dilate(self.invert, None, iterations=3)
        self.store_img("4_dilate.png", self.dilate)

    def add_padding(self, percent):
        img_height = self.img.shape[0]
        padding = int(img_height*float(percent))
        self.add_padding = cv2.copyMakeBorder(self.img, padding, padding, padding, padding, cv2.BORDER_CONSTANT, value=[255, 255, 255])

    def store_img(self, file_name, img):
        path_temp = self.path_img + '/img_TableCleanse/' + file_name
        cv2.imwrite(path_temp, img)

class LineRemove:
    def __init__(self, img, path_img):
        self.img = img
        self.path_img = path_img
        # create temp folder for processed img
        try:
            os.mkdir(self.path_img + '/img_LineRemove')
        except:
            shutil.rmtree(self.path_img + '/img_LineRemove')
            os.mkdir(self.path_img + '/img_LineRemove')
    
    # execute
    def execute(self):
        self.format_default()

        # remove any vertical lines
        self.remove_vertical_line()
        self.store_img("4_VerticalLine.png", self.vertical_line)
        # remove any horizontal lines
        self.remove_horizontal_line()
        self.store_img("5_HorizontalLine.png", self.horizontal_line)
        # combine removes
        self.combine_remove()
        # dilate
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
        self.dilate = cv2.dilate(self.combine, kernel, iterations=6)
        self.dilate = cv2.threshold(self.dilate, 0, 255, cv2.THRESH_TOZERO)[1]
        self.store_img("6_dilate.png", self.dilate)
        # remove actually :'(
        self.no_line = cv2.subtract(self.invert, self.dilate)
        self.store_img("7_noLine.png", self.no_line)
        # remove noise
        self.remove_noise()
        self.store_img("8_noise.png", self.remove_noise)
        return self.remove_noise
    
    def format_default(self):
        # gray
        self.gray = cv2.cvtColor(self.img, cv2.COLOR_BGR2GRAY)
        self.store_img("1_gray.png", self.gray)
        # threshold
        self.thresh = cv2.threshold(self.gray, 127, 255, cv2.THRESH_TOZERO)[1]
        self.store_img("2_threshold.png", self.thresh)
        # invert
        self.invert = cv2.bitwise_not(self.thresh)
        self.store_img("3_invert.png", self.invert)

    def remove_vertical_line(self):
        ver = np.array([[1],
               [1],
               [1],
               [1],
               [1],
               [1]])
        self.vertical_line = cv2.erode(self.invert, ver, iterations=10)
        self.vertical_line = cv2.dilate(self.vertical_line, ver, iterations=10)

    def remove_horizontal_line(self):
        hor = np.array([[1,1,1,1,1,1]])
        self.horizontal_line = cv2.erode(self.invert, hor, iterations=10)
        self.horizontal_line = cv2.dilate(self.horizontal_line, hor, iterations=10)


    def combine_remove(self):
        self.combine = cv2.add(self.vertical_line, self.horizontal_line)

    def remove_noise(self):
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2,2))
        self.remove_noise = cv2.erode(self.no_line, kernel, iterations = 1)
        self.remove_noise = cv2.dilate(self.remove_noise, kernel, iterations = 1)
        # self.remove_noise = cv2.morphologyEx(self.no_line, cv2.MORPH_CLOSE, kernel)

    def store_img(self, file_name, img):
        path_temp = self.path_img + '/img_LineRemove/' + file_name
        cv2.imwrite(path_temp, img)

class LineRemove_test:
    def __init__(self, img, path_img):
        self.img = img
        self.path_img = path_img
        # create temp folder for processed img
        try:
            os.mkdir(self.path_img + '/img_LineRemove_test')
        except:
            shutil.rmtree(self.path_img + '/img_LineRemove_test')
            os.mkdir(self.path_img + '/img_LineRemove_test')
    
    # execute
    def execute(self):
        self.format_default()
        self.clean_line()
        self.store_img("4_clean line.png", self.clean_line)
        return self.clean_line

    # default format
    def format_default(self):
        # gray
        self.gray = cv2.cvtColor(self.img, cv2.COLOR_BGR2GRAY)
        self.store_img("1_gray.png", self.gray)
        # threshold
        self.thresh = cv2.threshold(self.gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
        self.store_img("2_threshold.png", self.thresh)
        # invert
        self.invert = cv2.bitwise_not(self.gray)
        self.store_img("3_invert.png", self.invert)
    
    # clear all lines
    def clean_line(self):
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 5))
        self.clean_line = cv2.morphologyEx(self.invert, cv2.MORPH_OPEN, kernel)


    def store_img(self, file_name, img):
        path_temp = self.path_img + '/img_LineRemove_test/' + file_name
        cv2.imwrite(path_temp, img)

class OCRTool:
    def __init__(self, img, img_org, path_img):
        self.img_thresh = img
        self.img_org = img_org
        self.path_img = path_img
        # create temp folder for processed img
        try:
            os.mkdir(self.path_img + '/img_OCRTool')
        except:
            shutil.rmtree(self.path_img + '/img_OCRTool')
            os.mkdir(self.path_img + '/img_OCRTool')
    
    def execute(self):
        # dilate processed img
        self.dilate()
        self.store_img("1_dilate.png", self.dilate)
        # find contours
        self.find_contour()
        self.store_img("2_contours.png", self.img_contours)
        # convert contours to bounding boxes
        self.bounding_box()
        self.store_img("3_box.png", self.bounding_box)
        
    def dilate(self):
        kernel = np.array([
                [1,1,1,1,1,1,1,1,1,1],
                [1,1,1,1,1,1,1,1,1,1]
        ])
        self.dilate= cv2.dilate(self.img_thresh, kernel, iterations=3)
        simple_kernel = np.ones((3,3), np.uint8)
        self.dilate = cv2.dilate(self.dilate, simple_kernel, iterations=3)        
    def find_contour(self):
        result = cv2.findContours(self.dilate, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
        self.contours = result[0]
        self.img_contours = self.img_org.copy()
        cv2.drawContours(self.img_contours, self.contours, -1, (0, 255, 0), 2)
    def bounding_box(self):
        self.bounding_box_temp = []
        self.bounding_box = self.img_org.copy()
        for contour in self.contours:
            x, y, w, h = cv2.boundingRect(contour)
            self.bounding_box_temp.append((x, y, w, h))
            self.bounding_box = cv2.rectangle(self.bounding_box, (x, y), (x + w, y + h), (0, 255, 0), 3)
       


    def store_img(self, file_name, img):
        path_temp = self.path_img + '/img_OCRTool/' + file_name
        cv2.imwrite(path_temp, img)


path = r'C:\Users\vtung010\Downloads\Stockbiz Scraping\Python\Python'

wb = openpyxl.load_workbook('Excel/FS.xlsx')
ws_ICB = wb['SAMPLE']
image_loader = SheetImageLoader(ws_ICB)
img = image_loader.get('A2')

try:
    os.mkdir(path + "/Excel/resource")
    print("Create temp folder")
except:
    shutil.rmtree(path + "/Excel/resource")
    os.mkdir(path + "/Excel/resource")
    print("Delete existing temp folder")
    print("Create new temp folder")
img.thumbnail((900, 900))
img.save(path + '/Excel/resource/temp.png')

path_img = r'C:\Users\vtung010\Downloads\Stockbiz Scraping\Python\Python\Excel\resource'
cleanse = TableCleanse(path_img)
cleanse = cleanse.execute()
remove = LineRemove(cleanse, path_img)
remove = remove.execute()
remove_test = LineRemove_test(cleanse, path_img)
remove_test = remove_test.execute()
ocr = OCRTool(remove, cleanse, path_img)
ocr = ocr.execute()
cv2.waitKey(0)
# ---------------------------------------------------------------------------------
# # # Step 1 : Preprocessing
# img = cv2.imread('Excel/resource/temp.png')
# cv2.imshow('Img', img)
# # gray
# gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
# cv2.imshow('gray', gray)
# # thresh
# thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
# cv2.imshow('thresh', thresh)
# # invert
# invert = cv2.bitwise_not(thresh)
# cv2.imshow('invert', invert)
# # dilate
# dilate = cv2.dilate(invert, None, iterations=5)
# cv2.imshow('thresh', dilate)
# # # Step 2 : Perspective correction
# cv2.copyMakeBorder(dilate, padding, padding, padding, padding, cv2.BORDER_CONSTANT, value=[255, 255, 255])
# cv2.waitKey(0)
# ---------------------------------------------------------------------------------


# custom_config = r'--oem 1 --psm 6 -l vie'
# text = pytesseract.image_to_string(img, config=custom_config)
# rows = text.split('\n')



	

# print(rows)
# BusinessSegment = []
# segmentCount = 0
# for row in rows:
#     row_str = str(row)
#     try:
#         strBln = row_str[0]
#         # substring useless chars at beginning
#         while row_str[0].isalpha() == False:
#             row_str = row_str[1:]
#         # row = row_str
#     except:
#         continue



    # while row_str != '':
    #     str_start = 0
    #     whileBln = True
    #     charCount = 0
    #     while whileBln == True:
    #         if (row_str[charCount].isdigit() and row_str[charCount+1].isdigit()) or (row_str[charCount].isdigit() and row_str[charCount+2].isdigit()):



        
    
    # print(row)


# table_data = [row.split() for row in rows if row]

# for row in table_data:
#     print(row)






# original = img.copy()
# gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
# thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]

# # Remove text characters with morph open and contour filtering
# kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3,3))
# opening = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel, iterations=1)
# cnts = cv2.findContours(opening, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
# cnts = cnts[0] if len(cnts) == 2 else cnts[1]
# for c in cnts:
#     area = cv2.contourArea(c)
#     if area < 50:
#         cv2.drawContours(opening, [c], -1, (0,0,0), -1)

# # Repair table lines, sort contours, and extract ROI
# close = 255 - cv2.morphologyEx(opening, cv2.MORPH_CLOSE, kernel, iterations=1)
# cnts = cv2.findContours(close, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
# cnts = cnts[0] if len(cnts) == 2 else cnts[1]
# (cnts, _) = contours.sort_contours(cnts, method="top-to-bottom")
# for c in cnts:
#     area = cv2.contourArea(c)
#     if area < 50:
#         x,y,w,h = cv2.boundingRect(c)
#         cv2.rectangle(img, (x, y), (x + w, y + h), (36,255,12), -1)
#         ROI = original[y:y+h, x:x+w]

#         # Visualization
#         cv2.imshow('image', img)
#         cv2.imshow('ROI', ROI)
#         cv2.waitKey(20)

# cv2.imshow('opening', opening)
# cv2.imshow('close', close)
# cv2.imshow('image', img)
# cv2.waitKey()



# cv2.imwrite(path + '/sg-proccessed.png', img)

# text = pytesseract.image_to_string(img, lang='vie')
# print(text)
# print(img)

