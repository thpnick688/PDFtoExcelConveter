# -*- coding: utf-8 -*-
"""
Created on Mon Dec 13 13:19:10 2021

@author: rli23
"""
import os
import pytesseract
from pathlib import Path
import cv2
from pdf2image import convert_from_path
from matplotlib import pyplot as plt
import re
WD=Path('C:/Users/rli23/Desktop/Renyu/Pyscripts/ConvertToExcel-main/ConvertToExcel-main')

ManifestFolder=WD/'Manifests'
os.chdir(ManifestFolder)
print(ManifestFolder)
ManifestList=os.listdir(ManifestFolder)
pytesseract.pytesseract.tesseract_cmd=r'C:\Users\rli23\AppData\Local\Programs\Tesseract-OCR\tesseract'
for file in ManifestList:
    path=ManifestFolder/file
    pages=convert_from_path(path)
    for page in pages:
        image_name = file.replace('.pdf','.jpeg')  
        page.save(image_name, "JPEG")
imageList=[]
TotalFileList=os.listdir(ManifestFolder)
for file in TotalFileList:
    if re.search('.jpeg', file):
        imageList.append(file)

def mark_region(image_path):
    
    image = cv2.imread(image_path)

    THRESHOLD_REGION_IGNORE=175
    
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    blur = cv2.GaussianBlur(gray, (9,9), 0)
    thresh = cv2.adaptiveThreshold(blur,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV,11,30)

    # Dilate to combine adjacent text contours
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (9,9))
    dilate = cv2.dilate(thresh, kernel, iterations=4)

    # Find contours, highlight text areas, and extract ROIs
    cnts = cv2.findContours(dilate, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    cnts = cnts[0] if len(cnts) == 2 else cnts[1]

    line_items_coordinates = []
    for c in cnts:
        x,y,w,h = cv2.boundingRect(c)

        if w < THRESHOLD_REGION_IGNORE or h < THRESHOLD_REGION_IGNORE:
            continue
        else:
            image = cv2.rectangle(image, (x,y), (2200, y+h), color=(255,0,255), thickness=3)
            line_items_coordinates.append([(x,y), (2200, y+h)])
            line_items_coordinates.sort()
    return image, line_items_coordinates

def change_TO_BW(image):
    ori_image=image
    gray_image=cv2.cvtColor(ori_image, cv2.COLOR_BGR2GRAY)
    (thresh,BW_image)=cv2.threshold(gray_image, 127,255,cv2.THRESH_BINARY)
    plt.figure(figsize=(20,20))
    plt.imshow(BW_image)
    
    return BW_image

for file in imageList:
    outfile=file.replace('.jpeg','.txt')
    path=ManifestFolder/file
    
    output=open(ManifestFolder/outfile,'a')
    path=path.__str__()
    image_ROI, corp_coords=mark_region(path)
    plt.figure(figsize=(20,20))
    plt.imshow(image_ROI)
    image=cv2.imread(path)
    for coord in corp_coords:
        corp_img = image[coord[0][1]:coord[1][1], coord[0][0]:coord[1][0]]
        corp_BW_img=change_TO_BW(corp_img)
        text=pytesseract.image_to_string(corp_BW_img, lang='eng', config='--psm 11 --oem 3')
        for line in text:
            output.write(line)
    output.close()
for file in imageList:
    path=ManifestFolder/file
    os.remove(path)
print('OCR done')