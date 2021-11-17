#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Aug 19 21:45:49 2021

@author: renyuli
"""
import os
import re
import pandas as pds
from pathlib import Path

# =============================================================================
# Class PDFprocess contains methods to process/validate PDF file and convert
# into excel files. The methods here is not very suitable for PDF file generated
# by scanner
# =============================================================================
import PyPDF2
import tabula

class PDFprocess():
    def __init__(self,Manifestfolder):
        self.Manifestfolder=Path(Manifestfolder)
    
    def PDFvalidation (self):
        Manifestfolder=self.Manifestfolder
        ManifestList=os.listdir(Manifestfolder)
        PDFList=[]
        PDFscanList=[]
        for file in ManifestList:
            if re.search(r'.pdf',file):
               try:
                   InFile=open(Manifestfolder/ file,'rb')
                   PyPDF2.PdfFileReader(InFile,strict=False)
               except PyPDF2.utils.PdfReadError:
                   print(file,'not valid pdf file')
                   InFile.close()
               try:
                   tabula.read_pdf(Manifestfolder/ file,pages='all',lattice=True)[0]
               except IndexError:
                   print(file,'not valid pdf file')
                   PDFscanList.append(Manifestfolder+file)
               else:
                   PDFList.append(Manifestfolder/ file)
                   InFile.close()
            else:
                continue
        return PDFList, PDFscanList
    def Conversion(self,File):
         InFile=str(File)      
         OutFile=str(File).replace('.pdf','.xlsx')
         PDF=tabula.read_pdf(InFile,pages="all",lattice=True)
         TableNum=len(PDF)
         TableCount=0
         writer = pds.ExcelWriter(OutFile)
         while TableCount<TableNum:
             df=PDF[TableCount]
             TableCount=TableCount+1
             SheetNum='Sheet'+str(TableCount)            
             df.to_excel(writer,sheet_name=SheetNum)
         writer.save()
    def ScanPDFconversion(self,File):
        
        return File
        

# =============================================================================
# Class Wordprocess contains method that process/validate Word document manifest
# files, and convert into excel file. 
# =============================================================================
import docx

class Wordprocess():
    def __init__(self,Manifestfolder):
        self.Manifestfolder=Path(Manifestfolder)
    
    def Wordvalidation(self):
        Manifestfolder=self.Manifestfolder
        ManifestList=os.listdir(Manifestfolder)
        ManifestList=os.listdir(Manifestfolder)
        WordList=[]
        for file in ManifestList:
            if re.search(r'.docx',file):
                WordList.append(Manifestfolder/ file)
            elif re.search(r'.doc',file):
                WordList.append(Manifestfolder/ file)
            else:
                continue
        return WordList
    def Conversion(self,File):
        InFile=str(File)
        OutFile=str(File).replace('.docx','.xlsx')
        print('Done1')
        doc=docx.Document(InFile)
        Tables=doc.tables
        print('Done2')
        TableNum=len(Tables)
        TableCount=0
        df=[]
        writer = pds.ExcelWriter(OutFile)
        print('Done3')
        while TableCount<TableNum:
            Table=Tables[TableCount]        
            data=[]
            keys=None
            for i, row in enumerate(Table.rows): #Iterate the whole table 
                text = (cell.text for cell in row.cells) #get all the cells' text in the same row
                if i == 0:
                    keys = tuple(text) # This is the header row
                    continue
                row_data = dict(zip(keys, text))
                data.append(row_data)
            df = pds.DataFrame(data=data)
       
            TableCount=+1
            SheetNum='Sheet'+str(TableCount)            
            df.to_excel(writer,sheet_name=SheetNum)
        writer.save()






        
    
 
                
        