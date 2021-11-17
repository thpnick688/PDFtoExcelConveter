# -*- coding: utf-8 -*-
"""
Created on Fri Sep  3 21:01:30 2021

@author: meng
"""


from ConvertToExcel import PDFprocess
from ConvertToExcel import Wordprocess
import sys
import os
from pathlib import Path
# =============================================================================
# 
# =============================================================================
def path_setup(WD):
    os.chdir(WD)
    RootDIR=Path.cwd()
    Manifestfolder=RootDIR / 'Manifests'
    return Manifestfolder
def main():
    Manifestfolder=path_setup(sys.argv[1])
    PDF=PDFprocess(Manifestfolder)
    PDFList=PDF.PDFvalidation()[0]
    #PDFscanList=PDF.PDFvalidation()[1]
    
    try:
        file=PDFList[0]
        #file=PDFscanList[0]
    except IndexError:
        print('No Workable PDF file')
    else:
        for file in PDFList:
            Filepath=Path(file)
            Fname=Filepath.name
            print(Fname,'Converting')
            PDF.Conversion(file)
            print(Fname,'Done')
# =============================================================================
# The following code lines utlize the Wordprocess class to convert docx/doc file to
# Excel. If the folder contains no workable docx/doc file. The WordList will be empty and 
# Exception will happen to stop the rest code lines. 
# =============================================================================
    Word=Wordprocess(Manifestfolder)
    WordList=Word.Wordvalidation()
    try:
        file=WordList[0]
    except IndexError:
        print('No Workable Word file')
    else:
        for file in WordList:
            Filepath=Path(file)
            Fname=Filepath.name
            print(Fname,'Converting')
            Word.Conversion(file)
            print(Fname,'Done')
if __name__=='__main__':
    main()            
            
