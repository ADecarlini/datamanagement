# -*- coding: utf-8 -*-
"""
Created on Fri Sep 21 21:22:28 2018

@author: TMUY
"""

import os
from openpyxl import Workbook

EXCELFORMAT = ['xls', 'xlsx']
CVS_FORMAT = ['cvs']
INPUT_DIRECTORY = '.\\input'


def getinputfiles(dir):
    inputfiles = os.listdir(dir)
    print(type(inputfiles))
    return(inputfiles)

    
def checksupportedformat(filename):
    # file extension: 3 or 4 last characters
    fileextension = [str(filename[-3:]).lower(), 
                     str(filename[-4:]).lower()]
    tmp = fileextension[0] in EXCELFORMAT or fileextension[1] in EXCELFORMAT
    return(tmp)


def saveexcelfile(sheet, path, name = 'output.xlsx'):
    wb = Workbook()
    wb.active = sheet
    wb.save(name)
    print('saving outputfile:' + name)
    return(True)
    

#####


    
    
    
    
    
    
