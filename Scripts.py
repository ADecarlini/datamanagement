# -*- coding: utf-8 -*-
"""
edits excel files and unifies the data format.
data format shall be define in a configuration file (conf.xlsx)

conf.xlsx description:

    sheet1: cambio_de_nombre
        column 1: initial name of the column (to be changed)
        column 2: name to be considered in final file for that column
        
    sheet2: orden_de_columnas
        Column 1: final column name 
        Column 2: order of the column in final file

@author: Adecarlini
"""

# import classes
from numpy import append, insert
from pathlib import Path
import os
import openpyxl 


#constant definition

## configuration file found in root directory
CONF_FILE = 'conf.xlsx'
INPUT_DIRECTORY = '.\\input'

## configuration sheets name
NAME_SHEET = 'cambio_de_nombre'
NAME_SHEET_INITIAL_NAME_INDEX = '1'
NAME_SHEET_FINAL_NAME_INDEX = '2'

ORDER_SHEET = 'orden_de_columnas'
ORDER_SHEET_NAME_INDEX = 1
ORDER_SHEET_ORDER_INDEX = 2

# Main script
print('------------- \n')
print('current directory --> '+os.getcwd())




print('order sheet--> rows: ' + str(columnordersheet.max_row))



for i in range(1, columnordersheet.max_row, 1):
    print ('|' + str(columnordersheet.cell(column = 1, row = i).value)
        + '|' + str(columnordersheet.cell(column = 2, row = i).value)  
        + '|')


##functions

## loads all CSV and XLSX files in ./input dir
def getinputfiles(dir = INPUT_DIRECTORY):
    inputfiles = os.listdir(dir)
    return(inputfiles)
    
## retrieves the configuration options
def getconfigurationoptions(configurationfile = CONF_FILE):
    ### read configuration file
    configurationfile = openpyxl.load_workbook(CONF_FILE)
    columnordersheet = configurationfile[ORDER_SHEET]
    columnnamessheet = configurationfile[NAME_SHEET]

    ### number of rules in each sheet    
    columnordersheetlenght = configurationfile[ORDER_SHEET].max_row
    columnnamessheetlenght = configurationfile[NAME_SHEET].max_row
    
    
    namesconfig = 1
    orderconfig = 1
    return(namesconfig, orderconfig)



def returnsheetvalues(sheet, rownumber, colnumber = 2):
    temp = []
    for rowindex in range(1, rownumber,1):
        temp
        for colindex in range(1, colnumber,1):
            





def renameColumns(excelsheet):
    
    

a = columnnamessheet
tmp = [0] * a.max_row

for row in range(1,a.max_row,1):
    A = a['A'+str(row)].value
    B = a['B'+str(row)].value
    print(A+'|'+B)
    tmp[row] = [A,B]
    print(i)
    print(tmp[i])
