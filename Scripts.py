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

# import classes & functions
from readConfiguration import getinfoindex
from readConfiguration import getconfigurationfile
import os
import openpyxl 
import string
 

#constant definition
## configuration file found in root directory
CONF_FILE = 'conf.xlsx'
INPUT_DIRECTORY = '.\\input'
OUTPUT_DIRECTORY = '.\\output'
WORKING_DIRECTORY = 'F:\\Python\\projects\\gestionDeFicheros\\datamanagement'
## configuration sheets name
NAME_SHEET = 'cambio_de_nombre'
NAME_SHEET_INITIAL_NAME_INDEX = 1
NAME_SHEET_FINAL_NAME_INDEX = 2
COLUMN_INITAIL_NAME_INDEX = 0
COLUMN_FINAL_NAME_INDEX = 1
## order of the sheet configuration
ORDER_SHEET = 'orden_de_columnas'
ORDER_SHEET_NAME_INDEX = 1
ORDER_SHEET_ORDER_INDEX = 2
COLUMN_NAME_INDEX = 0
COLUMN_ORDER_INDEX = 1
## configuration info Index
NAME_INDEX = 1
ORDER_INDEX = 0

## setting working directory
os.chdir(WORKING_DIRECTORY)
print('final directory --> '+os.getcwd())

# Main script
print('-- running ----------- \n')
## load configuration
configurationfile = getconfigurationfile(WORKING_DIRECTORY, CONF_FILE,
                                         NAME_SHEET, ORDER_SHEET)
nameinformation = configurationfile[NAME_INDEX]
orderinformation = configurationfile[ORDER_INDEX]

#####
a = openpyxl.load_workbook(CONF_FILE)
a = a[NAME_SHEET]
a['A1'].value

## obtengo las letras del abecedario y accedo al excel ( a cada columna)
for columnletter in string.ascii_uppercase[:a.max_column]:  
    print(columnletter)
    print( a[columnletter + '1'].value)

print("hecho")



