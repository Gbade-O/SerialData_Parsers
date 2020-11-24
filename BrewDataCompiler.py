# -*- coding: utf-8 -*-
"""
Created on Wed Nov 18 15:36:35 2020

@author: GOgunwumi
"""

import os 
import pandas as pd 
import matplotlib.pyplot as plt
import numpy as np
import collections as col
import xlsxwriter
from openpyxl import load_workbook
import csv 
import re 
import codecs
import glob
from openpyxl.styles import numbers

MasterFilePath = r"C:\Users\gogunwumi\Documents\Temp\Scale Testing CFP300.xlsx"

##Create Writer to MainSheet 
writer = pd.ExcelWriter(MasterFilePath, engine = 'openpyxl')
   
##read in workbook
writer.book = load_workbook(MasterFilePath)

BrewFilePath = r"C:\Users\gogunwumi\Documents\Temp\BrewData"

os.chdir(BrewFilePath) ##Change to directory for Brewfiles 

files = [ name for name in os.listdir(".") if os.path.isfile(name)]


for file in files :
    
    xls = pd.ExcelFile(file)
    sheets = xls.sheet_names
    
    for name in sheets:
        
        df = pd.read_excel(file, sheet_name = name)
        
       
        
        ##Get/Find appropriate Sheet 
        Sheet = writer.book[name]
        
        #Set revelant formats before hand 
        
        
        #Brute force copy data to destination sheet 
        
        for i in range(2, len(df)):
            
           
            #Write nearby columns in for loop 
            for column in enumerate(Sheet.iter_cols(min_row = 1, max_row = 1, min_col = 4, max_col =9 , values_only = True),start = 4):
                
                Sheet.cell(row = i, column = column[0]).value = df.iloc[i-2,column[0]-1]
                
           
               
            
       
            

writer.book.save(MasterFilePath)

        
        
        
        
        
        
 
    
    