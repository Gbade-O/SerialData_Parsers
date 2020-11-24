# -*- coding: utf-8 -*-
"""
Created on Thu Nov 19 16:02:15 2020

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


MainPath = r"C:\Users\gogunwumi\OneDrive - SharkNinja\Projects\Coffee\CFP300\TTermData"

headers = ["TimeStamp","HeatSink","AirFryer NTC","PC NTC","Probe NTC1","Probe NTC2","High Prs Switch","Low Prs Switch","Solenoid Status","Software"]

df_temp = pd.DataFrame(np.nan,index = [0],columns = headers)


# #Create Workbook for Parsed data 
# wb = Workbook()
# ws = wb.active #Get current sheet
# ws.title = "Raw Serial Data"

file = "TestLog.csv"

with open(os.path.join(MainPath,file),'r') as csvfile :
    
    reader = csv.reader(csvfile) #Create reader object 
    
    rows = list(reader) # Import all data as list 
    
    Software_Rev = rows.pop(0)
    
    for row in rows : 
        
        if any(row):
            
            #Process Tstamp 
            df_temp.loc[0,"TimeStamp"] = row[0]
            
            for i in range(1,len(row)+1):
                
                #Process individual BDP commands 
                Command = row[i].split("$")[1][0:3]
                Data = row[i].split("$")[1][3:]
                if Command == "KN1":
                    df_temp.loc[0,"HeatSink"] = Command
                
                           
            
            
            
            
            
            
            
    
    
    
            
        
        
        
        
    
   
        
        
    
    