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


MainPath = r"C:\Users\gogunwumi\OneDrive - SharkNinja\Projects\OneLid\SystemResponse\Reheat"

headers = ["TimeStamp","HeatSink","AirFryer NTC","PC NTC","Probe NTC1","Probe NTC2","High Prs Switch","Low Prs Switch","Solenoid Status","Motor Setting","Motor RPM","FAN TRiAC","AF TRIAC"]

df_temp = pd.DataFrame(np.nan,index = [0],columns = headers)
df_final = pd.DataFrame(np.nan,index = [0],columns = headers)

file = "Reheat_12_3_20.csv"

##Read in DAQ Data and add to df structure 
DAQ_Path = r"C:\Users\gogunwumi\OneDrive - SharkNinja\Projects\OneLid\SystemResponse\SteamBake\201202-170537.csv"

with open(DAQ_Path,'r') as csvfile :
    
    reader = csv.reader(csvfile) #reader object 
    
    for i in range(1,25): 
        next(reader)
   
    DAQ_headers = next(reader)
    next(reader)
    DAQ_rows = list(reader) #Import all data as a list 
    DAQ_df = pd.DataFrame(DAQ_rows,columns = DAQ_headers)
    DAQ_df.drop(columns = ['Number','ms','Alarm1-10','Alarm11-20','AlarmOut'],axis = 1, inplace = True)
    
# #Create Workbook for Parsed data 
# wb = Workbook()
# ws = wb.active #Get current sheet
# ws.title = "Raw Serial Data"



with open(os.path.join(MainPath,file),'r') as csvfile :
    
    reader = csv.reader(csvfile) #Create reader object 
    
    rows = list(reader) # Import all data as list 
    
    Software_Rev = rows.pop(0)
    
    for row in rows : 
        
        if any(row):
            
            #Process Tstamp 
            df_temp.loc[0,"TimeStamp"] = row[0]
            
            for i in range(1,len(row)):
                
                #Process individual BDP commands 
                try :
                    Command = row[i].split("$")[1][0:3]
                    Data = row[i].split("$")[1][4:]
                    print(Command)
                
                    if Command == "KN1":
                        df_temp.loc[0,"HeatSink"] = float(str(int(Data,16)))/10  #Very ugly way to convert hex to float 
                    if Command == "KN2" :
                        df_temp.iloc[0,2] = float(str(int(Data,16)))/10
                    if Command == "KN3":
                        df_temp.iloc[0,3] = float(str(int(Data,16)))/10
                    if Command == "KN4":
                        df_temp.iloc[0,4] = float(str(int(Data,16)))/10
                    if Command == "KN5":
                        df_temp.iloc[0,5] = float(str(int(Data,16)))/10
                    if Command == "SW2":
                        df_temp.iloc[0,6] = int(row[i].split("$")[1][3:])
                    if Command == "SW1":
                        df_temp.iloc[0,7] = int(row[i].split("$")[1][3:])
                    if Command == "KR3":
                        df_temp.iloc[0,8] = int(row[i].split("$")[1][3:])
                    if Command == "KM1":
                        df_temp.iloc[0,9] = int(row[i].split("$")[1][3:],16)
                    if Command == "KP1":
                        df_temp.iloc[0,10] = int(row[i].split("$")[1][3:],16)
                    if Command == "KT1":
                        df_temp.iloc[0,11] = int(row[i].split("$")[1][3:],16)
                    if Command == "KT2":
                        df_temp.iloc[0,12] = int(row[i].split("$")[1][3:],16)
                except:
                    pass
                
            df_final = df_final.append(df_temp,ignore_index=True )
                           
            
    
    #Save local summary sheet , specific to the particular day 
    SaveName = file + "_Parsed.xlsx"
    SavePath = os.path.join(MainPath,SaveName)
    writer = pd.ExcelWriter(SavePath, engine='xlsxwriter')
    df_final.to_excel(writer)
    writer.save()        
            
            
            
            
            
            
    
    
    
            
        
        
        
        
    
   
        
        
    
    