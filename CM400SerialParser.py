# -*- coding: utf-8 -*-
"""
Created on Mon Nov 16 13:01:24 2020

@author: GOgunwumi
"""

import os 
import pandas as pd 
import matplotlib.pyplot as plt
import numpy as np
import collections as col
import xlsxwriter
import csv 
import re 
import codecs



headers = ["Tout","Tboil","Tptc","MaxT", "OffsetT","Pwm","Boil","Ptc","FRate","CBV","CTV","CLLCNT","RStP","RecB","Blok","Vtal","TIME","Brew","BrewSet","T","V","T","n","BSTP4"]
df_summary  = pd.DataFrame(columns = ["Brew Number", "Max Outlet Temp", "Max Boiler temp", "Recipe Blocks", "Total Volume", "Max Warm Plate temp", "Max Flowrate"])
df_Final  = pd.DataFrame(columns = ["Brew Number", "Max Outlet Temp", "Max Boiler temp", "Recipe Blocks", "Total Volume", "Max Warm Plate temp", "Max Flowrate"])


## Folder director 
MainPath = r"C:\Users\gogunwumi\Documents\Temp\Unit_1"
folders  = ["1117"]

for folder in folders :
    
    os.chdir(MainPath + "/" + folder)
    
    if os.path.exists("Parsed") == False :
        os.mkdir("Parsed")
    if os.path.exists("Plots") == False : 
        os.mkdir("Plots")
        
    files = [ name for name in os.listdir(".") if os.path.isfile(name) ]
    
    for file in files :
        with open(file,'r') as csvfile :
            df = pd.DataFrame(np.nan, index=[0], columns =headers) ##Pre-define data structure
            reader = csv.reader(csvfile)
            count =0
            Brew = file.split('.')[0][4:]
            for i in range(10):
                next(reader)
            for row in reader : 
                if any(row):
                    
                    data = row[0].split()
                    for header in data :
                        if header == "BrewSet":   ##Handle "BrewSet" differently , since it can't be stripped 
                            df.loc[count,header] =float(0)
                        elif "BSTP" in header :   ##Handle "BSTP" differently , since it can't be stripped
                            df.loc[count,"BSTP4"] = header
                            count = count +1 ##Only increment row at the end of serial stream
                        else:
                            try:
                                
                                items =header.split(":")
                                df.loc[count,items[0]] = float(items[1])   ##Split data by semi colon , convert to float and assign to dataframe
                            except : 
                                pass
        
        
        
        #Write DataFrame to excel file 
        SaveName = file.split(".")[0] + "_Parsed.xlsx"
        SaveAddress = os.path.join(MainPath,folder,"Parsed",SaveName)
        writer = pd.ExcelWriter(SaveAddress, engine = 'xlsxwriter')
        df.to_excel(writer)
        writer.save()
                    
        
        ##Get Critical data and save to summary file 
        df_summary.loc[i-1, "Brew File"] = Brew
        df_summary.loc[i-1,"Max Outlet Temp"] = df['Tout'].max()
        df_summary.loc[i-1,"Max Boiler temp"] = df['Tboil'].max()
        df_summary.loc[i-1,"Recipe Blocks"] = df['RecB'].max()
        df_summary.loc[i-1,"Total Volume"] = df['Vtal'].max()
        df_summary.loc[i-1,"Max Warm Plate temp"] = df['Tptc'].max()
        df_summary.loc[i-1,"Max Flowrate"] = df['FRate'].max()
        #df_summary = df_summary.astype(int)
        df_Final = df_Final.append(df_summary,ignore_index = True)
        
        # ##Make plot(s) and save to Plot directory
        # fig, ax1=  plt.subplots(figsize=(20, 10))
        # fig.suptitle("CM"+ folder + " Brew file#" + Brew, fontsize = 18)
        # nrows = len(df)
        # time = np.linspace(0,nrows,num=nrows)
        # ax1 = plt.subplot(3,1,1)
        # ax1.plot(time,df["Tout"],label = 'OutletTemp')
        # ax1.plot(time,df["Tboil"], label = 'Boiler Temp')
        # ax1.plot(time,df["Tptc"], label = 'Warmplate Temp')
        # ax1.plot(time,df["Pwm"], label = 'Pump PWM')
        # ax1.plot(time,df["RecB"], label = 'Recipe Block')
        # plt.yticks(np.arange(0,160,10))
        # ax1.legend(loc =0)
        # plt.xlabel('Time(s)', fontsize = 14)
        # plt.ylabel("Temp + Pump Rates", fontsize = 14)
        # plt.title(MainPath[-6:] + " Control Response", fontsize = 18)
        
        # ax1_2 = ax1.twinx() ##Plot on secondary axis
        # ax1_2.plot(time,df["FRate"], '-k',label = 'Flow rate')
        # ax1_2.legend(loc = 'upper right')
       
        # ax3 = plt.subplot(3,1,2)
        
        # ax3.plot(time,df["Boil"], label= 'Boiler ON')
        # ax3.plot(time, df["Ptc"], label = 'PTC ON')
        # ax3.plot(time,df["RecB"], label = 'Recipe Block')
        # ax3.legend(loc = 0)
        
        
        # ax2 = plt.subplot(3,1,3)
        # ax2.plot(time,df["CBV"], label = 'Current Block Volume')
        # ax2.plot(time,df["CTV"], label = 'Current Total Volume')
        # ax2.plot(time,df["Vtal"],label = 'Recipe Total Volume')
        # plt.yticks(np.arange(0,max(df["Vtal"])+100,100))
        # plt.xlabel('Time(s)')
        # plt.ylabel('Volume(mL)')
        # plt.title('Brew Volumes')
        # ax2.legend()
        
        # SaveName = file.split('.')[0]
        # SavePath = os.path.join(MainPath,folder,"Plots",SaveName)
        # plt.savefig(SavePath)
        # plt.show()        
                
                        
                    
     ##Save Final Summary df to MasterWorksheet under main path
    SaveName = "SummarySheet_ScaleTesting" + "_" +"_"+ MainPath[-6:]+"_"+ folder+ ".xlsx"
    SavePath = os.path.join(MainPath,folder,SaveName)
    writer = pd.ExcelWriter(SavePath, engine='xlsxwriter')
    df_Final.to_excel(writer,sheet_name = folder)
    writer.save()               
            
                
            

