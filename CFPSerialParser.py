# -*- coding: utf-8 -*-
"""
Created on Thu Oct 29 15:17:51 2020

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


#Get directory listing 
MainPath = r"C:\Users\gogunwumi\OneDrive - SharkNinja\Projects\Coffee\CFP300\BlownTCO"


folders = ["CFP_Unit1","CFP_Unit3"]  #Specify folders to process 

#Define data/structures to be used later 
rows =[]
headers = ["OutletTemp","Boiler Temp","WarmPlate Temp","Max Temp", "Calibrated Offset","Pump PWM","Boiler ON","PTC ON","Flow Rate","Current Block Volume","Current Total Volume","Clean Count","recipe Size","Recipe Brew","Recipe Block","Recipe Total Volume","Recipe Time"]
df = pd.DataFrame(np.nan, index=[0], columns =headers) ##Pre-define data structure 
df_summary  = pd.DataFrame(columns = ["Brew Number", "Max Outlet Temp", "Max Boiler temp", "Recipe Blocks", "Total Volume", "Max Warm Plate temp", "Max Flowrate"])
df_Final  = pd.DataFrame(columns = ["Brew Number", "Max Outlet Temp", "Max Boiler temp", "Recipe Blocks", "Total Volume", "Max Warm Plate temp", "Max Flowrate"])





#Iterate through data folders
for folder in folders :                  
    os.chdir(MainPath + "/" + folder)  #Navigate to current folder 
    files  = [ name for name in os.listdir(".") if os.path.isfile(name)]  #Get all files 
    
    directory = "Parsed"
    
    #Create folders to save files ,  if they don't exist already 
    if os.path.exists("Parsed")== False:
        os.mkdir("Parsed")
    
    if os.path.exists("Plots") == False:
        os.mkdir("Plots")
        
    

    #Iterate through files in folder 
    for file in files :
        Data =[]  
        
        
        filetype = file.split(".")[1].lower()  #What kind of file is it 
        
        #Handle csv and .txt files differently 
        if filetype == "csv":
              with open(file,'r', encoding = 'utf-8', errors = 'ignore') as csvfile :
                                   
                    csvreader = csv.reader(csvfile)
                    rows = list(csvreader)  #Read in all data at once 
                    
                    #Skip file if it contains only header row 
                    if len(rows) > 10 :
                        
                        del rows[0:5]         ##delete first 4 rows , usually junk data  
                        
                        #Process each row from csv file 
                        for row in rows :
                            #Try to parse row, if it has junk/special characters . Skip it and move on
                            try :
                                
                                data = row[0].split()
                                if 'B0' in data :
                                    Bool = 1
                                elif 'B1' in data:
                                    Bool  = 1
                                else:
                                    Bool =0
                                if Bool==1 :
                                    for k in range(0,len(data)):
                                        if k==6 or k==7:
                                            data[k] =  float(re.split('[BP]',data[k])[1])
                                        elif k<5:
                                            data[k] = float(data[k])/10
                                        else:
                                            data[k] = float(data[k])
                                    Data.append(data)
                            except:
                                
                                pass
                                        
                                
                        #Write Data matrix to dataframe structure and save to approptiate directory
                        df = pd.DataFrame(Data,columns=headers)
                        SaveName = file.split('.')[0]+"_Parsed.xlsx"
                        Brew = file.split('.')[0][3:]
                        SavePath = os.path.join(MainPath,folder,"Parsed",SaveName)
                        writer = pd.ExcelWriter(SavePath, engine='xlsxwriter')
                        df.to_excel(writer)
                        writer.save()
                        
                    
                        ##Get Critical data and save to summary file 
                        df_summary.loc[0, "Brew File"] = Brew
                        df_summary.loc[0,"Max Outlet Temp"] = df['OutletTemp'].max()
                        df_summary.loc[0,"Max Boiler temp"] = df['Boiler Temp'].max()
                        df_summary.loc[0,"Recipe Blocks"] = df['Recipe Block'].max()
                        df_summary.loc[0,"Total Volume"] = df['Recipe Total Volume'].max()
                        df_summary.loc[0,"Max Warm Plate temp"] = df['WarmPlate Temp'].max()
                        df_summary.loc[0,"Max Flowrate"] = df['Flow Rate'].max()
                        #df_summary = df_summary.astype(int)
                        df_Final = df_Final.append(df_summary,ignore_index = True)  ##DataFrame holds summary data from all files
                        
                        ##Make plot(s) and save to Plot directory
                        fig, ax1=  plt.subplots(figsize=(20, 10))
                        fig.suptitle("CFP"+ folder + " Brew file#" + Brew, fontsize = 18)
                        nrows = len(df)
                        time = np.linspace(0,nrows,num=nrows)
                        ax1 = plt.subplot(3,1,1)
                        ax1.plot(time,df["OutletTemp"],label = 'OutletTemp')
                        ax1.plot(time,df["Boiler Temp"], label = 'Boiler Temp')
                        ax1.plot(time,df["WarmPlate Temp"], label = 'Warmplate Temp')
                        ax1.plot(time,df["Pump PWM"], label = 'Pump PWM')
                        ax1.plot(time,df["Recipe Block"], label = 'Recipe Block')
                        plt.yticks(np.arange(0,160,10))
                        ax1.legend(loc =0)
                        plt.xlabel('Time(s)', fontsize = 14)
                        plt.ylabel("Temp + Pump Rates", fontsize = 14)
                        plt.title(MainPath[-6:] + " Control Response", fontsize = 18)
                        
                        ax1_2 = ax1.twinx() ##Plot on secondary axis
                        ax1_2.plot(time,df["Flow Rate"], '-k',label = 'Flow rate')
                        ax1_2.legend(loc = 'upper right')
                       
                        ax3 = plt.subplot(3,1,2)
                        
                        ax3.plot(time,df["Boiler ON"], label= 'Boiler ON')
                        ax3.plot(time, df["PTC ON"], label = 'PTC ON')
                        ax3.plot(time,df["Recipe Block"], label = 'Recipe Block')
                        ax3.legend(loc = 0)
                        
                        
                        ax2 = plt.subplot(3,1,3)
                        ax2.plot(time,df["Current Block Volume"], label = 'Current Block Volume')
                        ax2.plot(time,df["Current Total Volume"], label = 'Current Total Volume')
                        ax2.plot(time,df["Recipe Total Volume"],label = 'Recipe Total Volume')
                        plt.yticks(np.arange(0,max(df["Recipe Total Volume"])+100,100))
                        plt.xlabel('Time(s)')
                        plt.ylabel('Volume(mL)')
                        plt.title('Brew Volumes')
                        ax2.legend()
                        
                        #Save Plots to appropriate directory
                        SaveName = file.split('.')[0]
                        SavePath = os.path.join(MainPath,folder,"Plots",SaveName)
                        plt.savefig(SavePath)
                        plt.show()
                    
              
         
    
  
                  
        elif filetype == "txt":
            f = open(file,"r")
            data = f.read() #read in text file 
            lines = data.split('\n')
            lines.remove(lines[0])
            for i in range(1,len(lines)-1):
                temp = lines[i].split(' ')
                del temp[:2]
                data = temp
                for k in range(0,len(data)):
                    if k==6 or k==7:
                        data[k] =  float(re.split('[BP]',data[k])[1])
                    elif k<5:
                        data[k] = float(data[k])/10
                    else:
                        data[k] = float(data[k])
                Data.append(data)
            df = pd.DataFrame(Data,columns=headers)
            writer = pd.ExcelWriter(file.split('.')[0]+"_Parsed.xlsx", engine='xlsxwriter')
            df.to_excel(writer)
            writer.save()
        
           
            
    #Save local summary sheet , specific to the particular day 
    SaveName = "SummarySheet_ScaleTesting" + "_" +"_"+ MainPath[-6:]+"_"+ folder+ ".xlsx"
    SavePath = os.path.join(MainPath,folder,SaveName)
    writer = pd.ExcelWriter(SavePath, engine='xlsxwriter')
    df_Final.to_excel(writer,sheet_name = folder)
    writer.save()
    
    
    # ## Append MasterSheet with similar data to have comprehensive data 
    # FilePath = os.path.join(r"C:\Users\gogunwumi\Documents\Temp","MasterSheet.xlsx")
    
    # ##Create writer object to existing file 
    # writer = pd.ExcelWriter(FilePath, engine = 'openpyxl')   
    
    # #Create Workbook object that contains existing file 
    # writer.book = load_workbook(FilePath)
    
    # #Read in sheets and save to object 
    # writer.sheets = dict((ws.title,ws) for ws in writer.book.worksheets )
    
    # #Get Active sheet 
    # sheet = writer.book.active
    # reader = pd.read_excel(FilePath,sheet_name=sheet.title)
    
    # #Write df to specific sheet 
    # df_Final.to_excel(writer,sheet_name = sheet.title,index = False, header = False, startrow = len(reader)+1)
    
    # writer.close()
    
    
    
    

            
                       
            
        
        
                        
            
    
    