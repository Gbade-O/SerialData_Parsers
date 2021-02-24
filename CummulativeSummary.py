# -*- coding: utf-8 -*-
"""
Created on Sat Jan 16 19:51:04 2021

@author: GOgunwumi
"""

import xlsxwriter
import os 
import pandas as pd 


#Find summary files 
MainPath = r"C:\Users\gogunwumi\Documents\Temp"

folders = [folder for folder in os.listdir(MainPath) if os.path.exists(os.path.join(MainPath,folder,"$Cummulative_Summary"))] #Filter folders based on if they have a cummulative summary folder 

files = []
info = []

#Create list of files to be processed 
for folder in folders :
    
    Path = os.path.join(MainPath,folder,"$Cummulative_Summary")
    os.chdir(Path)
    temp_files = os.listdir(Path)
    if len(temp_files)>1 : #Directory could have more than one file , sort according to most recent 
        
        temp_files.sort(key = lambda x: os.path.getmtime(x))
        files.append(temp_files[-1])
        info.append(os.path.join(Path,temp_files[-1]))
        
    else : 
        
        files.append(temp_files[0]) #Add file name to list
        info.append(os.path.join(Path,temp_files[0]))



#Create df file from excel sheets and save to new document 
df = pd.DataFrame()
SavePath = os.path.join(MainPath,"MasterSheet","AllData.xlsx")
writer = pd.ExcelWriter(SavePath, engine='xlsxwriter')

for file in info :
    
    #Read excel file into df , then write to summary sheet 
    sheet = file[34:40]
    df = pd.read_excel(file, ignore_index = True)
    
    #Clean df file
    ind = df.index[df["Max Flowrate"]<= 200].tolist()
    df = df.drop(ind)
    
    ind = df.index[df["Max Outlet Temp"]< 90].tolist()
    df = df.drop(ind)
 
    #Section to determine Macro
    Macro_ind =df.index[ df["Max Boiler temp"]>=150 ]
   
    df.to_excel(writer,sheet_name =sheet )
    

#Add charts to created worksheet
workbook = writer.book #create workbook object 

#Loop through sheets and add necessary charts 
for sheet in workbook.sheetnames :
    
    worksheet = writer.sheets[sheet] #Choose current sheet as active worksheet
    max_row = worksheet.dim_rowmax + 1
    #Create chart object 
    chart1 = workbook.add_chart({'type' : 'line'}) #Outlet water chart 
    
    chart1.add_series({'name': [sheet,0,3],'values' : [sheet,1,3,max_row,3],})
    
    chart2 = workbook.add_chart({'type' : 'line'}) #Boiler temp chart 
    
    chart2.add_series({'name': [sheet,0,4],'values' : [sheet,1,4,max_row,4],})
    
    #Set chart axes 
    chart1.set_x_axis({'name': 'Brew Number'})
    chart1.set_y_axis({'name':'Max oultet water Temp ( degC )'})
    
    chart2.set_x_axis({'name': 'Brew Number'})
    chart2.set_y_axis({'name':'Max boiler Temp ( degC )'})
    
    #insert chart into the worksheet 
    worksheet.insert_chart('M45',chart1)
    
    worksheet.insert_chart('M2',chart2)
    
    
writer.save()    
