# -*- coding: utf-8 -*-
"""
Created on Fri Jun  9 15:03:28 2023

@author: Lucie Lu

Execute morningstar excel add-in function to download Sustainalytics ratings

Country investment list has been manually created in Morningstar direct


"""



import win32com.client as win32 

import pywintypes

import openpyxl

import os
import re
import pandas as pd
import numpy as np
from datetime import datetime

from time import sleep

import time

from tqdm import tqdm

# Convert 'Time' column to datetime, but handle None values
def convert_to_datetime(time):
    if time is None:
        return pd.NaT  # NaT represents missing datetime
    return datetime.datetime(
    time.year, time.month, time.day,
    time.hour, time.minute, time.second
)


def get_year(time):
    if time is pd.NaT:
        return pd.NaN  # if datetime is None return NaN
    else:
        return int(time.year)  # Otherwise return year



path="F:\\Dropbox\\IO-CEL\\data and code\\CELu codes\\Morningstar\\Sustainalytics\\"

os.chdir(path)

universe=pd.read_csv("MorningstarUniverse.csv")

secids=universe['secid']

#only need these scores I will not make my life more complex
fieldlist=['Sustainalytics_ESG_Risk_Score','Sustainalytics_Environmental_Risk_Score',
           'Sustainalytics_Social_Risk_Score','Sustainalytics_Governance_Risk_Score']

# Open up Excel and make it visible
excel = win32.Dispatch('Excel.Application')
excel.Visible = True

# Set calculation mode to automatic (using the value for xlCalculationAutomatic)
# xlCalculationAutomatic = -4105
# excel.Calculation = xlCalculationAutomatic


start_year=2009

end_year=2023

years=list(range(start_year,end_year))

#Step2: need to cut the list into smaller chunks
n_identifiers=len(secids)

batch_size=500

total_batches=n_identifiers//batch_size+(1 if n_identifiers % batch_size!=0 else 0)

# save lists of identifiers as csv

for batch_index in range(total_batches):
    
    start_index=batch_index*batch_size
    end_index= (batch_index+1)*batch_size
    identifiers=secids[start_index:end_index]
    batch_number=batch_index+1
    pd.DataFrame(identifiers, columns=['secid']).to_csv(path+f"MorningstarList{batch_number}.csv", index=False)

#workbook = excel.Workbooks.Open(path+'Sustainalytics_download.xlsx')

workbook = excel.Workbooks.Add()

#new_sheet=workbook.Sheets.Add(None,workbook.Sheets(workbook.Sheets.Count))


#new_sheet.Name='Sustainalytics'

#new_sheet.Name=field.replace('Sustainalytics_','') # remove prefix for sheet name

# Download in small batches so that I could resume the process if there is a connection issue 
#total_batches=255

total_batches=286


#103 batch number
#271
for batch_index in range(270, total_batches-1):
    
    batch_number=batch_index+1
    
    #for id in identifiers:

    # use tesla as an example that do not have data for the full sample
        # id="0P0000OQN8"
    progress_percentage = batch_number/ total_batches * 100
    
    print(f"Batch {batch_number} | Progress: {progress_percentage:.2f}%")
    
    start_time=time.time() #
    
    identifiers=pd.read_csv(path+f"MorningstarList{batch_number}.csv")['secid'].tolist()
    
    identifiers_array = [[str(id)] for id in identifiers]
    
    n_id=len(identifiers)
    
    #new workbook for this batch
    workbook = excel.Workbooks.Add()

    #new_sheet=workbook.Sheets.Add(None,workbook.Sheets(workbook.Sheets.Count))
    for year in years:
      sheet=workbook.Sheets.Add()
      sheet.Name=str(year)
    
    # initialize empty batch table
    
    thisBatch=pd.DataFrame()
    
    range_string = f"{sheet.Cells(2,1).Address}:{sheet.Cells(n_id+1, 1).Address}"
    
    for year in years:
        
        print(f'Downloading {year}')
        
        sheet=workbook.Sheets(str(year))
        
        sheet.Activate()
        
        sheet.Range(range_string).Value=identifiers_array
    
        thisYear=pd.DataFrame()
        
        thisYear['secid']=identifiers
        
        thisYear['year']=year
        

        #download four scores:
            
        sheet.Range("B1").Formula=f"=@MSTS(A2:A{n_id+1},\""+"Sustainalytics_ESG_Risk_Score"+f"\",\"1/1/{year}\",\"12/31/{year}\",\"CORR=R, DATES=TRUE, ASCENDING=TRUE, FREQ=Y, DAYS=T, FILL=B, HEADERS=TRUE\")"
        
        sleep(10)
        
        sheet.Range("D1").Formula=f"=@MSTS(A2:A{n_id+1},\""+"Sustainalytics_Environmental_Risk_Score"+f"\",\"1/1/{year}\",\"12/31/{year}\",\"CORR=R, DATES=TRUE, ASCENDING=TRUE, FREQ=Y, DAYS=T, FILL=B, HEADERS=TRUE\")"
        
        sleep(10)
      
        sheet.Range("F1").Formula=f"=@MSTS(A2:A{n_id+1},\""+"Sustainalytics_Social_Risk_Score"+f"\",\"1/1/{year}\",\"12/31/{year}\",\"CORR=R, DATES=TRUE, ASCENDING=TRUE, FREQ=Y, DAYS=T, FILL=B, HEADERS=TRUE\")"
      
        sleep(10)
        
        sheet.Range("H1").Formula=f"=@MSTS(A2:A{n_id+1},\""+"Sustainalytics_Governance_Risk_Score"+f"\",\"1/1/{year}\",\"12/31/{year}\",\"CORR=R, DATES=TRUE, ASCENDING=TRUE, FREQ=Y, DAYS=T, FILL=B, HEADERS=TRUE\")"
      
        sleep(10)
        
        tmp=pd.DataFrame(sheet.Range(f"C2:C{n_id+1}").Value,columns=['ESG_Risk_Score'])
        
        thisYear['ESG_Risk_Score']=tmp['ESG_Risk_Score']
        
        tmp=pd.DataFrame(sheet.Range(f"E2:E{n_id+1}").Value,columns=['Environmental_Risk_Score'])
        
        thisYear['Environmental_Risk_Score']=tmp['Environmental_Risk_Score']
        
        tmp=pd.DataFrame(sheet.Range(f"G2:G{n_id+1}").Value,columns=['Social_Risk_Score'])
        
        thisYear['Social_Risk_Score']=tmp['Social_Risk_Score']
        
        tmp=pd.DataFrame(sheet.Range(f"I2:I{n_id+1}").Value,columns=['Governance_Risk_Score'])
        
        thisYear['Governance_Risk_Score']=tmp['Governance_Risk_Score']
        
        # coarse numerical scores
        thisYear['ESG_Risk_Score']=pd.to_numeric(thisYear['ESG_Risk_Score'], errors='coerce')
            
        thisYear['Environmental_Risk_Score']=pd.to_numeric(thisYear['Environmental_Risk_Score'], errors='coerce')
            
        thisYear['Social_Risk_Score']=pd.to_numeric(thisYear['Social_Risk_Score'], errors='coerce')
            
        thisYear['Governance_Risk_Score']=pd.to_numeric(thisYear['Governance_Risk_Score'], errors='coerce')
            
        thisYear.to_csv(path+f"Sustainalytics_Batch{batch_number}_{year}.csv", index=False)
        
        thisBatch=pd.concat([thisBatch,thisYear])
    
        
    thisBatch.to_csv(path+f"Sustainalytics_Batch{batch_number}.csv", index=False)
    
    sleep(60)
    
    workbook.SaveAs(path+f"Sustainalytics_Batch{batch_number}.xlsx")
    
    workbook.Close(True)
     
    end_time = time.time()
    
    elapsed_time=end_time-start_time
    
    print(f"Batch {batch_number} took {elapsed_time:.3f} seconds to download")
    

#Downloading takes 3 days


#Concatenate query tables to combined table 


total_batches=286

#Do not use pywin32 as it will open excel and trigger morningstar function
def excel_range(file, sheet_name, start_cell, end_cell):  
    wb = openpyxl.load_workbook(file)
    sheet=wb[sheet_name]
    data=[]
    for row in sheet[start_cell:end_cell]:
        data.append([cell.value for cell in row])
        
    df=pd.DataFrame(data)
    
    return(df)


combined_table=pd.DataFrame()

for batch_index in range(0, total_batches): 
    
    start_time=time.time() #
    
    batch_number=batch_index+1
    
    identifiers=pd.read_csv(path+f"MorningstarList{batch_number}.csv")['secid'].tolist()
    
    n_id=len(identifiers)
    
    #workbook=excel.Workbooks.Open(path+f"Sustainalytics_Batch{batch_number}.xlsx")
    
    thisBatch=pd.DataFrame()
 
    for year in years:
        
        thisYear=pd.DataFrame()
        
        thisYear['secid']=identifiers
        
        thisYear['year']=year
        
        print(f'Converting batch {batch_number} {year}')
        
        #sheet=workbook.Sheets(str(year))
        
        #sheet.Activate()
        
        file_path=path+f"Sustainalytics_Batch{batch_number}.xlsx"
        
        tmp=excel_range(file_path,sheet_name=str(year),start_cell='C2',end_cell=f"C{n_id+1}")
        
        tmp.columns=['ESG_Risk_Score']
        
        #tmp=pd.read_excel(sheet.Range(f"C2:C{n_id+1}").Value,columns=['ESG_Risk_Score']) 
        
        if tmp['ESG_Risk_Score'].notna().sum()==0:
           print('Batch {batch_number} {year} has unprocessed request')
        
        thisYear['ESG_Risk_Score']=tmp['ESG_Risk_Score']
        
        tmp=excel_range(file_path,sheet_name=str(year),start_cell='E2',end_cell=f"E{n_id+1}")
        
        tmp.columns=['Environmental_Risk_Score']
    
        if tmp['Environmental_Risk_Score'].notna().sum()==0:
          print('Batch {batch_number} {year} has unprocessed request')
        
        thisYear['Environmental_Risk_Score']=tmp['Environmental_Risk_Score']
    
        tmp=excel_range(file_path,sheet_name=str(year),start_cell='G2',end_cell=f"G{n_id+1}")
        
        tmp.columns=['Social_Risk_Score']
        
        if tmp['Social_Risk_Score'].notna().sum()==0:
          print('Batch {batch_number} {year} has unprocessed request')
        
        thisYear['Social_Risk_Score']=tmp['Social_Risk_Score']

        tmp=excel_range(file_path,sheet_name=str(year),start_cell='I2',end_cell=f"I{n_id+1}")
        
        tmp.columns=['Governance_Risk_Score']
        
        if tmp['Governance_Risk_Score'].notna().sum()==0:
          print('Batch {batch_number} {year} has unprocessed request')
        
        thisYear['Governance_Risk_Score']=tmp['Governance_Risk_Score']
        
        # coarse numerical scores
        thisYear['ESG_Risk_Score']=pd.to_numeric(thisYear['ESG_Risk_Score'], errors='coerce')
            
        thisYear['Environmental_Risk_Score']=pd.to_numeric(thisYear['Environmental_Risk_Score'], errors='coerce')
            
        thisYear['Social_Risk_Score']=pd.to_numeric(thisYear['Social_Risk_Score'], errors='coerce')
            
        thisYear['Governance_Risk_Score']=pd.to_numeric(thisYear['Governance_Risk_Score'], errors='coerce')
        
        thisBatch=pd.concat([thisBatch,thisYear])
        
    end_time = time.time()
    
    elapsed_time=end_time-start_time
    
    print(f"Batch {batch_number} took {elapsed_time:.3f} seconds to convert")
    
    thisBatch.to_csv(path+f"Sustainalytics_Batch{batch_number}.csv", index=False)
    
    combined_table=pd.concat([combined_table,thisBatch])
        
combined_table.to_csv(path+'Sustainalytics_combined.csv', index=False)  #142846 unique secid, same as universe

combined_table_clean=combined_table[~(combined_table['ESG_Risk_Score'].isna()
                                       &combined_table['Environmental_Risk_Score'].isna()
                                       &combined_table['Social_Risk_Score'].isna()
                                       &combined_table['Governance_Risk_Score'].isna())]

combined_table_clean.to_csv(path+'Sustainalytics_combined_clean.csv', index=False)

#conversion takes 2.4 hours


#House-cleaning
       
pattern = r'Sustainalytics_Batch\d+_\d+\.csv'  # Regular expression pattern to match the file name format

# List all files in the folder
files_in_folder = os.listdir(path)

# Loop through the files and delete the ones with the specified name format
for file_name in files_in_folder:
    if re.match(pattern, file_name):
        file_path = os.path.join(path, file_name)
        os.remove(file_path)
        print(f"Deleted file: {file_name}")



    
