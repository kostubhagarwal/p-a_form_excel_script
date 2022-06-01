import os
import pandas as pd
from openpyxl import load_workbook

#initiate variables for file and dataframe

cwd = os.getcwd()
files = os.listdir(cwd)
df = pd.DataFrame()

print (files)

#loop to read through files, and append to df

for file in files:
    if file.endswith('.xlsm'):

        wb = load_workbook(file)
        wb.sheetnames
        sheet = wb.active
        x = sheet["J2"].value
       
        if x == 1:
            SheetN = "Sheet3"
            df=df.append(pd.read_excel(file, sheet_name = SheetN))
        if x == 2:
            SheetN = "Sheet2"
            df=df.append(pd.read_excel(file, sheet_name = SheetN))
        
        df=df.fillna("")

print(df)
df.to_csv('test.csv', index=False)


