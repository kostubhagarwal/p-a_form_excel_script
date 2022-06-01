import os
import pandas as pd

#initiate variables for file and dataframe

cwd = os.getcwd()
files = os.listdir(cwd)
df = pd.DataFrame()

print (files)

#loop to read through files, and append to df

for file in files:
    if file.endswith('.xlsm'):
        x = file.cell(4,6)
       
        if x == 0:
            SheetN = "Sheet2"

        if x == 1:
            SheetN = "Sheet3"

        df=df.append(pd.read_excel(file, sheet_name = SheetN))
        df=df.fillna("")

print(df)


