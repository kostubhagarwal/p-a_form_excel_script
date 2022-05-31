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
        
        df=df.append(pd.read_excel(file, sheet_name = "Sheet2"))
        df=df.fillna("")

print(df)


