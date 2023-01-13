import os
import pandas as pd
import numpy as np

path = 'C:/data_analyze/tek_excel/'

files = os.listdir(path)
files = [file for file in files if len(file)==21]
#
df_ch3 = pd.DataFrame()
df_ch4 = pd.DataFrame()

for file in files:
    if file[11:-5] == '08Ch3':
        if len(df_ch3) != 0:
            df = pd.read_excel(path + file)
            df_ch3 = pd.concat([df_ch3, df])
        else:
            df_ch3 = pd.read_excel(path + file)
    elif file[11:-5] == '08Ch4':
        if len(df_ch4) != 0:
            df = pd.read_excel(path + file)
            df_ch4 = pd.concat([df_ch4, df])
        else:
            df_ch4 = pd.read_excel(path + file)


df_summary = df_ch3
filename = 'summary 08 12 08Ch3.xlsx'
sheet = 'summary'
if not os.path.exists(path + filename):
    with pd.ExcelWriter(path + filename, mode='w', engine='openpyxl') as writer:
        df_summary.to_excel(writer, sheet_name=sheet, index=False)
else:
    with pd.ExcelWriter(path + filename, mode='a', engine='openpyxl') as writer:
        df_summary.to_excel(writer, sheet_name=sheet, index=False)


df_summary = df_ch4
filename = 'summary 08 12 08Ch4.xlsx'
sheet ='summary'
if not os.path.exists(path + filename):
    with pd.ExcelWriter(path + filename, mode='w', engine='openpyxl') as writer:
        df_summary.to_excel(writer, sheet_name=sheet, index=False)
else:
    with pd.ExcelWriter(path + filename, mode='a', engine='openpyxl') as writer:
        df_summary.to_excel(writer, sheet_name=sheet, index=False)


print('===============================')

