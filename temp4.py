import pandas as pd
import numpy as np
import sys
import os

path = '/Users/rainyseason/winston/Workspace/python/Pycharm Project/sinewave_analyze/Evaluation/summary/'
file = 'summary 08Ch3 phantom I sweep.xlsx'

sheets = pd.ExcelFile(path + file).sheet_names
df_origin = pd.read_excel(path + file, sheets[-1])
columns = [col for col in df_origin.columns if ('RF Volt' in col or 'Pwm Ch' in col) and 'deviation' not in col]

df = df_origin.copy()
for col in columns:
    for i in range(len(df)):
        if 'Volt' in col:
            df.at[i, col] = df.at[i, col] * 0.1
        elif 'Pwm Ch' in col:
            df.at[i, col] = df.at[i, col] * 10

sheet_name = 'with kmon correction'
if not os.path.exists(path):
    with pd.ExcelWriter(path + file, mode='w', engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name, index=False)
        print('create sheet: ', sheet_name)
else:
    with pd.ExcelWriter(path + file, mode='a', engine='openpyxl') as writer:
        try:
            df.to_excel(writer, sheet_name, index=False)
        except:
            print('그 시트 있다 안카나: ', sheet_name)