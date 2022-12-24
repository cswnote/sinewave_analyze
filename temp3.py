import pandas as pd
import math
import os
import openpyxl
import platform
import numpy as np
import pandas as pd

if platform.platform()[:3].lower() == 'mac':
    mac_m1 = True
elif platform.platform()[:3].lower() == 'win':
    mac_m1 = False

if mac_m1:
    path = '/Users/rainyseason/winston/Workspace/python/Pycharm Project/sinewave_analyze/Evaluation/'
    path_csv = path + 'tek_csv/'
    path_excel = path + 'tek_excel/'
    path_summary = path + 'summary/'
    path_information = path + 'test information/'
    path_kmon = path + 'kmon_csv'
else:
    path = 'D:/data_analyze/'
    path_csv = path + 'tek_csv/'
    path_excel = path + 'tek_excel/'
    path_summary = path + 'summary/'
    path_information = path + 'test information/'
    path_kmon = path + 'kmon_csv/'

evaluation_control_file = 'eval_control.xlsx'

files = os.listdir(path_excel)
files = [file for file in files if 'summary' in file ]
files.sort()

sheets = pd.ExcelFile(path_excel + files[-1]).sheet_names
print(sheets)
df = pd.read_excel(path_excel + files[-1], sheet_name=sheets[-1])
df.drop([0, 11], axis=0, inplace=True)
df.reset_index(inplace=True)
df.drop(['index'], axis=1, inplace=True)

sheets = pd.ExcelFile(path_excel + files[1]).sheet_names
print(sheets)
df_ch3 = pd.read_excel(path_excel + files[0], sheet_name=sheets[-1])
df_ch4 = pd.read_excel(path_excel + files[1], sheet_name=sheets[-1])

df_phantom = df.iloc[:13, :]
df_R_non_inter = df.iloc[13:26, :]
df_R = df.iloc[26:, :]

df_R.reset_index(inplace=True)
df_R = df_R.drop(['index'], axis=1)
df_R_non_inter.reset_index(inplace=True)
df_R_non_inter = df_R_non_inter.drop(['index'], axis=1)

print('=============')
print("{:.2f} = {:.2f} / {:.2f} Ch3 {6.1f}mA Ch4 {6.1f}mA".format(resistor, df_phantom.at[i, 'Vp_ch3'] / math.sqrt(2), df_phantom.at[i, 'Irms_ch3'] * 10 ** -3, df_phantom.at[i, 'Irms_ch3'], df_phantom.at[i, 'Irms_ch4']))