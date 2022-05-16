import openpyxl
import pandas as pd
import numpy as np
import os
import math

path = 'D:/work/data_analyze/'
csv_path = path + 'csv/1/'
excel_path = path + 'excel/'

filename = 'file name.xlsx'

csv_list = os.listdir(csv_path)

csv_list = [file for file in csv_list if file[:3] == 'tek' and file[-3:] == 'csv']

# for file in csv_list:
#     ohm = file.split('ohm')[0]
#     ohm = file.split(' ')[-1]
#     ohm = int(round(ohm))
#
#     scr = csv_list + file
#     dst = csv_list


df = pd.read_excel(path + filename)

for file in csv_list:
    for i in range(len(df)):
        if df.iloc[i, 0] <= int(file[3:7]) <= df.iloc[i, 1]:
            scr = csv_path + file
            dst = csv_path + file[:7] + ' RFAMP_01 ' + str(df.iloc[i, 2]) + ' ' + str(df.iloc[i, 3]) + 'ohm ' + 'PWM' + str(df.iloc[i, 4]) + '.csv'
            os.rename(scr, dst)








