import pandas as pd
import openpyxl
import os
import shutil
import platform

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
    path = 'D:/winston/OneDrive - (주)필드큐어/data_analyze/'
    path_csv = path + 'tek_csv/'
    path_excel = path + 'tek_excel/'
    path_summary = path + 'summary/'
    path_information = path + 'test information/'
    path_kmon = path + 'kmon_csv/'

evaluation_control_file = 'eval_control.xlsx'

files = os.listdir(path_excel)
files = [file for file in files if '.DS' not in file and '~' not in file]
files.sort()

files = ['tek00000.xlsx', 'tek00009.xlsx', 'tek00010.xlsx', 'tek00013.xlsx', 'tek00022.xlsx', 'tek00023.xlsx', 'tek00026.xlsx', 'tek00035.xlsx', 'tek00036.xlsx']

for file in files:
    df = pd.read_excel(path_excel + file)
    df.columns
    print('===================')