import os
import openpyxl

path = 'D:/data_analyze/'
path_csv = path + 'tek_csv/'
path_excel = path + 'tek_excel/'
path_summary = path + 'summary/'
path_information = path + 'test information/'
path_kmon = path + 'kmon_csv/'

files = os.listdir(path_excel)
files.sort()
files = [file for file in files if file.endswith('.xlsx')]

for idx, file in enumerate(files):
    scr = os.path.join(path_excel + file)
    if idx % 4 == 0:
        file = file[:7] + ' ' + 'RFAMP_01 500ohm Ch1 leakage.xlsx'
    elif idx % 4 == 1:
        file = file[:7] + ' ' + 'RFAMP_01 500ohm Ch2 leakage.xlsx'
    elif idx % 4 == 2:
        file = file[:7] + ' ' + 'RFAMP_01 500ohm Ch3 leakage.xlsx'
    elif idx % 4 == 3:
        file = file[:7] + ' ' + 'RFAMP_01 500ohm Ch4 leakage.xlsx'

    dst = os.path.join(path_excel + file)
    os.rename(scr, dst)


