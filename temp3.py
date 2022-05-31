import os
import openpyxl

path = 'D:/winston/OneDrive - (주)필드큐어/data_analyze/'
path_csv = path + 'tek_csv/'
path_excel = path + 'tek_excel/'
path_summary = path + 'summary/'
path_information = path + 'test information/'
path_kmon = path + 'kmon_csv/'

files = os.listdir(path_excel)
files = [file for file in files if file[:3] == 'tek']

for file in files:
    scr = path_excel + file
    dst = path_excel + file + '.xlsx'
    os.rename(scr, dst)