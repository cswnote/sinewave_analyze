import os
import openpyxl

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



files = os.listdir(path_excel)
files = [file for file in files if file[-4:] == 'xlsx']
files.sort()

for file in files:
    scr = path_excel + file
    dst = path_excel + file.split('xlsx')[0] + '.xlsx'
    os.rename(scr, dst)
