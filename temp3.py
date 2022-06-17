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

previous = 5173
abssent = []
for file in files:
    if previous - int(file.split('.xlsx')[0][3:7]) != -1:
        abssent.append(file)
    previous = int(file.split('.xlsx')[0][3:7])

print(abssent)


