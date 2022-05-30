import os
import openpyxl

path = 'D:/winston/OneDrive - (주)필드큐어/data_analyze/'
path_csv = path + 'tek_csv/'
path_excel = path + 'tek_excel/'
path_summary = path + 'summary/'
path_information = path + 'test information/'
path_kmon = path + 'kmon_csv/'

files = os.listdir(path_information)
files = [file for file in files if file[:10] == 'info_test_']

for file in files:
    wb = openpyxl.load_workbook(path_information + file)
    ws = wb.active
    if ws.cell(1, 20).value == 'Control':
        ws.cell(1, 20).value = 'Usr Control'
    print(ws.cell(1, 20).value)

    wb.save(path_information + file)
