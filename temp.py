import openpyxl
import os

path = '/Users/rainyseason/winston/Workspace/python/Pycharm Project/sinewave_analyze/Evaluation/'
path_csv = path + 'csv/'
path_excel = path + 'excel/'
path_summary = path + 'summary/'
path_information = path + 'test information/'

file_list = os.listdir(path_information)
file_list = [file for file in file_list if file[-5:] == '.xlsx']
file_list.sort()

for file in file_list:
    wb = openpyxl.load_workbook(path_information + file)
    ws = wb.active

    ws['ac1'].value = 'CP Pwm Set Ch 1'
    ws['ad1'].value = 'CP Pwm Set Ch 2'
    ws['ae1'].value = 'CP Pwm Set Ch 3'
    ws['af1'].value = 'CP Pwm Set Ch 4'

    wb.save(path_information + file)