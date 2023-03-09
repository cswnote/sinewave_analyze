import os
import openpyxl

path = "D:/winston/Desktop/"
file = "temperature monitoring.xlsx"
wb = openpyxl.load_workbook(path+file)
sheets = wb.sheetnames

for i, sheet in enumerate(sheets):
    print(sheet, end=' =====> ')
    if 'no data' in sheet:
        sheets[i] = sheet = sheet.replace('no data', 'nd')

    if 'no cover' in sheet:
        sheets[i] = sheet = sheet.replace('no cover', 'nc')
    elif 'cover' in sheet:
        sheets[i] = sheet = sheet.replace('cover', 'c')

    if '10L' in sheet:
        sheets[i] = sheet = sheet.replace('10L', '10.5L')

    print(sheet)
