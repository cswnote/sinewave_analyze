import openpyxl
import os

path = 'D:/work/data_analyze/excel/'
file = 'summary.xlsx'

wb = openpyxl.load_workbook(path + file)
ws = wb.active

sheets = wb.get_sheet_names()

name = ws.title

print(sheets)
