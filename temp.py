import openpyxl
import pandas as pd
import numpy as np
import os

path = 'D:/work/data_analyze/'
csv_path = path + 'csv/'

filename = 'file name.xlsx'

wb = openpyxl.load_workbook(path + filename, data_only=True)

sheet_list = wb.get_sheet_names()
w1 = wb[sheet_list[0]]
print(w1)