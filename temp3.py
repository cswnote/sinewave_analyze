import os
import pandas as pd
import openpyxl

path = 'D:\\data_analyze\\tek_excel\\'
files = os.listdir(path)
files = [file for file in files if '$' not in file]

prefix = 0

for i, file in enumerate(files):
    if prefix != file.split(' ')[0]:
        df = pd.read_excel(path + file)
        print(df)
    break
