import pandas as pd
import openpyxl
import os

path = 'D:\\data_analyze\\tek_excel'

files = os.listdir(path)
files = [file for file in files if file.endswith('.xlsx')]

for idx, file in enumerate(files):
    df = pd.read_excel(path + file)

