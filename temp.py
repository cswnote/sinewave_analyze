import pandas as pd
import openpyxl
import os

path = '/Users/rainyseason/winston/Workspace/python/Pycharm Project/sinewave_analyze/Evaluation/tek_csv'

files = os.listdir(path)
files = [file for file in files if file[:3] == 'tek']
files.sort()

absent = []
prev = 5846
for file in files:
    file = int(file.split('.csv')[0].split('tek0')[1])
    if file - prev != 1:
        absent.append(file)
    prev = file

print(absent)