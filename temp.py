import pandas as pd
import openpyxl
import os
import shutil

path = '/Users/rainyseason/winston/Workspace/python/Pycharm Project/sinewave_analyze/Evaluation/tek_excel/'


files = os.listdir(path)
files.sort()
for file in files:
    src = path + file

    if file[:3] == 'tek' and file[-4:] == 'xlsx':
        dst = path + file[:-5] + ' Scope CH2 and CH3' + file[-5:]
        os.rename(src, dst)