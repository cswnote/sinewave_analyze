import pandas as pd
import openpyxl
import os

path = 'D:/data_analyze/tek_excel/'

files = os.listdir(path)
files = [file for file in files if file[:3] == 'tek']
files.sort()

for file in files:
    src = path + file
    flist = file.split(' ')
    flist[2] = '300ohm'
    file =''
    for i in flist:
        file = file + i + ' '
    file = file.strip()
    dst = path + file
    os.rename(src, dst)
