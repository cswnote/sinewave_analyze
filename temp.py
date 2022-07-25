import pandas as pd
import openpyxl
import os

path = 'D:/data_analyze/tek_csv/'

files = os.listdir(path)
files = [file for file in files if file[-3:] == 'png']
files.sort()

for file in files:
    scr = os.path.join(path + file)
    if 'kmon' in file:
        file = file[:-4]
        dst = os.path.join(path + file)

        # print(f"{scr} to {dst}")

        os.rename(scr, dst)


