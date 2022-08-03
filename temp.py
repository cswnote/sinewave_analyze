import pandas as pd
import openpyxl
import os

path = 'D:/data_analyze/tek_excel/'

files = os.listdir(path)
files = [file for file in files if file[:4] == ' tek']
# files = [file for file in files if file[-3:] == 'csv']

files.sort()

for file in files:
    src = path + file

    file = file.strip()

    dst = path + file

    os.rename(src, dst)

# for file in files:
#     extension = file[-3:]
#     scr = os.path.join(path + file)
#     if extension == 'csv':
#         dst = os.path.join(path + 'csv/' + file)
#     elif extension == 'set':
#         dst = os.path.join(path + 'set/' + file)
#     elif extension == 'png' and len(file) < 13:
#         dst = os.path.join(path + 'png/' + file)
#     else:
#         dst = os.path.join(path + 'cap/' + file)
#
#     os.rename(scr, dst)
