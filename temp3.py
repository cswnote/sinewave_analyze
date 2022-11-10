import os
import openpyxl

path = 'D:/winston/Desktop/새 폴더/test 84/excel Ch2 Ch3/'

files = os.listdir(path)
files = [file for file in files if file[:3] == 'tek']
files.sort()

for file in files:
    src = os.path.join(path + file)
    dst = path + file[:-5] + ' scope CH2 and CH3' + file[-5:]
    os.rename(src, dst)

