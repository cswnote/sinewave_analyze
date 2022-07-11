import pandas as pd
import openpyxl
import os

path = '/Volumes/16G/'

files = os.listdir(path)
files = [file for file in files if file[:3] == 'tek']
files.sort()

base_num = -1776
for file in files:
    scr = path + file
    extenstion = file.split('.')[1]
    filenum = file.split('tek')[1] .split('.')[0]

    filenum = int(filenum) + base_num
    filenum = '{:05d}'.format(filenum)

    dst = path + 'tek' + filenum + '.' + extenstion

    os.rename(scr, dst)