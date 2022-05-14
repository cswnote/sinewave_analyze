import os
import sys
import pyautogui as pag
import time
import openpyxl
import FILE_MANAGEMENT
from datetime import datetime
# import keyboard # homebrew에서 설치되었을 때는 되었는데 지금은 안됨

path = os.getcwd() + '/Evaluation/Osciloscope/'
csv_path = path + 'csv/'
# path = '/Users/rainyseason/winston/Workspace/python/Pycharm Project/autotest/Evaluation'
filelist = os.listdir(csv_path)



for file in filelist:
    if file[:3] == 'tek':
        a.append(file.split('tek')[1][:4])


filelist = [int(file.split('tek')[1].split('_kmonCap')[0]) for file in filelist if file[:3] == 'tek']

for file in filelist:
    src = os.path.join(path, file + '.png')
    dst = file + '_kmonCap.png'
    dst = os.path.join(path, dst)
    os.rename(src, dst)
