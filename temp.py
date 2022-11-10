import pandas as pd
import openpyxl
import os
import shutil

main_path = 'D:/winston/Desktop/새 폴더/'
sub_path = ['test 86']

for sub in sub_path:
    path = main_path + sub + '/'

    files = os.listdir(path)
    for file in files:
        src = path + file

        if file[-3:] == 'csv':
            dst = 'D:/data_analyze/tek_csv/'
            shutil.copy(src, dst)
            dst = 'D:/winston/OneDrive - (주)필드큐어/4_개발과제/PL150/4_개발단계평가/2_TP/RF Board/raw data/scope/csv/' + file
            os.rename(src, dst)
        elif file[-3:] == 'set':
            dst = 'D:\\winston\\OneDrive - (주)필드큐어\\4_개발과제\\PL150\\4_개발단계평가\\2_TP\\RF Board\\raw data\\scope\\set\\' + file
            os.rename(src, dst)
        elif file[-3:] == 'png':
            dst = 'D:\\winston\\OneDrive - (주)필드큐어\\4_개발과제\\PL150\\4_개발단계평가\\2_TP\\RF Board\\raw data\\scope\\capture\\' + file
            os.rename(src, dst)