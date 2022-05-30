import os

mac_m1 = True
if mac_m1:
    path = '/Users/rainyseason/winston/Workspace/python/Pycharm Project/sinewave_analyze/Evaluation/'
    path_csv = path + 'tek_csv/'
    path_excel = path + 'tek_excel/'
    path_summary = path_excel
    path_information = path + 'test infomation/'
    path_kmon = path + 'kmon_csv'



files = os.listdir(path_excel)
files = [file for file in files if file[:3] == 'tek']
files.sort()

for file in files:
    os.rename(path_excel + file, path_excel + file.split(' ')[0] + '.xlsx')