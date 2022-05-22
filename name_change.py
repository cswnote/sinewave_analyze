import openpyxl
import os
import pandas as pd

mac_m1 = True

if mac_m1:
    path = '/Users/rainyseason/winston/Workspace/python/Pycharm Project/sinewave_analyze/Evaluation/sample/'
    path_csv = path + 'csv/'
    path_excel = path + 'excel/'
    path_summary = path + 'summary/'
    path_information = path + 'test infomation/'

df_name = pd.read_excel(path + 'name.xlsx')

info_files = os.listdir(path_information)
info_files = [file for file in info_files if file[:10] == 'info_test_' and not('all' in file)]
test_files = os.listdir(path_excel)
test_files = [file for file in test_files if file.endswith('xlsx')]

for i in range(len(df_name)):
    df_test_info = pd.read_excel(path_information + df_name.at[i, 'filename'] + '.xlsx')
    start = int(df_test_info.at[0, 'filename'][3:])
    end = int(df_test_info.at[len(df_test_info) - 1, 'filename'][3:])
    for file in test_files:
        if start <= int(file.split('.')[0][3:]) <= end:
            extension = file.split('.')[1]
            file = file.split('.')[0]
            scr = path_excel + file + '.' + extension

            idx = df_test_info.index[df_test_info['filename'] == file].tolist()[0]

            for j in range(1, len(df_name.columns)):
                if 'field' in df_name.columns[j]:
                    column = df_name.at[i, df_name.columns[j]]
                    # file = file + ' ' + str(column) + ' ' + str(df_test_info.at[idx, column])
                    file = file + ' ' + str(column).split(' ')[1] + ' ' + str(df_test_info.at[idx, column])
                else:
                    if 'ohm' == df_name.columns[j].lower():
                        file = file + ' ' + str(df_name.at[i, df_name.columns[j]]) + 'ohm'
                    else:
                        file = file + ' ' + str(df_name.at[i, df_name.columns[j]])
                    # file = file + str(df_name.iat[i, j]) # 위와 동일

            dst = path_excel + file + '.' + extension
            os.rename(scr, dst)
