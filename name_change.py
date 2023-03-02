import platform

import openpyxl
import os
import pandas as pd
import math

class NAME_CHANGE():
    def __init__(self, eval_file, path):
        super().__init__()

        self.path = path
        self.tek_csv = path + 'tek_csv/'
        self.tek_excel_path = path + 'tek_excel/'
        self.kmon_csv_path = path + 'kmon_csv/'
        self.test_info_path = path + 'test information/'
        self.eval_file = eval_file


    def file_name_change(self, sheet):
        path = self.path
        df_name = pd.read_excel(self.path + self.eval_file, sheet_name=sheet)

        info_files = df_name.iloc[:, 0].tolist()
        # info_files = [file for file in info_files if file[:10] == 'info_test_' and not('all' in file)]
        test_files = os.listdir(self.tek_excel_path)
        test_files = [file for file in test_files if file.endswith('xlsx') and file[:3] == 'tek']
        test_files.sort()

        for i in range(len(df_name)):
            df_test_info = pd.read_excel(self.test_info_path + df_name.at[i, 'filename'] + '.xlsx')
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
                            if type(df_name.at[i, df_name.columns[j]]) is str:
                                column = df_name.at[i, df_name.columns[j]]
                                # file = file + ' ' + str(column) + ' ' + str(df_test_info.at[idx, column])
                                file = file + ' ' + str(column).split(' ')[1] + \
                                       str(df_test_info.at[idx, column + ' ' + df_name.at[j, 'channel']])
                            elif type(df_name.at[i, df_name.columns[j]]) is float:
                                if not math.isnan(df_name.at[i, df_name.columns[j]]):
                                    column = str(df_name.at[i, df_name.columns[j]])
                                    # file = file + ' ' + str(column) + ' ' + str(df_test_info.at[idx, column])
                                    file = file + ' ' + str(column).split(' ')[1] + \
                                           str(df_test_info.at[idx, column + ' ' + df_name.at[j, 'channel']])

                        else:
                            if 'ohm' == df_name.columns[j].lower():
                                file = file + ' ' + (str(df_name.at[i, df_name.columns[j]]) + 'ohm').replace(' ', '')
                            else:
                                file = file + ' ' + str(df_name.at[i, df_name.columns[j]]).replace(' ', '')
                            # file = file + str(df_name.iat[i, j]) # 위와 동일

                    dst = path_excel + file + '.' + extension
                    os.rename(scr, dst)

if __name__ == '__main__':
    if platform.platform() == 'Windows-10-10.0.19044-SP0':
        mac_m1 = False
    if mac_m1:
        path = '/Users/rainyseason/winston/Workspace/python/Pycharm Project/sinewave_analyze/Evaluation/'
        path_csv = path + 'tek_csv/'
        path_excel = path + 'tek_excel/'
        path_summary = path_excel
        path_information = path + 'test infomation/'
        path_kmon = path + 'kmon_csv'
    else:
        path = 'C:/data_analyze/'
        path_csv = path + 'tek_csv/'
        path_excel = path + 'tek_excel/'
        path_summary = path + 'summary/'
        path_information = path + 'test information/'
        path_kmon = path + 'kmon_csv/'

    for paths, names, files in os.walk(path_excel):
        for file in files:
            print(file)



    evaluation_control_file = 'eval_control.xlsx'
    # control_file = 'eval_control.xlsx'
    name_sheet = 'file name'

    nc = NAME_CHANGE(evaluation_control_file, path)
    nc.file_name_change(name_sheet)
