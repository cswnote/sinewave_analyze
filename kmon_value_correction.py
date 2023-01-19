import pandas as pd
import numpy as np
import sys
import os
import platform
import gc

if __name__ == '__main__':
    if platform.platform() == 'macOS-13.1-arm64-arm-64bit':
        path = '/Volumes/Winston 2T/work/fieldcure/multi channel test data/'
    elif platform.platform() == 'Windows-10-10.0.19044-SP0':
        path = 'C:/data_analyze/'

    path_summary = path + 'tek_excel/'

    df_ctrl = pd.read_excel(path + 'eval_control.xlsx', sheet_name='kmon digit change set RFamp')
    files = list(df_ctrl['File'].dropna())
    df_ctrl = df_ctrl[['Name', 'digit']]
    df_ctrl = df_ctrl.dropna()

    for file in files:
        sheets = pd.ExcelFile(path_summary + file).sheet_names
        df = pd.read_excel(path_summary + file, sheet_name=sheets[-1])
        df_origin = df.copy()
        for i in range(len(df)):
            for j in range(len(df_ctrl)):
                if df_ctrl.at[j, 'digit'].split(' ')[0] == 'reduce':
                    df.at[i, df_ctrl.at[j, 'Name']] = df.at[i, df_ctrl.at[j, 'Name']] * int(df_ctrl.at[j, 'digit'].split(' ')[1])  * 0.1
                elif df_ctrl.at[j, 'digit'].split(' ')[0] == 'raise':
                    df.at[i, df_ctrl.at[j, 'Name']] = df.at[i, df_ctrl.at[j, 'Name']] * int(df_ctrl.at[j, 'digit'].split(' ')[1]) * 10

        filename = path_summary + file
        sheet = sheets[-1] + ' correction'
        df_write = df
        index_ = False
        if not os.path.exists(filename):  # excel.path로 변경
            with pd.ExcelWriter(filename, mode='w', engine='openpyxl') as writer:
                df_write.to_excel(writer, sheet_name=sheet, index=index_)
        else:
            with pd.ExcelWriter(filename, mode='a', engine='openpyxl') as writer:
                df_write.to_excel(writer, sheet_name=sheet, index=index_)

        del df_write, index_, sheet, filename
        gc.collect()

    print('===================')