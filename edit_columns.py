import math
import os
import gc
import numpy as np
import openpyxl
import platform
import pandas as pd


class Get_summary():
    def __init__(self, path, eval_file):
        super().__init__()

        self.path = path
        # self.tek_excel_path = path + 'tek_excel/auto mode test/'
        self.tek_csv_path = path + 'tek_csv/'
        self.tek_excel_path = path + 'tek_excel/'
        self.kmon_csv_path = path + 'kmon_csv/'
        self.test_info_path = path + 'test information/'
        self.eval_file = eval_file

        self.measure_value = {'filename': []}
        self.lost_files = []

    def remove_columns(self, file):
        df_ctrl = pd.ExcelFile(self.path + self.eval_file)
        df = pd.ExcelFile(self.tek_excel_path + file)
        # sheets = df.sheet_names[2:]
        del [[df]]
        gc.collect()
        df_ctrl = df_ctrl.parse('edit columns')

        sheets = list(df_ctrl.iloc[:, 0])

        for idx, sheet in enumerate(sheets):
            remain_items = list(df_ctrl.iloc[idx, 1:])

            for i in range(len(remain_items) - 1, -1, -1):
                try:
                    if math.isnan(remain_items[i]):
                        del remain_items[i]
                except:
                    pass

            remain_items = [item.lower() for item in remain_items if True]

            file_name = file
            sheet_name = sheet
            print('in process: ', sheet)
            df_summary = pd.read_excel(self.tek_excel_path + file_name, sheet_name=sheet)
            if len(df_summary) != 0:
            # if True:
                for item in remain_items:
                    if 'curr' in item.lower():
                        remain_items.append('vp')
                        remain_items.append('irms')
                        sheet_name = sheet_name + ' ' + 'Curr'
                    if 'volt' in item.lower():
                        remain_items.append('vp')
                        remain_items.append('irms')
                        sheet_name = sheet_name + ' ' + 'Volt'

                drop_list = list(df_summary.columns)[7:]

                for item in remain_items:
                    for i in range(len(drop_list) - 1, -1 , -1):
                        if item in drop_list[i].lower().replace(' ', ''):
                            drop_list.remove(drop_list[i])

                if 'curr' in remain_items and 'volt' in remain_items:
                    for i in range(len(df_summary) - 1, -1, -1):
                        if df_summary.at[i, 'Curr'] == '-' and df_summary.at[i, 'Volt'] == '-':
                            df_summary.drop([i], inplace=True)
                elif 'curr' in remain_items:
                    for i in range(len(df_summary) - 1, -1, -1):
                        if df_summary.at[i, 'Curr'] == '-':
                            df_summary.drop([i], inplace=True)
                elif 'volt' in remain_items:
                    for i in range(len(df_summary) - 1, -1, -1):
                        if df_summary.at[i, 'Volt'] == '-':
                            df_summary.drop([i], inplace=True)
                elif 'pwm' in remain_items:
                    for i in range(len(df_summary) - 1, -1, -1):
                        if df_summary.at[i, 'PWM'] == '-':
                            df_summary.drop([i], inplace=True)

                df_summary.drop(drop_list, axis=1, inplace=True)
                df_summary.reset_index(inplace=True)
                df_summary.drop(['index'], axis=1, inplace=True)

                calculate_first_col = 'Volt'
                cal_list = list(df_summary.columns)
                cal_list = cal_list[cal_list.index(calculate_first_col):]

                for col in reversed(cal_list):
                    if 'deviation' in col or '[mA' in col or '[mV' in col or '[A' in col or '[V' in col:
                        cal_list.remove(col)

                for column in cal_list:
                    if 'curr' in column.lower():
                        df_summary.rename(columns={column: column + '[mA]'}, inplace=True)
                        column = column + '[mA]'
                    elif 'volt' in column.lower():
                        df_summary.rename(columns={column: column + '[V]'}, inplace=True)
                        column = column + '[V]'

                    for i in range(len(df_summary)):
                        try:
                            df_summary.at[i, column] = int(df_summary.at[i, column]) * 0.1
                        except:
                            print('do not calculate, {}, data type is {}'.format(df_summary.at[i, column],
                                                                                 type(df_summary.at[i, column])))

                if not os.path.exists(self.tek_excel_path + file_name):  # excel.path로 변경
                    with pd.ExcelWriter(self.tek_excel_path + file_name, mode='w', engine='openpyxl') as writer:
                        df_summary.to_excel(writer, sheet_name=sheet_name, index=False)
                else:
                    with pd.ExcelWriter(self.tek_excel_path + file_name, mode='a', engine='openpyxl') as writer:
                        try:
                            df_summary.to_excel(writer, sheet_name=sheet_name, index=False)
                        except:
                            print('그 시트 있다이: {}'.format(sheet_name))


    def gather_scope_value_by_ctrl_set(self, gather_cols, gather_sheets, file):
        filename = file
        # summary = pd.ExcelFile(self.tek_excel_path + filename)
        sheets = gather_sheets[10:16]

        # sheets = sheets[2:]

        for idx, sheet in enumerate(sheets):
            print(idx, sheet)
            # df_summary = summary.parse(sheet_name=sheet)
            df_summary = pd.read_excel(self.tek_excel_path + filename, sheet_name=sheet)

            remain_columns = ['Board', 'ohm', 'PWM', 'Volt[V]', 'Curr[mA]']
            # remain_columns = ['Board', 'ohm']
            added_item = []

            # for col in gather_cols:
            #     remain_columns.append(col)

            delete_columns = list(df_summary.columns)

            if 'curr' in sheet.lower():
                # remain_columns.insert(2, 'Curr[mA]')
                added_item.append('Curr[mA]')
                sheet_name = sheet + ' ' + 'All Ch'
            if 'volt' in sheet.lower():
                # remain_columns.insert(2, 'Volt[V]')
                added_item.append('Volt[V]')
                sheet_name = sheet + ' ' + 'All Ch'
            if 'pwm' in sheet.lower():
                # remain_columns.insert(2, 'PWM')
                added_item.append('PWM')
                sheet_name = sheet + ' ' + 'All Ch'


            df = pd.DataFrame(columns=remain_columns)

            blank_value = []
            for i in range(len(df.columns)):
                blank_value.append('')

            max_length_ch = df_summary['Ch'].value_counts()
            max_ch = max_length_ch.index[0]

            max_length_ch = df_summary['Ch'].value_counts()
            max_length_ch = max_length_ch.index[0]

            start = 0
            end = 0
            before = ''
            for i in range(len(df_summary)):
                if df_summary.at[i, 'Ch'] == max_ch and before == '':
                    start = i
                    before = df_summary.at[i, 'Ch']
                elif before != df_summary.at[i, 'Ch'] and before != '':
                    end = i - 1
                    break
                elif before == df_summary.at[i, 'Ch'] and i == len(df_summary) - 1:
                    end = i

            for i in range(start, end + 1):
                df.loc[i - start] = blank_value

            for index, column in enumerate(remain_columns):
                for i in range(start, end + 1):
                    a = df_summary.at[i, column]
                    df.at[i - start, column] = df_summary.at[i, column]

            # dict = {}
            add_columns = []
            channels = df_summary['Ch'].unique()
            channels.sort()
            for col in gather_cols[idx]:
                for ch in channels:
                    # dict.setdefault(col + ch, [])
                    add_columns.append(col + ch)
            # channels = list(dict.keys())
            channels = add_columns

            for col in channels:
                df[col] = np.nan

            remain_columns = remain_columns[2:]
            row = 0
            inner_row = 0
            while True:
                for col in channels:
                    if df_summary.at[row, 'Ch'] in col:
                        column = col.split(df_summary.at[row, 'Ch'])[0]
                        while True:
                            same_flag = True
                            for set_value in remain_columns:
                                if df.at[inner_row, set_value] != df_summary.at[row, set_value]:
                                    same_flag = False
                            inner_row += 1
                            if same_flag:
                                df.at[inner_row - 1, col] = df_summary.at[row, column]
                                break
                            elif inner_row == end - start + 1:
                                break
                if inner_row == end - start + 1:
                    inner_row = 0
                row += 1
                if row == len(df_summary):
                    break

            # filename = 'summary derivation.xlsx'
            if not os.path.exists(self.tek_excel_path + filename):  # excel.path로 변경
                with pd.ExcelWriter(self.tek_excel_path + filename, mode='w', engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    print('create sheet: ', sheet_name)
            else:
                with pd.ExcelWriter(self.tek_excel_path + filename, mode='a', engine='openpyxl') as writer:
                    try:
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                    except:
                        print('그 시트 있다 안카나: ', sheet_name)



if __name__ == '__main__':
    if platform.platform()[:3].lower() == 'mac':
        mac_m1 = True
    elif platform.platform()[:3].lower() == 'win':
        mac_m1 = False

    if mac_m1:
        path = '/Users/rainyseason/winston/Workspace/python/Pycharm Project/sinewave_analyze/Evaluation/'
        path_csv = path + 'tek_csv/'
        path_excel = path + 'tek_excel/'
        path_summary = path + 'summary/'
        path_information = path + 'test information/'
        path_kmon = path + 'kmon_csv'
    else:
        path = 'D:/data_analyze/'
        path_csv = path + 'tek_csv/'
        path_excel = path + 'tek_excel/'
        path_summary = path + 'summary/'
        path_information = path + 'test information/'
        path_kmon = path + 'kmon_csv/'

    evaluation_control_file = 'eval_control.xlsx'

    sum = Get_summary(path, evaluation_control_file)
    file = 'summary 500ohm.xlsx'
    # sum.remove_columns(file)

    df = pd.ExcelFile(path_excel + file)
    gather_sheets = df.sheet_names
    # gather_sheets = [sheet for sheet in gather_sheets if len(sheet) > 18]
    del [[df]]
    gc.collect()
    gather_cols = [['Vpeak[V]'], ['Irms[mA]'], ['Vpeak[V]'], ['Irms[mA]']]
    sum.gather_scope_value_by_ctrl_set(gather_cols, gather_sheets, file)
