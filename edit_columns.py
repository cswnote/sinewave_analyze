import math
import os
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

    def remove_columns(self):
        df_ctrl = pd.ExcelFile(self.path + self.eval_file)
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

            file_name = 'summary.xlsx'
            sheet_name = sheet
            df_summary = pd.read_excel(self.tek_excel_path + file_name, sheet_name=sheet)
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

            df_summary.drop(drop_list, axis=1, inplace=True)

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
                    if df_summary.at[i, 'Curr'] == '-':
                        df_summary.drop([i], inplace=True)
            elif 'pwm' in remain_items:
                for i in range(len(df_summary) - 1, -1, -1):
                    if df_summary.at[i, 'PWM'] == '-':
                        df_summary.drop([i], inplace=True)

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
                        df_summary.at[i, column] = df_summary.at[i, column] * 0.1
                    except:
                        print('do not calculate, {}, data type is {}'.format(df_summary.at[i, column],
                                                                             type(df_summary.at[i, column])))

            if not os.path.exists(self.tek_excel_path + file_name):  # excel.path로 변경
                with pd.ExcelWriter(self.tek_excel_path + file_name, mode='w', engine='openpyxl') as writer:
                    df_summary.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                with pd.ExcelWriter(self.tek_excel_path + file_name, mode='a', engine='openpyxl') as writer:
                    df_summary.to_excel(writer, sheet_name=sheet_name, index=False)


    def gather_scope_value_by_ctrl_set(self):
        summary = pd.ExcelFile(self.tek_excel_path + 'summary.xlsx')
        sheets = summary.sheet_names

        sheets = sheets[2:]

        for idx, sheet in enumerate(sheets):
            df_summary = summary.parse(sheet_name=sheet)

            # remain_columns = ['Board', 'ohm', 'Vpeak[V]', 'Irms[mA]']
            remain_columns = ['Board', 'ohm', 'Irms[mA]']
            delete_columns = list(df_summary.columns)

            if 'curr' in sheet.lower():
                remain_columns.insert(2, 'Curr[mA]')
            if 'volt' in sheet.lower():
                remain_columns.append(2, 'Volt[V]')

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
                elif before != df_summary.at[i, 'Ch']:
                    end = i - 1
                    break
                elif before == df_summary.at[i, 'Ch'] and i == len(df_summary) - 1:
                    end = i

            for i in range(start, end + 1):
                df.loc[i - start] = blank_value

            for idx, column in enumerate(remain_columns):
                if idx == 3:
                    break
                for i in range(start, end + 1):
                    a = df_summary.at[i, column]
                    df.at[i - start, column] = df_summary.at[i, column]

                channels = []
                start = []
                end = []
                before = ''
                for i in range(len(df_summary)):
                    if before == '':
                        start.append(i)
                        channels.append(df_summary.at[i, 'Ch'])
                        before = df_summary.at[i, 'Ch']
                    elif before != df_summary.at[i, 'Ch']:
                        start.append(i)
                        channels.append(df_summary.at[i, 'Ch'])
                        end.append(i - 1)
                        before = df_summary.at[i, 'Ch']
                    elif before == df_summary.at[i, 'Ch'] and i == len(df_summary) - 1:
                        end.append(i)



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
        path = 'D:/winston/OneDrive - (주)필드큐어/data_analyze/'
        path_csv = path + 'tek_csv/'
        path_excel = path + 'tek_excel/'
        path_summary = path + 'summary/'
        path_information = path + 'test information/'
        path_kmon = path + 'kmon_csv/'

    evaluation_control_file = 'eval_control.xlsx'

    sum = Get_summary(path, evaluation_control_file)
    # sum.remove_columns()
    sum.gather_scope_value_by_ctrl_set()
