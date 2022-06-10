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
        self.tek_excel_path = path + 'tek_excel/temp/'
        self.kmon_csv_path = path + 'kmon_csv/'
        self.test_info_path = path + 'test information/'
        self.eval_file = eval_file

        self.measure_value = {'filename': []}
        self.lost_files = []

    def remain_columns(self):
        df_ctrl = pd.read_excel(self.path + self.eval_file, sheet_name='edit columns')

        sheets = list(df_ctrl.iloc[:, 0])
        sheet_name = ''

        for idx, sheet in enumerate(sheets):
            remain_items = list(df_ctrl.iloc[idx, 1:])

            for i in range(len(remain_items)):
                try:
                    if math.isnan(remain_items[i]):
                        del remain_items[i]
                except:
                    pass

            file_name = 'summary.xlsx'
            sheet_name = sheet
            df_sum = pd.read_excel(self.tek_excel_path + file_name, sheet_name=sheet)
            for item in remain_items:
                    if 'curr' in item.lower():
                        remain_items.append('irms')
                        sheet_name = sheet_name + ' ' + 'Curr'
                    if 'volt' in item.lower():
                        remain_items.append('vp')
                        sheet_name = sheet_name + ' ' + 'Volt'

            drop_list = list(df_sum.columns)[7:]

            for item in remain_items:
                for i in range(len(drop_list) - 1, -1 , -1):
                    if item in drop_list[i].lower().replace(' ', ''):
                        drop_list.remove(drop_list[i])

            df_sum.drop(drop_list, axis=1, inplace=True)

            if not os.path.exists(self.tek_excel_path + file_name):  # excel.path로 변경
                with pd.ExcelWriter(self.tek_excel_path + file_name, mode='w', engine='openpyxl') as writer:
                    df_sum.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                with pd.ExcelWriter(self.tek_excel_path + file_name, mode='a', engine='openpyxl') as writer:
                    df_sum.to_excel(writer, sheet_name=sheet_name, index=False)


if __name__ == '__main__':

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
    sum.remain_columns()
