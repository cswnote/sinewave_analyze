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

    def edit_columns(self):
        df_ctrl = pd.read_excel(self.path + self.eval_file, sheet_name='edit columns')
        df_sum = pd.read_excel(self.tek_excel_path + 'summary.xlsx', sheet_name='with kmon')

        remain_items = list(df_ctrl.iloc[:, 0])

        for item in remain_items:
            if 'curr' in item.lower():
                remain_items.append('irms')

        drop_list = list(df_sum.columns)[7:]

        for item in remain_items:
            for i in range(len(drop_list) - 1, -1 , -1):
                if item in drop_list[i].lower().replace(' ', ''):
                    drop_list.remove(drop_list[i])

        print('==============')

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
    sum.edit_columns()
