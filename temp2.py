import math
import os
import openpyxl
import platform
import numpy as np
import pandas as pd

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

path_ch3 = path_excel + 'Ch3'
path_ch4 = path_excel + 'Ch4'

files = os.listdir(path_summary)
files = [file for file in files if '~' not in file and 'summary' in file]
files

sheets = pd.ExcelFile(path_summary + files[1]).sheet_names

df = pd.DataFrame()

df_ch4 = pd.read_excel(path_summary + files[1], sheet_name=sheets[-1])
df_ch3 = pd.read_excel(path_summary + files[0], sheet_name=sheets[-1])

df['filename'] = df_ch3['filename']
df['Vp_ch3'] = df_ch3['Vpeak[V]']
df['Vp_ch4'] = df_ch4['Vpeak[V]']
df['Irms_ch3'] = df_ch3['Irms[mA]']
df['Irms_ch4'] = df_ch4['Irms[mA]']
df['angle_ch3'] = df_ch3['Delay(degree)']
df['angle_ch4'] = df_ch4['Delay(degree)']
df['Vmean_ch3'] = df_ch3['Vmean']
df['Vmean_ch4'] = df_ch4['Vmean']
df['FFT_Vrms_ch3'] = df_ch3['FFT V rms']
df['FFT_Vrms_ch4'] = df_ch4['FFT V rms']
df['FFT_Vdc_ch3'] = df_ch3['FFT V dc abs']
df['FFT_Vdc_ch4'] = df_ch4['FFT V dc abs']
df['FFT_Irms_ch3'] = df_ch3['FFT I rms[mA]']
df['FFT_Irms_ch4'] = df_ch4['FFT I rms[mA]']
df['FFT_Idc_ch3'] = df_ch3['FFT I dc abs[mA]']
df['FFT_Idc_ch4'] = df_ch4['FFT I dc abs[mA]']

target_list = ['RF Volt Ch 3', 'RF Volt Ch 4', 'RF Curr Ch 3', 'RF Curr Ch 4', 'CP Pwm Ch 3', 'CP Pwm Ch 4']

for target in target_list:
    col = target.replace(' ', '')
    if 'Volt' in target:
        col = 'Vmcu_' + col[-3:].lower()
    elif 'Curr' in target:
        col = 'Imcu_' + col[-3:].lower()
    elif 'Pwm' in target:
        col = 'PWM_' + col[-3:].lower()

    if target[-1] == '3':
        df[col] = df_ch3[target]
        print("{}가 df에 {}이름으로 기록되었습니다.".format(target, col))
    elif target[-1] == '4':
        df[col] = df_ch4[target]
        print("{}가 df에 {}이름으로 기록되었습니다.".format(target, col))

print('====================')
filename = 'summary.xlsx'
sheet = '08Ch3 Ch4 merge'
if not os.path.exists(path_summary + filename):
    with pd.ExcelWriter(path_summary + filename, mode='w', engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=sheet, index=False)
else:
    with pd.ExcelWriter(path_summary + filename, mode='a', engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=sheet, index=False)