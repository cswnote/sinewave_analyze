import math
import os
import openpyxl
import platform
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

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

drop_columns = ['Board', 'PWM', 'V Frequency[MHz]', 'Volt', 'Ave. RP Coff', 'Real P[W]', 'Vmean', 'Imean[mA]', 'FFT V freq[MHz]',
       'FFT V rms', 'FFT V dc abs', 'FFT I freq[MHz]', 'FFT I rms[mA]', 'FFT I dc abs[mA]', 'Usr Status',
       'RF Volt Ch 1', 'RF Volt Ch 2', 'RF Curr Ch 1', 'RF Curr Ch 2', 'CP Pwm Ch 1', 'CP Pwm Ch 2', 'Loop Time 0.1 us']

files = os.listdir(path_summary)
files = [file for file in files if '00 07' in file and '~' not in file ]
files.sort()

sheet = (pd.ExcelFile(path_summary + files[-1]).sheet_names)[0]

df_ch3 = pd.read_excel(path_summary + files[0], sheet)
df_ch4 = pd.read_excel(path_summary + files[1], sheet)

df_all_dict = {'df_ch3': df_ch3, 'df_ch4': df_ch4}

df_all_dict['df_ch3_300'] = df_ch3.iloc[15:41, :]
df_all_dict['df_ch4_300'] = df_ch3.iloc[15:41, :]

df_all_dict['df_ch3_ch3_curr_sweep_ch4_open'] = df_ch3.iloc[54:305, :]
df_all_dict['df_ch4_ch3_curr_sweep_ch4_open'] = df_ch4.iloc[54:305, :]

df_all_dict['df_ch3_ch4_curr_sweep_ch3_open'] = df_ch3.iloc[305:556, :]
df_all_dict['df_ch4_ch4_curr_sweep_ch3_open'] = df_ch4.iloc[305:556, :]

df_all_dict['df_ch3_ch3_curr_sweep_ch4_220ma'] = df_ch3.iloc[556:807, :]
df_all_dict['df_ch4_ch3_curr_sweep_ch4_220ma'] = df_ch4.iloc[556:807, :]

df_all_dict['df_ch3_ch4_curr_sweep_ch3_220ma'] = df_ch3.iloc[807:, :]
df_all_dict['df_ch4_ch4_curr_sweep_ch3_220ma'] = df_ch4.iloc[807:, :]

del df_ch3
del df_ch4

for key, df in df_all_dict.items():
    df.reset_index(inplace=True)
    df.drop('index', axis=1, inplace=True)

for key, df in (df_all_dict.items()):
    for col in drop_columns:
        df.drop([col], axis=1, inplace=True)
    for col in list(df.columns):
        if 'deviation' in col:
            df.drop([col], axis=1, inplace=True)

for idx, key in enumerate(df_all_dict.keys()):
    if idx > 3:
        if key == 'df_ch4_ch3_curr_sweep_ch4_220ma':
            for i in range(len(df_all_dict[key])):
                df_all_dict[key].at[i, 'Curr'] = 223
        elif key == 'df_ch3_ch4_curr_sweep_ch3_220ma':
            for i in range(len(df_all_dict[key])):
                df_all_dict[key].at[i, 'Curr'] = 222
        elif key != 'df_ch4_ch3_curr_sweep_ch4_open' and key != 'df_ch3_ch4_curr_sweep_ch3_open':
            for i in range(len(df_all_dict[key])):
                df_all_dict[key].at[i, 'Curr'] = i

x = df_all_dict['df_ch3_ch3_curr_sweep_ch4_open']['Irms[mA]']
y = df_all_dict['df_ch3_ch3_curr_sweep_ch4_open']['Irms[mA]']
y2 = df_all_dict['df_ch4_ch3_curr_sweep_ch4_open']['Irms[mA]']

# # channle 3 변경과 channel 4 단락 또는 220mA 고정 값 비교
x = df_all_dict['df_ch3_ch3_curr_sweep_ch4_open']['Irms[mA]']   # ch3 sweep, ch4 open 에서 , ch3 기준점
y = df_all_dict['df_ch3_ch3_curr_sweep_ch4_open']['Irms[mA]']   # ch3 sweep, ch4 open 에서 , ch3 기준점
y2 = df_all_dict['df_ch3_ch3_curr_sweep_ch4_220ma']['Irms[mA]'] # ch3 sweep, ch4 220ma 고정에서, ch3
y3 = (df_all_dict['df_ch3_ch3_curr_sweep_ch4_open']['Vpeak[V]'] / math.sqrt(2)) / df_all_dict['df_ch3_ch3_curr_sweep_ch4_open']['Irms[mA]'] * 1000    # ch3 sweep, ch4 open 에서, ch3 impedance
y4 = (df_all_dict['df_ch3_ch3_curr_sweep_ch4_220ma']['Vpeak[V]'] / math.sqrt(2)) / df_all_dict['df_ch3_ch3_curr_sweep_ch4_220ma']['Irms[mA]'] * 1000    # ch3 sweep, ch4 220ma 고정에서, ch4 impedance
y5 = df_all_dict['df_ch3_ch3_curr_sweep_ch4_open']['CP Pwm Ch 3']   # ch3 sweep, ch4 open 에서, ch3 PWM
y6 = df_all_dict['df_ch3_ch3_curr_sweep_ch4_220ma']['CP Pwm Ch 3']  # ch3 sweep, ch4 220ma 고정에서, ch4 PWM
y7 = (df_all_dict['df_ch3_ch3_curr_sweep_ch4_open']['Vpeak[V]'] / math.sqrt(2))    # ch3 sweep, ch4 open 에서, ch3 Vrms
y8 = (df_all_dict['df_ch3_ch3_curr_sweep_ch4_220ma']['Vpeak[V]'] / math.sqrt(2))    # ch3 sweep, ch4 220ma 고정에서, ch4 Vrms

size = 40
plt.figure(figsize=(10, 20))
plt.scatter(x[:], y[:], s=size, c='r', label='I singel')
plt.scatter(x[:], y2[:], s=size, c='b', label='I multi')
plt.scatter(x[:], y3[:], s=size, c='g', label='impedance single')
plt.scatter(x[:], y4[:], s=size, c='y', label='impedance multi')
plt.scatter(x[:], y5[:], s=size, c='orange', label='PWM  single')
plt.scatter(x[:], y6[:], s=size, c='m', label='PWM multi')
plt.scatter(x[:], y7[:], s=size, c='r', marker='x', label='Vrms single')
plt.scatter(x[:], y8[:], s=size, c='b', marker='x', label='Vrms multi')


plt.grid()
plt.xlabel('Isingel')
plt.ylabel('Imulti')
plt.legend()
plt.show()

print("======================")