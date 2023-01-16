import pandas as pd
import numpy as np
import sys
import os
import platform
import gc
from collections import defaultdict

if platform.platform() == 'macOS-13.1-arm64-arm-64bit':
    path = '/Volumes/Winston 2T/work/fieldcure/multi channel test data/'
elif platform.platform() == 'Windows-10-10.0.19044-SP0':
    path = ''

path_summary = path + 'summary/'

'''  data dictionary 구조
     data['phantom type']['test code']['측정 channel']['측정조건'] = df
     data['R']['channel 저항']['interference 저항']['측정 channel']['측정조건'] = df
'''
data_phantom = {'2ch_gel': {'230106_1': {'08ch3': {}, '08ch4':{}}}}
data_R_model = {'4ch_R150': {'R150': {'08ch3': {}, '08ch4': {}}}}

files = os.listdir(path_summary)
files = [file for file in files if '08 12 08' in file and '._' not in file]

df_all_ch3 = pd.read_excel(path_summary + files[0], sheet_name='with kmon correction')
df_all_ch4 = pd.read_excel(path_summary + files[1], sheet_name='with kmon correction')

start = 0
stop = 251
test_item = 'ch3_curr_sweep_ch4_220ma'
print("df_ch3 {} ~ {}".format(df_all_ch3.at[start, 'filename'], df_all_ch3.at[stop - 1, 'filename']))
print("df_ch4 {} ~ {}".format(df_all_ch4.at[start, 'filename'], df_all_ch4.at[stop - 1, 'filename']))
print(test_item)
df_ch3 = df_all_ch3.iloc[start:stop, :]
df_ch4 = df_all_ch4.iloc[start:stop, :]
data_phantom['2ch_gel']['230106_1']['08ch3'][test_item] = df_ch3
data_phantom['2ch_gel']['230106_1']['08ch4'][test_item] = df_ch4

start = stop
stop = stop + 251
test_item = 'ch4_curr_sweep_ch3_220ma'
print("df_ch3 {} ~ {}".format(df_all_ch3.at[start, 'filename'], df_all_ch3.at[stop - 1, 'filename']))
print("df_ch4 {} ~ {}".format(df_all_ch4.at[start, 'filename'], df_all_ch4.at[stop - 1, 'filename']))
print(test_item)
df_ch3 = df_all_ch3.iloc[start:stop, :]
df_ch4 = df_all_ch4.iloc[start:stop, :]
data_phantom['2ch_gel']['230106_1']['08ch3'][test_item] = df_ch3
data_phantom['2ch_gel']['230106_1']['08ch4'][test_item] = df_ch4

start = stop
stop = stop + 75
test_item = 'ch3_volt_sweep_ch4_50vp'
print("df_ch3 {} ~ {}".format(df_all_ch3.at[start, 'filename'], df_all_ch3.at[stop - 1, 'filename']))
print("df_ch4 {} ~ {}".format(df_all_ch4.at[start, 'filename'], df_all_ch4.at[stop - 1, 'filename']))
print(test_item)
df_ch3 = df_all_ch3.iloc[start:stop, :]
df_ch4 = df_all_ch4.iloc[start:stop, :]
data_phantom['2ch_gel'] = {'230109_1': {}}
data_phantom['2ch_gel']['230109_1'] = {'08ch3': {}, '08ch4': {}}
data_phantom['2ch_gel']['230109_1']['08ch3'][test_item] = df_ch3
data_phantom['2ch_gel']['230109_1']['08ch4'][test_item] = df_ch4

start = stop
stop = stop + 75
test_item = 'ch4_volt_sweep_ch3_50vp'
print("df_ch3 {} ~ {}".format(df_all_ch3.at[start, 'filename'], df_all_ch3.at[stop - 1, 'filename']))
print("df_ch4 {} ~ {}".format(df_all_ch4.at[start, 'filename'], df_all_ch4.at[stop - 1, 'filename']))
print(test_item)
df_ch3 = df_all_ch3.iloc[start:stop, :]
df_ch4 = df_all_ch4.iloc[start:stop, :]
data_phantom['2ch_gel']['230109_1']['08ch3'][test_item] = df_ch3
data_phantom['2ch_gel']['230109_1']['08ch4'][test_item] = df_ch4

start = 652
stop = 903
test_item = 'ch3_curr_sweep_all_220ma'
print("df_ch3 {} ~ {}".format(df_all_ch3.at[start, 'filename'], df_all_ch3.at[stop - 1, 'filename']))
print("df_ch4 {} ~ {}".format(df_all_ch4.at[start, 'filename'], df_all_ch4.at[stop - 1, 'filename']))
print(test_item)
df_ch3 = df_all_ch3.iloc[start:stop, :]
df_ch4 = df_all_ch4.iloc[start:stop, :]
data_R_model['4ch_R150']['R150']['08ch3'][test_item] = df_ch3
data_R_model['4ch_R150']['R150']['08ch4'][test_item] = df_ch4

start = 903
stop = 1154
test_item = 'ch4_curr_sweep_all_220ma'
print("df_ch3 {} ~ {}".format(df_all_ch3.at[start, 'filename'], df_all_ch3.at[stop - 1, 'filename']))
print("df_ch4 {} ~ {}".format(df_all_ch4.at[start, 'filename'], df_all_ch4.at[stop - 1, 'filename']))
print(test_item)
df_ch3 = df_all_ch3.iloc[start:stop, :]
df_ch4 = df_all_ch4.iloc[start:stop, :]
data_R_model['4ch_R150']['R150']['08ch3'][test_item] = df_ch3#_ch4_volt_sweep_ch3_50vp
data_R_model['4ch_R150']['R150']['08ch4'][test_item] = df_ch4#_ch4_volt_sweep_ch3_50vp

start = 1154
stop = start + 51
test_item = 'ch3_volt_sweep_all_25vp'
print("df_ch3 {} ~ {}".format(df_all_ch3.at[start, 'filename'], df_all_ch3.at[stop - 1, 'filename']))
print("df_ch4 {} ~ {}".format(df_all_ch4.at[start, 'filename'], df_all_ch4.at[stop - 1, 'filename']))
print(test_item)
df_ch3 = df_all_ch3.iloc[start:stop, :]
df_ch4 = df_all_ch4.iloc[start:stop, :]
data_R_model['4ch_R150']['R150']['08ch3'][test_item] = df_ch3
data_R_model['4ch_R150']['R150']['08ch4'][test_item] = df_ch4

start = stop
stop = stop + 51
test_item = 'ch4_volt_sweep_all_25vp'
print("df_ch3 {} ~ {}".format(df_all_ch3.at[start, 'filename'], df_all_ch3.at[stop - 1, 'filename']))
print("df_ch4 {} ~ {}".format(df_all_ch4.at[start, 'filename'], df_all_ch4.at[stop - 1, 'filename']))
print(test_item)
df_ch3 = df_all_ch3.iloc[start:stop, :]
df_ch4 = df_all_ch4.iloc[start:stop, :]
data_R_model['4ch_R150']['R150']['08ch3'][test_item] = df_ch3
data_R_model['4ch_R150']['R150']['08ch4'][test_item] = df_ch4

start = stop
stop = stop + 251
test_item = 'ch3_curr_sweep_ch4_220ma'
print("df_ch3 {} ~ {}".format(df_all_ch3.at[start, 'filename'], df_all_ch3.at[stop - 1, 'filename']))
print("df_ch4 {} ~ {}".format(df_all_ch4.at[start, 'filename'], df_all_ch4.at[stop - 1, 'filename']))
print(test_item)
df_ch3 = df_all_ch3.iloc[start:stop, :]
df_ch4 = df_all_ch4.iloc[start:stop, :]
data_R_model['2ch_R150'] = {'R150': {'08ch3': {}, '08ch4': {}}}
data_R_model['2ch_R150']['R150']['08ch3'][test_item] = df_ch3
data_R_model['2ch_R150']['R150']['08ch4'][test_item] = df_ch4

start = stop
stop = stop + 251
test_item = 'ch4_curr_sweep_ch3_220ma'
print("df_ch3 {} ~ {}".format(df_all_ch3.at[start, 'filename'], df_all_ch3.at[stop - 1, 'filename']))
print("df_ch4 {} ~ {}".format(df_all_ch4.at[start, 'filename'], df_all_ch4.at[stop - 1, 'filename']))
print(test_item)
df_ch3 = df_all_ch3.iloc[start:stop, :]
df_ch4 = df_all_ch4.iloc[start:stop, :]
data_R_model['2ch_R150']['R150']['08ch3'][test_item] = df_ch3
data_R_model['2ch_R150']['R150']['08ch4'][test_item] = df_ch4

start = stop
stop = stop + 50
test_item = 'ch3_volt_sweep_ch4_25vp'
print("df_ch3 {} ~ {}".format(df_all_ch3.at[start, 'filename'], df_all_ch3.at[stop - 1, 'filename']))
print("df_ch4 {} ~ {}".format(df_all_ch4.at[start, 'filename'], df_all_ch4.at[stop - 1, 'filename']))
print(test_item)
df_ch3 = df_all_ch3.iloc[start:stop, :]
df_ch4 = df_all_ch4.iloc[start:stop, :]
data_R_model['2ch_R150']['R150']['08ch3'][test_item] = df_ch3
data_R_model['2ch_R150']['R150']['08ch4'][test_item] = df_ch4

start = stop
stop = stop + 50
test_item = 'ch4_volt_sweep_ch3_25vp'
print("df_ch3 {} ~ {}".format(df_all_ch3.at[start, 'filename'], df_all_ch3.at[stop - 1, 'filename']))
print("df_ch4 {} ~ {}".format(df_all_ch4.at[start, 'filename'], df_all_ch4.at[stop - 1, 'filename']))
print(test_item)
df_ch3 = df_all_ch3.iloc[start:stop, :]
df_ch4 = df_all_ch4.iloc[start:stop, :]
data_R_model['2ch_R150']['R150']['08ch3'][test_item] = df_ch3
data_R_model['2ch_R150']['R150']['08ch4'][test_item] = df_ch4



print('==================')