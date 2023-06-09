import os
import pandas as pd

path = 'C:/Users/winston/workspace/PyCharm/sinewave_analyze/data/tek_excel/'
df_summary = pd.read_excel(path + 'summary.xlsx', sheet_name='with kmon')
cols = ['filename', 'Board', 'ohm', 'Ch', 'PWM', 'Vpeak[V]', 'Irms[mA]', 'RF Volt Ch 1', 'RF Volt Ch 2', 'RF Volt Ch 3', 'RF Volt Ch 4', 'RF Curr Ch 1', 'RF Curr Ch 2', 'RF Curr Ch 3', 'RF Curr Ch 4']
df_summary = df_summary(cols)

path = 'C:/Users/winston/workspace/PyCharm/sinewave_analyze/data/kmon_csv/'
df_kmon = pd.read_excel(path + 'info_kmon_00.xlsx')
cols = ['Unnamed: 0', 'Usr Control', 'CP Pwm Set Ch 1', 'RF Volt Ch 1', 'RF Curr Ch 1', 'CP Pwm Ch 1']
df_kmon = df_kmon[cols]



print('asdf')