import os
import pandas as pd
import FILE_MANAGEMENT
import TEK_CSV
import GET_SUMMARY
import platform
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.ticker import ScalarFormatter

def apply_fft(sampling_time, y_data):
    n = len(y_data)
    k = np.arange(n)
    T = n * sampling_time
    freq = k / T
    freq = freq[range(int(n / 2))]
    y_fft = np.fft.fft(y_data) / n
    y_fft = y_fft[range(int(n / 2))] * 2 / np.sqrt(2)

    return freq, y_fft


path = 'C:/Users/winston/Documents/PL150/EMI scope'
files = os.listdir(path)
files = [file for file in files if file[-3:]=='csv']
data = {}


for file in files:
    key = int(file.split('.')[0][3:])
    data[key] = {}
    data[key]['df'] = pd.read_csv(os.path.join(path, file), skiprows=lambda x: x<20)
    temp = pd.read_csv(os.path.join(path, file), skiprows=lambda x: (x > 15 or x < 2), index_col=0)
    cols = temp.columns.to_list()
    temp = temp.rename(columns={cols[0]: 1, cols[1]: 2, cols[2]: 3, cols[3]: 4})
    data[key]['scope_info'] = temp


for file in data.keys():
    cols = data[file]['df'].columns.to_list()
    cols.remove('TIME')
    data[file]['fft'] = {}
    for col in cols:
        data[file]['fft'][col] = pd.DataFrame(columns = ['freq', 'complex', 'abs'])
        sampling_period = float(data[file]['scope_info'].at['Sample Interval', 1])
        data[file]['fft'][col]['freq'], data[file]['fft'][col]['complex'] = apply_fft(sampling_period, data[file]['df'][col])
        data[file]['fft'][col]['abs'] = abs(data[file]['fft'][col]['complex'])

file = 3
fft = 'fft'
ch = 'CH' + '1'

x = data[file][fft][ch]['freq']
y = 20 * np.log10(data[file][fft][ch]['abs'])

fig = plt.figure(figsize=(20, 10))

start = x.index[x.apply(lambda x: x==150 * 10**3)].values[0]
stop = x.index[x.apply(lambda x: x==30 * 10**6)].values[0] - 1

marker_size = 5
font_size = 15

plt.scatter(x[start:stop], y[start:stop])
plt.xscale('log')
plt.show()

print('=====================')