import os
import pandas as pd

path = "C:/data/PL150/RFAMP Voltage Current Accuracy/1st_correction_function/summary/"

file_src = 'summary.xlsx'
file_dst = 'summary 00 39.xlsx'

df_src = pd.read_excel(path + file_src, sheet_name='summary')
df_dst = pd.read_excel(path + file_dst, sheet_name='summary')

for i in range(len(df_src)):
    if df_dst.apply(lambda x: x['filename'] == df_src['filename'][i], axis=1).any():
        idx = df_dst.index[df_dst.apply(lambda x: x['filename'] == df_src['filename'][i], axis=1)]
        df_dst.loc[idx[0]] = df_src.iloc[i]
    else:
        df_dst = df_dst.append(df_src.iloc[i: i + 1])

df_dst.sort_values(['filename'], inplace=True)
df_dst.reset_index(inplace=True, drop=True)

df_save = df_dst
filename = 'merge.xlsx'
sheet = 'summary'
if not os.path.exists(path + filename):
    with pd.ExcelWriter(path + filename, mode='w', engine='openpyxl') as writer:
        df_save.to_excel(writer, sheet_name=sheet, index=False)
else:
    with pd.ExcelWriter(path + filename, mode='a', engine='openpyxl') as writer:
        try:
            df_save.to_excel(writer, sheet_name=sheet, index=False)
        except:
            print("Shhet name '{}' is already in 'summary.xlsx'.".format(sheet))

print('=========================')