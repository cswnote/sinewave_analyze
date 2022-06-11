# import os
# import openpyxl
#
# path = 'D:/winston/OneDrive - (주)필드큐어/data_analyze/'
# path_csv = path + 'tek_csv/'
# path_excel = path + 'tek_excel/'
# path_summary = path + 'summary/'
# path_information = path + 'test information/'
# path_kmon = path + 'kmon_csv/'
#
# files = os.listdir(path_kmon)
# # files = [file for file in files if file[:3] == 'tek']
#
# for file in files:
#     scr = path_kmon + file
#     file = file.split('_')[0] + '_kmon_' + file[10:]
#     dst = path_kmon + file
#     os.rename(scr, dst)


a = ['Volt', 'Curr']

for item in a:
    print(item)

a = [item.lower() for item in a if True]
print(a)