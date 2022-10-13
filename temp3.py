import os
import openpyxl

path = 'D:/winston/Desktop/새 폴더/test79/'

files = os.listdir(path)
files = [file for file in files if file[:3] == 'tek']
files.sort()

for file in files:
    src = os.path.join(path + file)
    num = int(file.split('.')[0][3:])

    num -= 60
    name = 'tek' + str(num) + '.' + file.split('.')[1]

    dst = os.path.join(path + name)
    os.rename(src, dst)


