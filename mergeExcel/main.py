import os
import tkinter

import pandas as pd


# 获取文件夹下文件全路径名
def get_files(path):
    fs = []
    for root, dirs, files in os.walk(path):
        for file in files:
            fs.append(os.path.join(root, file))
    return fs

# 合并文件


def merge():
    files = get_files('F:\other\excel')
    arr = []
    for i in files:
        arr.append(pd.read_excel(i, skiprows=4))
    writer = pd.ExcelWriter('F:\other\excel\merge.xlsx')
    pd.concat(arr).to_excel(writer, 'Sheet1', index=False)
    writer.save()

# 弹出完成框


def alert():
    # 初始化Tk()
    myWindow = tkinter.Tk()
    # 设置标题
    myWindow.title('完毕！')
    # 设置窗口大小
    myWindow.geometry('380x300')
    tkinter.Label(myWindow, text="OK, Success!").pack()
    # 进入消息循环
    myWindow.mainloop()


if __name__ == '__main__':
    merge()
    alert()
