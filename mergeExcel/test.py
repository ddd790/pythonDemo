import os

import pandas as pd


# 获取文件夹下文件全路径名
def get_files():
    sExcelFile = "F:/excel服务器/其他/工作簿1(1).xlsx"
    df = pd.read_excel(sExcelFile, sheet_name='历史记录')
    # 获取最大行，最大列
    nrows = df.shape[0]
    ncols = df.columns.size

    print("=========================================================================")
    print('Max Rows:'+str(nrows))
    print('Max Columns:'+str(ncols))


if __name__ == '__main__':
    get_files()
