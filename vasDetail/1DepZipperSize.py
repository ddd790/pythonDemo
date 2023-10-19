import os
import pandas as pd
import datetime
import pymssql
from tkinter import *
import re
import time
import copy
from dateutil import parser


class VAS_GUI():
    # 1部拉链配码excel读取
    def commit_batch(self):
        print('数据操作进行中......')
        # sql服务器名
        self.serverName = '192.168.0.11'
        # 登陆用户名和密码
        self.userName = 'sa'
        self.passWord = 'jiangbin@007'

        # 数据库的列
        self.dbCol = ['PO', 'ITEM', 'StockCategory', 'Size', 'ScheduledQty', 'Style', 'Fabric', 'FileName', 'CreateDate']

        # 循环文件，处理合并，并存入数据库
        # self.local_vas_detail_file = r'\\192.168.0.3\01-业务一部资料\Ada\VAS-PPR'
        self.local_vas_detail_file = r'D:\zipperSize'
        self.df_data = pd.DataFrame(columns=self.dbCol)

        for lroot, ldirs, lfiles in os.walk(self.local_vas_detail_file):
            for file in lfiles:
                if str(file).__contains__('SIZESCALES') and (str(file).__contains__('.xls') or str(file).__contains__('.xlsx')) and not str(file).__contains__('~'):
                    self.arrange_excel_data(os.path.join(lroot, file))
        # 追加数据(1是删除所有数据，2是删除当天数据，3是不删除直接追加)
        self.batch_update_db(self.df_data, 1)
        print('已经完成数据操作！')
        input('按回车退出 ')

    def batch_update_db(self, temp_data, deleteFlag):
        self.table_value = []
        # 将新的数据追加到旧数据中
        self.table_value.append([tuple(row) for row in temp_data.values])
        # 更新数据库
        self.update_db(deleteFlag)

    def arrange_excel_data(self, io):
        # 文件名
        self.fileNameVersion = str(io).lstrip().split('\\')[2].split('.')[0]
        self.dataItem = ['A', 'PO', 'ITEM', 'Material', 'Stock Category', 'Size', 'Scheduledqty']
        # 不满足格式条件的excel，需要转成csv，然后转成DataFrame
        new_data = self.file_to_dataframe(io)
        # excelData = pd.read_excel(io, header=0, keep_default_na=False)
        # new_data = ''
        # new_data = pd.DataFrame(excelData.values, columns=self.dataItem)
        # formartTitle = list(new_data)
        # print(formartTitle)

        # 删除Material为*和空的列
        material_all = list(new_data['Material'].unique())
        material_all.remove('*')
        material_all.remove('')

        # 新的dataFrame的数据
        new_df = new_data.filter(self.dataItem[1:]).where(new_data['Material'].isin(material_all)).dropna()
        # 尺码取前两位
        new_df['Size'] = new_df['Size'].str.slice(0, 2)
        new_df[['款式缩写', '面料']] = new_df.Material.str.split('-', expand=True)
        new_df.drop(axis = 1, columns = 'Material', inplace = True)
        new_df['fileName'] = self.fileNameVersion
        new_df['CreateDate'] = str(datetime.datetime.now()).split('.')[0]
        new_df.columns = self.dbCol
        self.df_data = self.df_data.append(new_df, ignore_index=True)
        # 去掉前面的0
        self.df_data['ITEM'] = self.df_data['ITEM'].astype('int')
        self.df_data['ITEM'] = self.df_data['ITEM'].astype('str')

    def file_to_dataframe(self, io):
        file_name = self.excel_csv_change(io, 1)
        formatCsvData = self.csv_to_dataframe(file_name)
        change_file_name = self.excel_csv_change(file_name, 2)
        return formatCsvData

    def excel_csv_change(self, io, flag):
        # 原文件后缀名
        suffix_name = '.xls' if flag == 1 else '.csv'
        # 新文件后缀名
        new_suffix_name = '.csv' if flag == 1 else '.xls'
        # flag = 1为excel2csv, flag = 2为csv2excel
        index = io.find(suffix_name)
        new_file_name = io[:index]+new_suffix_name
        os.replace(io, new_file_name)
        return new_file_name

    def csv_to_dataframe(self, io):
        csv = ''
        for decode in ('gbk', 'utf-8', 'gb18030'):
            try:
                csv = pd.read_csv(io, encoding=decode, skip_blank_lines=False, delimiter="@", header=None)
                break
            except:
                pass
        # csv数据, 并去掉空格
        formatCsv = []
        split_flag = False
        for csvIdx in range(0, len(csv)):
            csvValues = ['' for v in range(7)]
            # 截取【PO Summary by FERT/CATEGORY/SIZE】和【PO Summary by FERT/EX-FACTORY/SIZE】之间的字符串
            csvTxt = str(csv.iloc[csvIdx].values[0])
            if csvTxt.__contains__('PO Summary by FERT/CATEGORY/SIZE'):
                split_flag = True
                continue
            if split_flag:
                tempCsvVal = csvTxt.replace('\t', '@').split('@')
                for tempIdx in range(0, 7):
                    if tempIdx > len(tempCsvVal) - 1:
                        csvValues[tempIdx] = ''
                    else:
                        csvValues[tempIdx] = str(tempCsvVal[tempIdx]).strip()
                formatCsv.append(csvValues)
            if csvTxt.__contains__('PO Summary by FERT/EX-FACTORY/SIZE'):
                split_flag = False
                break
        # 去掉前两个
        formatCsv = formatCsv[2:]
        df = pd.DataFrame(formatCsv, columns=self.dataItem)
        return df

    def update_db(self, deleteFlag):
        # 建立连接并获取cursor
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, "ESApp1")
        cursor = conn.cursor()
        # deleteFlag == 1，删除原数据
        if deleteFlag == 1:
            cursor.execute('TRUNCATE TABLE D_1DepZipperSize')
        # 组装插入的值
        insertValue = []
        for tabVal in self.table_value:
            insertValue += tabVal
        insertSql = 'INSERT INTO D_1DepZipperSize VALUES ('
        for colVal in self.dbCol:
            if colVal == 'CreateDate':
                insertSql += '%s'
            elif colVal == 'Size' or colVal == 'Scheduledqty':
                insertSql += '%d, '
            else:
                insertSql += '%s, '
        insertSql += ')'
        cursor.executemany(insertSql, insertValue)
        conn.commit()
        conn.close()

    def select_all_data(self):
        # 建立连接并获取cursor
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, "ESApp1")
        cursor = conn.cursor()
        strCol = ",".join(str(i) for i in self.dbCol)
        select_sql = 'select ' + strCol + ' from D_1DepZipperSize'
        cursor.execute(select_sql)
        row = cursor.fetchall()
        self.old_all_data = pd.DataFrame(data=list(row), columns=self.dbCol)
        cursor.close()
        conn.close()


def gui_start():
    VAS = VAS_GUI()
    VAS.commit_batch()


if __name__ == '__main__':
    gui_start()
