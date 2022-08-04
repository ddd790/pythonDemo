import os
import time
import shutil
import pandas as pd
import datetime
import pymssql
import numpy as np
from dateutil import parser
from tkinter import *


class VAS_GUI():
    # 批量获取服务器数据，进行累加操作
    def get_files(self):
        print('数据操作进行中......')
        # sql服务器名
        self.serverName = '192.168.0.6'
        # 登陆用户名和密码
        self.userName = 'sa'
        self.passWord = 'MS_guanli09'
        # 数据库名
        self.dbName = 'ESApp1'
        # trimlist文件的列对应的表
        self.add_col_data_title = ['ColName', 'ColCName']
        # 查询已存在的列的记录
        self.old_col_all_data = self.search_TrimList_File_Col()
        # 已经存在的文件的列(去重)
        self.file_col_name_old = list(set(self.old_col_all_data['ColName']))

        # 追加的dataFrame的title
        self.add_data_title = ['Version', 'PO', 'Style',
                               'Fabric', 'ColName', 'Value1', 'Value2', 'Dep', 'ColNo']
        self.local_vas_detail_file = 'd:\excelTrimListHistoryFile'

        # 修改时间有误的list
        errorTimeFileList = [
            'TRIMLIST-4900118168-V1', 'TRIMLIST-4900117745-V3']
        # 循环本地临时文件，处理合并
        self.table_value = []
        self.col_table_value = []
        self.new_col_name = []
        for lroot, ldirs, lfiles in os.walk(self.local_vas_detail_file):
            for lfile in lfiles:
                if lfile.split('.')[0][0] == '~':
                    continue
                mtime = parser.parse(time.ctime(os.path.getmtime(
                    os.path.join(lroot, lfile))))
                # ctime = time.ctime(os.path.getctime(
                #     os.path.join(lroot, lfile)))
                create_time = mtime.strftime('%Y-%m-%d %H:%M:%S')
                if lfile.split('.')[0] in errorTimeFileList:
                    create_time = '2021-07-13 10:32:54'
                if lfile.split('.')[0] == 'TRIMLIST-4900120872-V2':
                    create_time = '2021-09-23 10:32:54'
                if lfile.split('.')[0] == 'TRIMLIST-4900118169-V1':
                    create_time = '2021-08-23 10:32:54'
                self.file_to_dataframe(os.path.join(lroot, lfile), str(
                    lfile).split('-')[2].split('.')[0], create_time)
        # 更新D_TrimListFileCol数据库
        file_col_name_new = list(set(self.new_col_name))
        for new_col in file_col_name_new:
            t = (new_col, '')
            self.col_table_value.append([t])
        self.update_TrimList_File_Col()

        # 查询最新的D_TrimListFileCol数据,找到对应的key-value
        # self.now_col_data = self.search_TrimList_File_Col()
        # self.col_key_value = {}
        # for index, now_col in self.now_col_data.iterrows():
        #     self.col_key_value[now_col['ColName']] = now_col['TFCID']

        # 根据excel文件的内容进行D_TrimListFileVal的更新
        self.update_db()
        print('已经完成计算操作！')

    def file_to_dataframe(self, io, version, create_time):
        # 读取文件
        excelData = pd.read_excel(io, header=None, keep_default_na=False)
        # csv中的title
        formartExcelTitle = []
        # csv数据
        formatExcelvalue = excelData.values
        for csvIdx in range(0, len(excelData.values[0])):
            formartExcelTitle.append(csvIdx)
        df = pd.DataFrame(formatExcelvalue, columns=formartExcelTitle)
        # title
        excelTitleOld = np.array(df[0])
        excelTitle = np.append(excelTitleOld, 'ColNo')
        # 正常数据，偶数列的数据
        dataVal = []
        # 描述数据，奇数列的数据
        disVal = []
        for tempIndex in formartExcelTitle:
            str_arr = df[tempIndex].values
            for arr_i in range(len(str_arr)):
                str_arr[arr_i] = str(str_arr[arr_i]).replace(
                    '=', '').replace('"', '')
            if tempIndex != 0 and tempIndex % 2 == 0:
                new_list = np.append(str_arr, tempIndex)
                disVal.append(new_list)
            elif tempIndex % 2 != 0:
                new_list = np.append(str_arr, tempIndex)
                dataVal.append(new_list)

        valueDf = pd.DataFrame(dataVal, columns=excelTitle)
        valueDf['Version'] = version[1:]
        disDf = pd.DataFrame(disVal, columns=excelTitle)
        disDf['Purchasing Document'] = valueDf['Purchasing Document'][0]
        disDf['Version'] = version[1:]

        # 列名处理
        for col_name in excelTitle:
            if str(col_name).strip() != '' and str(col_name).strip() not in self.file_col_name_old:
                self.new_col_name.append(str(col_name).strip())

        # 对excel读取的数据进行整理，整理成符合要求的格式
        self.arrange_excel_data(valueDf, disDf, create_time)

    def arrange_excel_data(self, valueDf, disDf, create_time):
        # 根据数据结构整理dataframe
        title_list = list(valueDf)
        for idx, value in valueDf.iterrows():
            dep_val = '一部'
            for val in title_list:
                if str(val).strip() == 'Fabric':
                    dep = str(disDf.loc[idx][val]).strip()
                    if dep.__contains__('Raincoats') or dep.__contains__('Overcoats'):
                        dep_val = '五部'
                    break
            for val in title_list:
                val_t = (value['Version'], str(value['Purchasing Document']).strip(), str(value['Style']).strip(),
                         str(value['Fabric']).strip(), str(val).strip(), str(value[val]).strip(), str(disDf.loc[idx][val]).strip(), dep_val, str(value['ColNo']), create_time)
                self.table_value.append([val_t])

    # 查询D_TrimListFileCol表中已经存在的列
    def search_TrimList_File_Col(self):
        # 建立连接并获取辅料填写的数据（采购表）
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        select_sql = 'select TFCID, ColName, ColCName from D_TrimListFileCol'
        cursor.execute(select_sql)
        row = cursor.fetchall()
        old_col_all_data = pd.DataFrame(
            data=list(row), columns=['TFCID', 'ColName', 'ColCName'])
        cursor.close()
        conn.close()
        return old_col_all_data

    # 更新D_TrimListFileCol表中的列
    def update_TrimList_File_Col(self):
        dbCol = self.add_col_data_title[:]
        # 建立连接并获取cursor
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        # 组装插入的值
        insertValue = []
        for tabVal in self.col_table_value:
            insertValue += tabVal
        insertSql = 'INSERT INTO D_TrimListFileCol VALUES ('
        for colVal in dbCol:
            if colVal == 'ColCName':
                insertSql += '%s'
            else:
                insertSql += '%s, '
        insertSql += ')'
        cursor.executemany(insertSql, insertValue)
        conn.commit()
        conn.close()

    def update_db(self):
        dbCol = self.add_data_title[:]
        dbCol.append('CreateDate')
        # 建立连接并获取cursor
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        cursor.execute('TRUNCATE TABLE D_TrimListFileVal')
        # 组装插入的值
        insertValue = []
        for tabVal in self.table_value:
            insertValue += tabVal
        insertSql = 'INSERT INTO D_TrimListFileVal VALUES ('
        for colVal in dbCol:
            if colVal == 'CreateDate':
                insertSql += '%s'
            else:
                insertSql += '%s, '
        insertSql += ')'
        cursor.executemany(insertSql, insertValue)
        conn.commit()
        conn.close()


def gui_start():
    VAS = VAS_GUI()
    VAS.get_files()


if __name__ == '__main__':
    gui_start()
