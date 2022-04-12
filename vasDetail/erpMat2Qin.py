# coding=utf8
import pandas as pd
import datetime
import pymssql
import numpy as np
import math
from tkinter import *


class VAS_GUI():
    # 批量获取服务器数据，进行累加操作
    def get_datas(self):
        print('数据操作进行中......')
        # erp服务器名
        self.serverNameErp = '192.168.0.11'
        # erp登陆用户名和密码
        self.userNameErp = 'sa'
        self.passWordErp = 'jiangbin@007'
        # erp数据库名
        self.dbNameErp = 'MDFNEW'
        # sql服务器名
        self.serverName = '192.168.0.6'
        # 登陆用户名和密码
        self.userName = 'sa'
        self.passWord = 'MS_guanli09'
        # 数据库名
        self.dbName = 'ESApp1'
        # 工厂信息列
        # 追加的dataFrame的title 加工工厂 from及条件
        self.select_sql = '''select * from AA_D_MaterialCheckInfo'''

        # 循环操作表格
        print('开始操作 【D_MaterialCheckInfo_Erp】 表，请耐心等待！' +
              str(datetime.datetime.now()).split('.')[0])
        self.select_column_value('D_MaterialCheckInfo_Erp')
        option_data = self.select_erp_value(self.select_sql)
        if not option_data.empty:
            self.update_db('D_MaterialCheckInfo_Erp', option_data)

        print('所有表操作完毕！' + str(datetime.datetime.now()).split('.')[0])

    def select_column_value(self, table_name):
        # 获取表的信息数据（列和类型）
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        select_sql = "select column_name,data_type from INFORMATION_SCHEMA.COLUMNS where Table_Name = '" + table_name + "'"
        cursor.execute(select_sql)
        row_list = list(cursor.fetchall())
        # 日期列
        self.datetime_val = []
        # title列
        self.add_data_title = []
        # 自定义字段，需要对sql中的汉字进行转码的字段
        self.change_code_col = []
        # title对应类型
        self.dic_add_col_data = {}
        for row in row_list:
            self.add_data_title.append(row[0])
            self.dic_add_col_data[row[0]] = row[1]
            if row[1] == 'datetime':
                self.datetime_val.append(row[0])
        cursor.close()
        conn.close()

    def select_erp_value(self, select_sql):
        # 根据表名和sql获取ERP信息的数据
        conn = pymssql.connect(
            self.serverNameErp, self.userNameErp, self.passWordErp, self.dbNameErp, charset='utf8')
        cursor = conn.cursor()
        cursor.execute(select_sql.encode("utf8"))
        row = cursor.fetchall()
        data_value = pd.DataFrame(
            data=list(row), columns=self.add_data_title)
        cursor.close()
        conn.close()
        return data_value

    def update_db(self, tableName, data_value):
        dbCol = self.add_data_title
        self.strCol = ",".join(str(i) for i in self.add_data_title)
        # 建立连接并获取cursor
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        cursor.execute('TRUNCATE TABLE ' + tableName)
        # 组装插入的值
        insertValue = []
        table_value = []
        # 主要转码的字段
        for change_code_row in self.change_code_col:
            if change_code_row in dbCol:
                data_value[change_code_row] = np.where(
                    data_value[change_code_row].notnull(), data_value[change_code_row], '')
                data_value[change_code_row] = data_value[change_code_row].map(
                    lambda x: x.encode('latin-1').decode('gbk'))
        # 去除日期型的NAT数据
        for time_row in self.datetime_val:
            data_value[time_row] = pd.to_datetime(
                data_value[time_row]).dt.floor('d')
            data_value[time_row] = np.where(
                data_value[time_row].notnull(), data_value[time_row].dt.strftime('%Y-%m-%d %H:%M:%S'), '')
        table_value.append([tuple(None if isinstance(i, float) and math.isnan(
            i) else i for i in t) for t in data_value.values])
        for tabVal in table_value:
            insertValue += tabVal
        # print(insertValue)
        insertSql = 'INSERT INTO ' + tableName + \
            ' (' + self.strCol + ')' + ' VALUES ('
        for index in range(len(dbCol)):
            # 判断数据类型
            if self.dic_add_col_data[dbCol[index]] == 'int' or self.dic_add_col_data[dbCol[index]] == 'decimal':
                insertSql += '%d'
            else:
                insertSql += '%s'
            # 判断是否是最后一个
            if index != len(dbCol) - 1:
                insertSql += ', '
        insertSql += ')'
        cursor.executemany(insertSql, insertValue)
        conn.commit()
        conn.close()


def gui_start():
    VAS = VAS_GUI()
    VAS.get_datas()


if __name__ == '__main__':
    gui_start()
