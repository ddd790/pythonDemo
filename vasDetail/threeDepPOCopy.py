import pandas as pd
import datetime
import pymssql
import numpy as np
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
        # 日期列
        self.datetime_val = []
        # 查找所有列（三部PO订单表_明细），放入数组
        self.select_po_column_value()
        # 查询今天的数据
        self.today_data = self.select_po_value('三部PO订单表_明细')
        # 查询第二表（昨天）的数据
        self.yesterday_data = self.select_po_value('业务三部PO表_明细_昨天')
        if not self.yesterday_data.empty:
            # 根据第二表（昨天）的数据，更新第三表（前天）的数据
            self.update_db('业务三部PO表_明细_前天', self.yesterday_data)
        print('前天数据已经更新完成！' + str(datetime.datetime.now()).split('.')[0])
        # 根据今天的数据的数据，更新第二表（昨天）的
        self.update_db('业务三部PO表_明细_昨天', self.today_data)
        print('昨天数据已经更新完成！' + str(datetime.datetime.now()).split('.')[0])
        print('已经完成操作！' + str(datetime.datetime.now()).split('.')[0])

    def select_po_column_value(self):
        # 获取订单信息列的数据
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        select_sql = "select column_name,data_type from INFORMATION_SCHEMA.COLUMNS where Table_Name = '三部PO订单表_明细'"
        cursor.execute(select_sql)
        row_list = list(cursor.fetchall())
        self.add_data_title = []
        self.dic_add_po_data = {}
        for row in row_list:
            self.add_data_title.append(row[0])
            self.dic_add_po_data[row[0]] = row[1]
            if row[1] == 'datetime':
                self.datetime_val.append(row[0])
        cursor.close()
        conn.close()

    def select_po_value(self, tableName):
        # 根据表名获取订单信息的数据
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        self.strCol = ",".join(str(i) for i in self.add_data_title)
        select_sql = "select " + self.strCol + " from " + tableName
        cursor.execute(select_sql)
        row = cursor.fetchall()
        data_value = pd.DataFrame(
            data=list(row), columns=self.add_data_title)
        cursor.close()
        conn.close()
        return data_value

    def update_db(self, tableName, data_value):
        dbCol = self.add_data_title[:]
        # 建立连接并获取cursor
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        cursor.execute('TRUNCATE TABLE ' + tableName)
        # 组装插入的值
        insertValue = []
        table_value = []
        # 去除日期型的NAT数据
        for time_row in self.datetime_val:
            data_value[time_row] = pd.to_datetime(
                data_value[time_row]).dt.floor('d')
            data_value[time_row] = np.where(
                data_value[time_row].notnull(), data_value[time_row].dt.strftime('%Y-%m-%d %H:%M:%S'), None)
        table_value.append([tuple(row) for row in data_value.values])
        for tabVal in table_value:
            insertValue += tabVal
        insertSql = 'INSERT INTO ' + tableName + \
            ' (' + self.strCol + ')' + ' VALUES ('
        for index in range(len(dbCol)):
            # 判断数据类型
            if self.dic_add_po_data[dbCol[index]] == 'int' or self.dic_add_po_data[dbCol[index]] == 'decimal':
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
    VAS.get_files()


if __name__ == '__main__':
    gui_start()
