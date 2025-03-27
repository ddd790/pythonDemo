import os
import pandas as pd
import datetime
import pymssql
import shutil
from tkinter import *
import numpy as np


class VAS_GUI():
    # 批量获取服务器数据，进行累加操作
    def get_files(self):
        print('文件操作进行中......')
        # sql服务器名
        self.serverName = '192.168.0.11'
        # 登陆用户名和密码
        self.userName = 'sa'
        self.passWord = 'jiangbin@007'
        # 数据库名
        self.dbName = 'ESApp1'
        # 追加的dataFrame的title
        self.add_data_title = ['Style', 'Qty', 'Shipping', 'Size']
        # 数字类型的字段
        self.number_item = ['Qty']
        self.local_cai_detail_file = 'd:\\4DepPo'
        self.ship_date = []

        # 循环文件，处理合并
        self.table_value = []
        # 删除文件的list
        self.StyleList = []
        for lroot, ldirs, lfiles in os.walk(self.local_cai_detail_file):
            for lfile in lfiles:
                if not str(lfile).__contains__('~'):
                    print('文件名：' + str(lfile).split('.')[0])
                    # 读取第一个sheet页的船期
                    ship_df = pd.read_excel(os.path.join(lroot, lfile), sheet_name=0, header=None, nrows=20)
                    # 船期的list
                    ship_list = list(ship_df.iloc[12:20, 4].dropna())
                    # 取季节号中的年
                    season_year = str(ship_df.iloc[10, 1])[-4:]
                    df = pd.read_excel(os.path.join(lroot, lfile), sheet_name=1, nrows=1000)
                    self.format_dataframe(df, lfile, ship_list, season_year)

        # 更新数据库，删除文件
        self.update_db()
        # 删除目录内文件
        self.delete_files_in_folder(self.local_cai_detail_file)
        input('按回车退出 ')

    def format_dataframe(self, df, lfile, ship_list, season_year):
        df.drop([len(df)-1], inplace=True)
        df.dropna(axis=1, how='all', inplace=True)
        # 第一列为size列
        size_list = list(map(str,list(df.iloc[:, 0].dropna())))
        del size_list[0]
        # 创建一个空列表来存储需要删除的列名
        columns_to_drop = []
        # 遍历每一列
        for col in df.columns:
            # 检查该列是否有任何单元格包含 "SIZE"
            if df[col].str.contains('SIZE').any():
                columns_to_drop.append(col)
        # 删除这些列
        df.drop(columns=columns_to_drop, inplace=True)
        # 创建一个空列表来存储需要删除的行索引
        rows_to_drop = []
        # 遍历每一行
        for index, row in df.iterrows():
            # 检查该行是否有任何单元格不是数字也不是NaN
            if not row.apply(lambda x: np.isreal(x) and not np.isnan(x)).all():
                rows_to_drop.append(index)
        # 删除这些行
        df = df.drop(rows_to_drop)
        df = df.reset_index(drop=True)
        # 新建dataframe
        table_data_add = pd.DataFrame(data=None, columns=self.add_data_title)
        idx = 0
        # 遍历每一列
        for col in df.columns:
            # 新建dataframe
            table_data = pd.DataFrame(data=None, columns=self.add_data_title)
            table_data['Qty'] = df[col].astype(int)
            table_data['Size'] = pd.DataFrame(size_list)
            table_data['Shipping'] = self.change_shipping_date(ship_list[idx], season_year)
            table_data['Style'] = lfile.split(' ')[0]
            idx = idx + 1
            table_data_add = table_data_add.append(table_data, ignore_index=True)
        table_data_add['FileName'] = str(lfile).split('.')[0]
        table_data_add['CreateDate'] = str(datetime.datetime.now()).split('.')[0]
        self.table_value.append([tuple(row) for row in table_data_add.values])
        self.StyleList.extend(table_data_add['Style'].tolist())

    def update_db(self):
        dbCol = self.add_data_title[:]
        dbCol.append('FileName')
        dbCol.append('CreateDate')
        # 建立连接并获取cursor
        conn = pymssql.connect(self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        # 组装删除的值
        del_tuple = tuple(self.StyleList)
        # 删除已经存在的文件
        delSql = 'delete from D_4DepCai_S where Style = (%s)'
        cursor.executemany(delSql, del_tuple)
        # 组装插入的值
        insertValue = []
        for tabVal in self.table_value:
            insertValue += tabVal
        insertSql = 'INSERT INTO D_4DepCai_S VALUES ('
        for colVal in dbCol:
            if colVal == 'CreateDate':
                insertSql += '%s'
            elif colVal in self.number_item:
                insertSql += '%d, '
            else:
                insertSql += '%s, '
        insertSql += ')'
        cursor.executemany(insertSql, insertValue)
        conn.commit()
        conn.close()

    def change_shipping_date(self, ship_date, year):
        date_list = ship_date.split(' ')
        key_list = {
            'JAN': '01',
            'FEB': '02',
            'MAR': '03',
            'MARCH': '03',
            'APR': '04',
            'APRIL': '04',
            'MAY': '05',
            'JUNE': '06',
            'JULY': '07',
            'AUG': '08',
            'SEPT': '09',
            'OCT': '10',
            'NOV': '11',
            'DEC': '12',
        }
        month = key_list[date_list[0]]
        # 如果date_list长度为1，则day为'01'
        day = '01'
        if len(date_list) > 1:
            day = date_list[1][:-2]
        if len(day) == 1:
            day = '0' + day
        retrun_date = year + '-' + month + '-' + day + ' 00:00:00'
        # 如果retrun_date小于当前日期，则加一年
        if datetime.datetime.now() > datetime.datetime.strptime(retrun_date, '%Y-%m-%d %H:%M:%S'):
            year = str(int(year) + 1)
            retrun_date = year + '-' + month + '-' + day + ' 00:00:00'
        return retrun_date

    def delete_files_in_folder(self, folder_path):
        # 确保提供的路径是一个有效的文件夹路径
        if not os.path.isdir(folder_path):
            print(f"Path {folder_path} is not a valid directory.")
            return

        # 获取文件夹中的所有文件
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            
            # 检查是否为文件
            if os.path.isfile(file_path):
                try:
                    # 删除文件
                    os.remove(file_path)
                    print(f"Deleted file: {file_path}")
                except Exception as e:
                    print(f"Failed to delete {file_path}. Error: {e}")

def gui_start():
    VAS = VAS_GUI()
    VAS.get_files()


if __name__ == '__main__':
    gui_start()
