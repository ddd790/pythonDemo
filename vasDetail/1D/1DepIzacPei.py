import pandas as pd
import datetime
import pymssql
from tkinter import *


class VAS_GUI():
    # 一部IZAC配码拆分尺码
    def get_files(self):
        print('操作进行中......')
        # sql服务器名
        self.serverName = '192.168.0.11'
        # 登陆用户名和密码
        self.userName = 'sa'
        self.passWord = 'jiangbin@007'
        # 数据库名
        self.dbName = 'ESApp1'
        # 追加的dataFrame的title
        self.add_data_title = ['用料名称', '尺码分档要求', '规格尺寸', '型号', '款式', '季节号', 'size', 'delKey']
        # 定义尺码映射表
        self.size_map = {'XXS': 22, 'XS': 24, 'S': 26, 'M': 28, 'L': 30, 'XL': 32, 'XXL': 34, 'XXXL': 36, 'XXXXL': 38, 'XXXXXL': 40}
        # 反向映射表
        self.reverse_size_map = {v: k for k, v in self.size_map.items()}

        # 循环文件，处理合并
        self.table_value = []
        # # 删除文件的list
        self.delList = []
        # 获取数据库内容
        self.get_size_info()
        # 添加新的列
        self.pei_data[['start', 'end']] = self.pei_data.apply(self.split_intervals, axis=1)
        self.pei_data['size'] = self.pei_data.apply(self.generate_numbers, axis=1)
        self.pei_data = self.pei_data.explode('size', ignore_index=True)
        self.pei_data.drop(['start', 'end'], axis=1, inplace=True)
        self.pei_data['size'] = self.pei_data['size'].astype(str)
        # 拼接出删除的列
        self.pei_data['delKey'] = self.pei_data['款式'] + '-' + self.pei_data['季节号']
        self.pei_data['CreateDate'] = str(datetime.datetime.now()).split('.')[0]
        self.delList.extend(self.pei_data['delKey'].tolist())
        self.table_value.append([tuple(row) for row in self.pei_data.values])
        # 更新数据库
        self.update_db()
        input('按回车退出 ')
    # 拆分区间字符串
    def split_intervals(self, row):
        start, end = row['尺码分档要求'].split('-')
        return pd.Series([start, end])
    # 生成区间内的所有数字, 步长为2
    def generate_numbers(self, row):
        if self.is_number(row['start']):
            return pd.Series(range(int(row['start']), int(row['end']) + 1, 2)).tolist()
        else:
            start_index = self.size_map[row['start']]
            end_index = self.size_map[row['end']]
            sizes = [self.reverse_size_map[i] for i in range(start_index, end_index + 1, 2)]
            return pd.Series(sizes).tolist()
    # 查询配码表中已经存在的列
    def get_size_info(self):
        # 建立连接并获取配码的数据
        conn = pymssql.connect(self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        select_sql = '''SELECT a.用料名称, a.尺码分档要求, a.规格尺寸, a.型号, a.款式, a.季节号 
                        from 一部IZAC配码_明细 a, 一部IZAC配码_主表 b
                        WHERE a.ExcelServerRCID = b.ExcelServerRCID
                        AND b."修改时间" >= DATEADD(MONTH, -1, GETDATE())'''
        cursor.execute(select_sql)
        row = cursor.fetchall()
        # 配码数据
        self.pei_data = pd.DataFrame(data=list(row), columns=['用料名称', '尺码分档要求', '规格尺寸', '型号', '款式', '季节号'])
        cursor.close()
        conn.close()

    def update_db(self):
        dbCol = self.add_data_title[:]
        dbCol.append('CreateDate')
        # 建立连接并获取cursor
        conn = pymssql.connect(self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        # 组装删除的值
        del_tuple = tuple(self.delList)
        # 删除已经存在的文件
        delSql = 'delete from D_1DepIzacPei where delKey = (%s)'
        cursor.executemany(delSql, del_tuple)
        # 组装插入的值
        insertValue = []
        for tabVal in self.table_value:
            insertValue += tabVal
        insertSql = 'INSERT INTO D_1DepIzacPei VALUES ('
        for colVal in dbCol:
            if colVal == 'CreateDate':
                insertSql += '%s'
            else:
                insertSql += '%s, '
        insertSql += ')'
        cursor.executemany(insertSql, insertValue)
        conn.commit()
        conn.close()
    def is_number(self, s):
        try:
            float(s)
            return True
        except ValueError:
            pass

        try:
            import unicodedata
            unicodedata.numeric(s)
            return True
        except (TypeError, ValueError):
            pass

        return False
    
def gui_start():
    VAS = VAS_GUI()
    VAS.get_files()


if __name__ == '__main__':
    gui_start()
