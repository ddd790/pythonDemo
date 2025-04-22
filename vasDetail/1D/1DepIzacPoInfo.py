import pandas as pd
import datetime
import pymssql
from tkinter import *


class VAS_GUI():
    # 1部izacPO共享数据读取
    def commit_batch(self):
        print('数据操作进行中......')
        # sql服务器名
        self.serverName = '192.168.0.11'
        # 登陆用户名和密码
        self.userName = 'sa'
        self.passWord = 'jiangbin@007'
        # 数据库名
        self.dbName = 'ESApp1'
        # 追加的dataFrame的title
        self.add_data_title = ['来单日期', '订单备注', '客户', '季节号', 'PO号', '款式名称', '英文款名', '中文款式', 'classification', '线上颜色', '面料颜色', 'QTY', '价格', 
                               '总金额', '交期', '工厂', '是否有工艺', '是否录BOM', '生产部门', '客人到料时间', '采购到料时间', '给生产到料时间', '供应商', '面料号', '客人PO成分', '供应商成分', '克重', 
                               '幅宽', '面料价格', '实际采购面料价格', '面料特点', '采购单耗', '大货单耗', '加减裁数量', '款式图', '开发款式', '序号', '工艺特点', '样品轮数', '样品数量', 
                               '样品状态', '走货发票号', '挂法', '贸易方式', 'HS编码', '加工费含税']
        # 数字类型的字段
        self.number_item = ['QTY', '样品数量', '加减裁数量', '价格', '总金额', '面料价格', '实际采购面料价格', '采购单耗', '大货单耗', '加工费含税']
        # 日期类型的字段
        self.date_item = ['来单日期', '客人到料时间', '采购到料时间', '给生产到料时间', '交期']
        # 循环文件，处理合并，并存入数据库
        self.local_vas_detail_file = r'\\192.168.0.3\01-业务一部资料\A  IZAC\5.H25\H25 IZAC PO.xlsx'
        # self.local_vas_detail_file = r'D:\temp\H25 IZAC PO.xlsx'
        self.table_value = []
        # 删除文件的list
        self.keyList = []
        # 读取A到AT列的内容
        df = pd.read_excel(self.local_vas_detail_file, sheet_name=0, skiprows=1, usecols='A:BD', dtype=str)
        table_data = pd.DataFrame(df)
        add_data = pd.DataFrame(data=None, columns=self.add_data_title)
        # title对应的excel列
        col_idx = [0, 23, 1, 2, 3, 4, 5, 6, 7, 8, 8, 9, 10, 11, 12, 14, 15, 16, 17, 49, 19, 18, 21, 22, 13, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 47, 35, 36, 37, 38, 39, 40, 41, 42, 43]
        for i in range(len(self.add_data_title)):
            add_data[self.add_data_title[i]] = table_data.iloc[:, col_idx[i]]
        # 将add_data的NaN替换为空字符串
        add_data = add_data.drop_duplicates()
        add_data['P_KEY'] = add_data['季节号'].astype(str) + '_' + add_data['PO号'].astype(str) + '_' + add_data['款式名称'].astype(str) + '_' + add_data['面料颜色'].astype(str)
        add_data['CreateDate'] = str(datetime.datetime.now()).split('.')[0]
        add_data['P_KEY'] = add_data['P_KEY'].str.strip()
        # keyList去重
        self.keyList = list(set(add_data['季节号'].tolist()))
        for column in add_data:
            if column in self.number_item[:3]:
                # 将add_data的NaN替换为0
                add_data[column].fillna(0, inplace=True)
                add_data[column] = add_data[column].astype(int)
            elif column in self.number_item[3:]:
                add_data[column].fillna(0, inplace=True)
                add_data[column] = add_data[column].astype(float)
            elif column in self.date_item:
                add_data[column].fillna('1977-01-01', inplace=True)
                add_data[column] = pd.to_datetime(add_data[column])
            else:
                add_data[column].fillna('', inplace=True)
                add_data[column] = add_data[column].astype(str)
        self.table_value.append([tuple(row) for row in add_data.values])
        # 追加数据
        print(self.table_value)
        self.update_db()
        print('已经完成数据操作！')
        input('按回车退出 ')

    def update_db(self):
        dbCol = self.add_data_title[:]
        dbCol.append('P_KEY')
        dbCol.append('CreateDate')
        # 建立连接并获取cursor
        conn = pymssql.connect(self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        # 删除已经存在的文件
        if len(self.keyList) > 0:
            keylist = list(set(self.keyList))
            del_tuple = []
            for tuple_po in keylist:
                del_tuple.append((tuple_po, tuple_po))
            delSql = 'delete from D_1DepIzacPoInfo where 季节号 = (%s)'
            cursor.executemany(delSql, del_tuple)
        # 组装插入的值
        insertValue = []
        for tabVal in self.table_value:
            insertValue += tabVal
        insertSql = 'INSERT INTO D_1DepIzacPoInfo VALUES ('
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

def gui_start():
    VAS = VAS_GUI()
    VAS.commit_batch()


if __name__ == '__main__':
    gui_start()
