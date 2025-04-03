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
        self.add_data_title = ['来单日期', '客户', '季节号', '订单号', '款式名称', '客户面料号', '款式缩写', '款式', '英文款名', 
                               '中文款名', '类别', '面料描述', '面料颜色', '数量', '价格', '总金额', '零售单价', '交期', 
                               '客人PO成分', '工厂', '是否有工艺', '是否录BOM', '生产部门', '面料发货时间', '面料到厂时间', 
                               '辅料采购发货时间', '给采购的最晚到料时间', '供应商', '面料号', '面料成分', '克重', '幅宽', 
                               '面料价格', '实际采购面料价格', '采购单耗', '大货单耗', '加减裁数量', '加工费', 
                               '面料成本', '辅料成本', '运费', '操作费用', '关税', '税率', '汇率', '样品轮数', '样品数量', 
                               '样品状态', '走货发票号', '挂法', '贸易方式', '海关编码']
        # 数字类型的字段
        self.number_item = ['数量', '加减裁数量', '样品轮数', '样品数量', '价格', '总金额', '零售单价', '面料价格', '实际采购面料价格', '采购单耗', '大货单耗', 
                            '加工费', '面料成本', '辅料成本', '运费', '操作费用', '关税', '税率', '汇率']
        # 日期类型的字段
        self.date_item = ['来单日期', '交期', '面料发货时间', '面料到厂时间', '辅料采购发货时间', '给采购的最晚到料时间']
        # 循环文件，处理合并，并存入数据库
        self.local_vas_detail_file = r'\\192.168.0.3\01-业务一部资料\A  JV\JV PO表.xlsx'
        # self.local_vas_detail_file = r'D:\temp\H25 IZAC PO.xlsx'
        self.table_value = []
        # 删除文件的list
        self.keyList = []
        # 读取A到AZ列的内容,从第三行开始读取
        add_data = pd.read_excel(self.local_vas_detail_file, sheet_name=0, header=1, usecols='A:AZ', dtype=str)
        # 将df的列名转换为dataFrame的列名
        add_data.columns = self.add_data_title
        # 将add_data的NaN替换为空字符串
        add_data = add_data.drop_duplicates()
        # 将P_KEY的值修改为从1开始的字符串
        add_data['P_KEY'] = (add_data.index + 1).astype(str)
        add_data['CreateDate'] = str(datetime.datetime.now()).split('.')[0]
        # keyList去重
        self.keyList = list(set(add_data['季节号'].tolist()))
        for column in add_data:
            if column in self.number_item[:4]:
                # 将add_data的NaN替换为0
                add_data[column].fillna(0, inplace=True)
                add_data[column] = add_data[column].astype(int)
            elif column in self.number_item[4:]:
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
            delSql = 'delete from D_1DepJVPoInfo where 季节号 = (%s)'
            cursor.executemany(delSql, del_tuple)
        # 组装插入的值
        insertValue = []
        for tabVal in self.table_value:
            insertValue += tabVal
        insertSql = 'INSERT INTO D_1DepJVPoInfo VALUES ('
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
