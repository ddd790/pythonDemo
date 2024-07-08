import os
import pandas as pd
import datetime
import pymssql
import shutil
from tkinter import *


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
        self.add_data_title = ['PO_No', 'Dept', 'Spec_Style', 'SKU', 'SKU_DESCRIPTION', 'Ticket_Style', 'RMS_Color_Code ', 'RMS_Color_Description','SIZE_CODE', 
                               'SIZECODE_DESCRIPTION', 'Price', 'Qty', 'Total_Qty', 'Tot_Color_Qty', 'Vendor_id', 'Vendor', 'Factory',  'POL', 'Delivery_Terms', 
                               'GAC', 'NDC', 'Mode']
        # 数字类型的字段
        self.number_item = ['Price', 'Qty', 'Total_Qty', 'Tot_Color_Qty']
        self.local_cai_detail_file = 'd:\EXPREE裁单'

        # 循环文件，处理合并
        self.table_value = []
        # 删除文件的list
        self.SKUList = []
        for lroot, ldirs, lfiles in os.walk(self.local_cai_detail_file):
            for lfile in lfiles:
                if not str(lfile).__contains__('~'):
                    print('文件名：' + str(lfile).split('.')[0])
                    df = pd.read_excel(os.path.join(lroot, lfile), sheet_name=0, nrows=1000,converters={'Ticket Style ':str, 'RMS Color Code ':str})
                    table_data = pd.DataFrame(df)
                    table_data.columns = self.add_data_title
                    table_data['FileName'] = str(lfile).split('.')[0]
                    table_data['CreateDate'] = str(datetime.datetime.now()).split('.')[0]
                    self.SKUList.extend(table_data['SKU'].tolist())
                    self.table_value.append([tuple(row) for row in table_data.values])

        # 更新数据库，删除文件
        self.update_db()
        if os.path.exists(self.local_cai_detail_file):
            shutil.rmtree(self.local_cai_detail_file)
        os.mkdir(self.local_cai_detail_file)
        input('按回车退出 ')

    def update_db(self):
        dbCol = self.add_data_title[:]
        dbCol.append('FileName')
        dbCol.append('CreateDate')
        # 建立连接并获取cursor
        conn = pymssql.connect(self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        # 组装删除的值
        del_tuple = tuple(self.SKUList)
        # 删除已经存在的文件
        delSql = 'delete from D_3DepCai_E where SKU = (%s)'
        cursor.executemany(delSql, del_tuple)
        # 组装插入的值
        insertValue = []
        for tabVal in self.table_value:
            insertValue += tabVal
        insertSql = 'INSERT INTO D_3DepCai_E VALUES ('
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
    VAS.get_files()


if __name__ == '__main__':
    gui_start()
