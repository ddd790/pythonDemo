import os
import pandas as pd
import datetime
import pymssql
import pdfplumber
import shutil
import time
from dateutil import parser
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
        self.add_data_title = ['FileName', 'BrandType', 'CreateDate']
        # 数字类型的字段
        self.number_item = ['Qty', 'Price']
        networked_directory = r'\\192.168.0.3\01-业务一部资料\=14785212\PEERLESS\国内埃塞柬埔寨订单信息\临时'
        self.local_pdf_detail_file = 'd:\peerlessPdfBrandType'

        # 删除目录内文件
        if os.path.exists(self.local_pdf_detail_file):
            shutil.rmtree(self.local_pdf_detail_file)
        os.mkdir(self.local_pdf_detail_file)
        # copy服务器的TRIMLIST文件到本地
        for root, dirs, files in os.walk(networked_directory):
            if root.__contains__('2024') or root.__contains__('2025'):
                for file in files:
                    if str(file).__contains__('PO-') and (str(file).__contains__('.pdf') or str(file).__contains__('.PDF')):
                        shutil.copy2(os.path.join(root, file), self.local_pdf_detail_file)

        # 查询已存在的记录
        self.select_po_old_value()
        # 已经存在的文件名(去重)
        file_name_old = list(set(self.old_all_data['FileName']))

        self.pdf_data_val = []
        # 文件名的list
        self.fileNameList = []
        for lroot, ldirs, lfiles in os.walk(self.local_pdf_detail_file):
            for file in lfiles:
                file_name = str(file).split('.')[0]
                # print('文件名：' + str(file).split('-V')[0])
                self.fileNameList.append(str(file).split('.')[0])
                if (str(file).__contains__('.pdf') or str(file).__contains__('.PDF')) and file_name not in file_name_old:
                    mtime = parser.parse(time.ctime(os.path.getmtime(os.path.join(lroot, file))))
                    create_time = mtime.strftime('%Y-%m-%d %H:%M:%S')
                    self.file_to_dataframe_pdfplumber(file, create_time)
        # 将pdf结果转成dataFrame
        table_data = pd.DataFrame(self.pdf_data_val, columns=self.add_data_title)
        self.table_value = []
        self.table_value.append([tuple(row) for row in table_data.values])

        # 更新数据库
        self.update_db()
        print('已经完成操作！' + str(datetime.datetime.now()).split('.')[0])

    def file_to_dataframe_pdfplumber(self, fileName, create_time):
        # 文件名
        tx_fileName = str(fileName).split('.')[0]
        # print('新增文件名：' + tx_fileName)
        # PO
        tx_po = tx_fileName.split('-')[1]
        pdfreader = pdfplumber.open(self.local_pdf_detail_file + '\\' + fileName)
        tx_val = []
        # 循环读取pdf内容
        for index in range(len(pdfreader.pages)):
            pageReader = pdfreader.pages[index]
            pageObj = pageReader.extract_text()  # 获取内容
            tx_val.append(pageObj)
        # 每页中的元素拼接
        tx = '\n'.join(tx_val)
        detail_info = []
        # 文件名
        detail_info.append(tx_fileName)
        if tx.__contains__('PRIVATE LABEL'):
            detail_info.append('PRIVATE LABEL')
        else:
            detail_info.append('')
        detail_info.append(create_time)
        self.pdf_data_val.append(detail_info)

    def update_db(self):
        dbCol = self.add_data_title[:]
        # 建立连接并获取cursor
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        # 组装插入的值
        insertValue = []
        for tabVal in self.table_value:
            insertValue += tabVal
        insertSql = 'INSERT INTO D_Peerless_Order_BrandType (' + (
            ",".join(str(i) for i in dbCol)) + ') VALUES ('
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

    def select_po_old_value(self):
        # 建立连接并获取PO数据
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        strCol = ",".join(str(i) for i in self.add_data_title)
        select_sql = 'select ' + strCol + ' from D_Peerless_Order_BrandType'
        cursor.execute(select_sql)
        row = cursor.fetchall()
        self.old_all_data = pd.DataFrame(data=list(row), columns=self.add_data_title)
        cursor.close()
        conn.close()

def gui_start():
    VAS = VAS_GUI()
    VAS.get_files()


if __name__ == '__main__':
    gui_start()
