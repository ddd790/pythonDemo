import os
import pandas as pd
import pdfplumber
import re
import datetime
import shutil
from decimal import Decimal
import pymssql
import stat


class VAS_GUI():
    # 电子发票数据提取
    def get_files(self):
        # print('数据操作进行中......' + str(datetime.datetime.now()).split('.')[0])
        # sql服务器名
        self.serverName = '192.168.0.11'
        # 登陆用户名和密码
        self.userName = 'sa'
        self.passWord = 'jiangbin@007'
        # 数据库名
        self.dbName = 'ESApp1'
        # 追加的dataFrame的title
        # self.add_data_title = ['发票号码', '开票日期','购买方名称', '购买方纳税人识别号', '销售方名称', '销售方纳税人识别号', '项目名称', '规格型号', '单位', '数量', '单价', '金额',  
        #                      '税率', '税额', '价税合计', '备注', '文件名', '文件类型', '创建时间']
        self.add_data_title = ['InvoiceNo', 'InvoiceDate', 'BuyName', 'BuyNo', 'SellName', 'SellNo', 'Name', 'Size', 'Unit', 'Number', 'UnitPrice', 'Price',
                               'Rate', 'Tax', 'TotalPrice', 'Remarks', 'FileName', 'FileType', 'CreateDate']
        # 数字类型的字段
        self.number_item = ['Number', 'UnitPrice', 'Price', 'Rate', 'Tax', 'TotalPrice',]
        # 服务器发票文件路径
        networked_directory = r'\\192.168.0.3\04-业务四部共享'
        # self.local_list_file = 'd:\\fapiaoTest'
        self.local_list_file = 'd:\\haiyun\\1'
        # 遍历服务器文件
        # for root, dirs, files in os.walk(networked_directory):
        #     for file in files:
        #         if (str(file).__contains__('.xlsx') or str(file).__contains__('.xls')) and not str(file).__contains__('~') and str(file).__contains__('海运大表'):
        #             shutil.copy2(os.path.join(root, file), self.local_list_file)
        all_data = pd.DataFrame()
        # 遍历文件夹中的所有文件
        for filename in os.listdir(self.local_list_file):
            if filename.endswith('.xlsx') or filename.endswith('.xls'):
                file_path = os.path.join(self.local_list_file, filename)
                # 读取Excel文件的第一个sheet
                df = pd.read_excel(file_path, sheet_name=0)
                # 将所有单元格内容转换为字符串
                df = df.astype(str)
                all_data = pd.concat([all_data, df], ignore_index=True)
        
        # 将合并后的数据写入新的Excel文件
        all_data.to_excel('results.xlsx', index=False)
        print('------------------------------------------------------------')
        print('已经完成操作！' + str(datetime.datetime.now()).split('.')[0])
        # 回车退出
        input('按回车退出 ')

def gui_start():
    VAS = VAS_GUI()
    VAS.get_files()


if __name__ == '__main__':
    gui_start()
