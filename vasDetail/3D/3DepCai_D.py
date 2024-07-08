import os
import pandas as pd
import pdfplumber
import re
import datetime
import shutil
from decimal import Decimal
from aip import AipOcr
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
        self.add_data_title = ['PK','Order_No', 'Style_No', 'Fabric_No', 'Fabric_color', 'Size', 'Number', 'FileName', 'CreateDate']
        # 数字类型的字段
        self.number_item = ['Number']
        # 服务器发票文件路径
        self.local_list_file = 'D:\\DIGEL裁单'
        # 删除文件的list
        self.keyList = []
        # 最终dataframe
        self.table_data = pd.DataFrame(data=None, columns=self.add_data_title)
        self.table_value = []
        # 循环文件，处理合并
        for lroot, ldirs, lfiles in os.walk(self.local_list_file):
            for lfile in lfiles:
                print(lfile)
                self.file_to_dataframe(os.path.join(lroot, lfile), str(lfile).split('.')[0])
        self.table_value.append([tuple(row) for row in self.table_data.values])
        # 更新数据库
        self.update_db()
        # 删除目录内文件
        # if os.path.exists(self.local_list_file):
        #     shutil.rmtree(self.local_list_file, onerror=self.readonly_handler)
        # os.mkdir(self.local_list_file)
        # 回车退出
        print('------------------------------------------------------------')
        print('已经完成操作！' + str(datetime.datetime.now()).split('.')[0])
        input('按回车退出 ')

    def file_to_dataframe(self, io, lfile):
        pdf_df = pd.DataFrame(data=None, columns=self.add_data_title)
        pdf = pdfplumber.open(io)
        # 尺码列表
        size_info_list = []
        size_list = []
        number_list = []
        style_No_list = []
        fabric_No_list = []
        fabric_color_list = []
        order_no = ''
        # 裁单的key
        key_dic = {}
        # 打开的PDF文件  
        for page in pdf.pages:
            text = page.extract_text()
            order_no = text.split('Order No.:')[1].strip()[:5]
            detali_info_list = text.split('Delivery date')[1:]
            for detali_info in detali_info_list:
                # print(detali_info)
                # 面料号和颜色号拼成key
                size_info_list = []
                style = str(detali_info.split('\n')[1].split(' ')[2]).strip()
                fabric = str(detali_info.split('\n')[1].split(' ')[1]).strip()
                color = str(detali_info.split('\n')[2]).strip()
                tmp_key = fabric + '^' + color + '^' + style + '^' + order_no
                for d_info in detali_info.split('\n'):
                    if d_info.strip().__contains__('/'):
                        size_info_list.append(d_info)
                key_dic[tmp_key] = size_info_list
        # print(key_dic)
        for key, value in key_dic.items():
            for size_info in list(value):
                # print(size_info)
                # print(key)
                for tmp in size_info.split(' '):
                    if str(tmp).strip().__contains__('/'):
                        size_list.append(str(tmp).strip().split('/')[1].replace(',', '.'))
                        number_list.append(int(str(tmp).strip().split('/')[0]))
                        self.keyList.append(key)
                        fabric_No_list.append(key.split('^')[0])
                        fabric_color_list.append(key.split('^')[1])
                        style_No_list.append(key.split('^')[2])
        # 'PK','Order_No','style_No', 'Fabric_No', 'Fabric_color', 'Size', 'Number', 'CreateDate'
        pdf_df.loc[:, self.add_data_title[0]] = self.keyList
        pdf_df.loc[:, self.add_data_title[1]] = order_no
        pdf_df.loc[:, self.add_data_title[2]] = style_No_list
        pdf_df.loc[:, self.add_data_title[3]] = fabric_No_list
        pdf_df.loc[:, self.add_data_title[4]] = fabric_color_list
        pdf_df.loc[:, self.add_data_title[5]] = size_list
        pdf_df.loc[:, self.add_data_title[6]] = number_list
        pdf_df.loc[:, self.add_data_title[7]] = lfile
        pdf_df.loc[:, self.add_data_title[8]] = str(datetime.datetime.now()).split('.')[0]
        self.table_data = self.table_data.append(pdf_df, ignore_index=True)
        pdf.close()

    # 获取一个字符串中两个字母中间的值(one为None时从第一位取, two为None时取到最后)
    def get_value_two_word(self, txt_str, one, two):
        if one == None:
            return txt_str[:txt_str.find(two)]
        if two == None:
            return txt_str[txt_str.find(one) + len(one):]
        return txt_str[txt_str.find(one) + len(one):txt_str.find(two)]

    def update_db(self):
        dbCol = self.add_data_title[:]
        # 建立连接并获取cursor
        conn = pymssql.connect(self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        if len(self.keyList) > 0:
            keylist = list(set(self.keyList))
            del_tuple = []
            for tuple_po in keylist:
                del_tuple.append((tuple_po, tuple_po))
            delSql = 'delete from D_3DepCai_D where PK = (%s)'
            cursor.executemany(delSql, del_tuple)
        # 组装插入的值
        insertValue = []
        for tabVal in self.table_value:
            insertValue += tabVal
        insertSql = 'INSERT INTO D_3DepCai_D (' + (",".join(str(i) for i in dbCol)) + ') VALUES ('
        for colVal in dbCol:
            if colVal == 'CreateDate':
                insertSql += '%s'
            elif colVal in self.number_item:
                insertSql += '%d, '
            else:
                insertSql += '%s, '
        insertSql += ')'
        # print(insertValue)
        cursor.executemany(insertSql, insertValue)
        conn.commit()
        conn.close()
    
    # 文件只读删除的解决
    def readonly_handler(self, func, path, exc_info):
        os.chmod(path, stat.S_IWRITE)
        func(path)

def gui_start():
    VAS = VAS_GUI()
    VAS.get_files()


if __name__ == '__main__':
    gui_start()
