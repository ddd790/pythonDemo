import os
import pandas as pd
import pdfplumber
import re
import pymssql
import datetime
import shutil

class VAS_GUI():
    # 批量获取服务器数据，进行累加操作
    def get_files(self):
        print('数据操作进行中......' + str(datetime.datetime.now()).split('.')[0])
        # sql服务器名
        self.serverName = '192.168.0.11'
        # 登陆用户名和密码
        self.userName = 'sa'
        self.passWord = 'jiangbin@007'
        # 数据库名
        self.dbName = 'ESApp1'
        # 追加的dataFrame的title
        self.add_data_title = ['OrderNo', 'SupplierNo', 'Style', 'Colour', 'Size', 'Qty', 'Total', 'Price', 'Currency', 'DeliveryDate', 'SetNo', 'FileName']
        # 数字类型的字段
        self.number_item = ['Qty', 'Total', 'Price']
        # 文件位置
        self.local_pdf_detail_file = 'd:\\BB裁单'

        # 最终dataframe
        self.table_data = pd.DataFrame(data=None, columns=self.add_data_title)
        # 删除列表
        self.delete_key = []
        # 循环文件，处理合并
        for lroot, ldirs, lfiles in os.walk(self.local_pdf_detail_file):
            for lfile in lfiles:
                self.file_to_dataframe(os.path.join(lroot, lfile), str(lfile).split('.')[0])
        self.table_data['CreateDate'] = str(datetime.datetime.now()).split('.')[0]
        self.table_value = []
        self.table_value.append([tuple(row) for row in self.table_data.values])
        # 更新数据库
        self.update_db()
        # 删除目录内文件
        self.delete_files_in_folder(self.local_pdf_detail_file)
        print('已经完成操作！' + str(datetime.datetime.now()).split('.')[0])
        input('按回车退出 ')

    def file_to_dataframe(self, io, lfile):
        pdf = pdfplumber.open(io)
        # PO号, 供应商编号, 款号
        po_no = ''
        supplier_no = ''
        style_no = ''
        for page in pdf.pages:
            # 文件前面的非表格内容
            file_txt = str(page.extract_text()).split('Total')[0]
            po_no = self.get_value_two_word(file_txt, 'Order-No.:', 'Supplier-No.:').strip()
            supplier_no = self.get_value_two_word(file_txt, 'Supplier-No.:', 'VAT no.:').replace('\n', '').strip()
            style_no = self.get_value_two_word(file_txt, 'Our Style:', 'Collection-Code:').replace('\n', '').strip()
            table = page.extract_tables()[0]
            # 尺码
            size_list = []
            # 尺码索引的集合
            size_index = []
            # 循环表格的第一行，获取尺码
            for idx, row in enumerate(table[0]):
                # 判断row是否含有字符串'Quantity'
                if idx > 0 and row.find('Quantity') == -1:
                    row = row.replace('Sizes\n', '')
                    size_list.append(row)
                    size_index.append(idx)
                elif idx == 0:
                    continue
                else:
                    break
            # 循环表格的第二行到最后一行
            for idx, row in enumerate(table[1:]):
                if row[0] == '':
                    continue
                # 循环size_list的长度，获取每一行的数据
                for idx_s, size in enumerate(size_list):
                    # ['PO号', '供应商编号', '款号', '颜色', '尺码', '数量', '总数量', '价格', '币种', '船期', 'set号', '文件名']
                    tmp_date_row = [str(po_no), str(supplier_no), str(style_no), str(row[0]).replace(' ', ''), str(size), row[size_index[idx_s]].replace(' ', ''), row[size_index[-1] + 1], 
                                    row[size_index[-1] + 2].replace(',', '.'), str(row[size_index[-1] + 3]), self.change_date(row[size_index[-1] + 4]), 
                                    row[size_index[-1] + 5], lfile]
                    # 创建要追加的新行数据
                    new_row = pd.DataFrame([tmp_date_row], columns=self.add_data_title)
                    # 使用concat()函数对dataframe进行合并
                    self.table_data = pd.concat([self.table_data, new_row], axis=0).reset_index(drop=True)
                    self.delete_key.append(po_no + '_' + supplier_no + '_' + style_no + '_' + row[0].replace(' ', ''))
        # 删除table_data中的数量是空的行, 并重新设置索引
        self.table_data = self.table_data[self.table_data['Qty'] != ''].reset_index(drop=True)
        # 将table_data中的数量转化为数字
        self.table_data['Qty'] = self.table_data['Qty'].astype(int)
        self.table_data['Total'] = self.table_data['Total'].astype(int)
        self.table_data['Price'] = self.table_data['Price'].astype(float)
        self.table_data['DelKey'] = po_no + '_' + supplier_no + '_' + style_no
        # DelKey列的内容追加Colour的内容
        self.table_data['DelKey'] = self.table_data['DelKey'] + '_' + self.table_data['Colour']
        # 关闭文件
        pdf.close()

    # 获取一个字符串中两个字母中间的值(one为None时从第一位取, two为None时取到最后)
    def get_value_two_word(self, txt_str, one, two):
        if one == None:
            return txt_str[:txt_str.find(two)]
        if two == None:
            return txt_str[txt_str.find(one) + len(one):]
        return txt_str[txt_str.find(one) + len(one):txt_str.find(two)]
    
    # 转化日期格式
    def change_date(self, date_str):
        # date_str的格式为'28.06.2025', 转化为'2025-06-28 00:00:00'
        return date_str[-4:] + '-' + date_str[3:5] + '-' + date_str[:2] + ' 00:00:00'

    def update_db(self):
        dbCol = self.add_data_title[:]
        dbCol.append('DelKey')
        dbCol.append('CreateDate')
        # 建立连接并获取cursor
        conn = pymssql.connect(self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        if len(self.delete_key) > 0:
            keylist = list(set(self.delete_key))
            del_tuple = []
            for tuple_po in keylist:
                del_tuple.append((tuple_po, tuple_po))
            delSql = 'delete from D_3DepCai_B where DelKey = (%s)'
            cursor.executemany(delSql, del_tuple)
        # 组装插入的值
        insertValue = []
        for tabVal in self.table_value:
            insertValue += tabVal
        insertSql = 'INSERT INTO D_3DepCai_B (' + (",".join(str(i) for i in dbCol)) + ') VALUES ('
        for colVal in dbCol:
            if colVal == 'CreateDate':
                insertSql += '%s'
            elif colVal in self.number_item:
                insertSql += '%d, '
            else:
                insertSql += '%s, '
        insertSql += ')'
        # print(insertSql)
        # print(insertValue)
        cursor.executemany(insertSql, insertValue)
        conn.commit()
        conn.close()
    
    def delete_files_in_folder(self, folder_path):
        # 确保提供的路径是一个有效的文件夹路径
        if not os.path.isdir(folder_path):
            # print(f"Path {folder_path} is not a valid directory.")
            return

        # 获取文件夹中的所有文件
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            
            # 检查是否为文件
            if os.path.isfile(file_path):
                try:
                    # 删除文件
                    os.remove(file_path)
                    # print(f"Deleted file: {file_path}")
                except Exception as e:
                    print(f"Failed to delete {file_path}. Error: {e}")


def gui_start():
    VAS = VAS_GUI()
    VAS.get_files()


if __name__ == '__main__':
    gui_start()
