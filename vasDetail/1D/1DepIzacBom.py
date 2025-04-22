import os
import pandas as pd
import datetime
import pymssql
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
        self.base_title = ['品类', '品号', '用料颜色', '规格', '供应商', '用料名称', '单耗', '单位', '缩水率', '损耗率', '备注']
        self.jv_title = ['品类', '品号', '用料颜色', '规格', '供应商', '用料名称', '单耗', '单位', '损耗率', '单价', '美金单价', '克重', '成份', '起订量', '小缸费', '生产周期', '备注']
        self.add_data_title = self.base_title.copy()
        self.add_data_title.extend(['BOM类型', '款号', 'PO号', '颜色', 'delKey', 'rowNum', '客户'])
        # 数字类型的字段
        self.number_item = ['单耗']
        self.local_cai_detail_file = 'd:\\IZACBOM'

        # 循环文件，处理合并
        self.table_value = []
        # 删除文件的list
        self.delList = []
        for lroot, ldirs, lfiles in os.walk(self.local_cai_detail_file):
            for lfile in lfiles:
                print(lfile)
                if not str(lfile).__contains__('~'):
                    print('文件名：' + str(lfile).split('.')[0])
                    df = pd.read_excel(os.path.join(lroot, lfile), sheet_name=None, skiprows=1, names=self.base_title, keep_default_na=False)
                    self.format_dataframe(df, lfile, os.path.join(lroot, lfile))
        # 更新数据库，删除文件
        self.update_db()
        # self.delete_files_in_folder(self.local_cai_detail_file)
        input('按回车退出 ')

    def format_dataframe(self, df, lfile, file_path):
        # 新建dataframe
        table_data_add = pd.DataFrame(data=None, columns=self.add_data_title)
        for sheet_name, sheet_data in df.items():
            # 读取当前sheet
            df_po = pd.read_excel(file_path, sheet_name=sheet_name, header=None, keep_default_na=False)
            # 根据A1的内容判断客户
            customer_flag = str(df_po.iloc[0, 0])
            # 获取B1和D1的内容
            style_no = str(df_po.iloc[0, 1])  # B1单元格的内容
            po_no = str(df_po.iloc[0, 3])  # D1单元格的内容
            customer = 'IZAC'
            # 如果A1的内容是“STYLE NUMBER*订单号”，为JV客户
            if customer_flag.__contains__('STYLE NUMBER*订单号'):
                style_no = str(df_po.iloc[0, 3])  # B1单元格的内容
                po_no = str(df_po.iloc[0, 1])  # D1单元格的内容
                customer = 'JV'
            # 对每个sheet页的数据进行处理
            table_data= pd.DataFrame(sheet_data, columns=self.base_title)
            # 如果customer_flag的内容是“STYLE NUMBER*订单号”，按照新列读取
            if customer_flag.__contains__('STYLE NUMBER*订单号'):
                table_data= pd.DataFrame(None, columns=self.base_title)
                df_po.drop(df_po.index[0:2], inplace=True)
                df_po = df_po.reset_index(drop=True)
                # ['品类', '品号', '用料颜色', '规格', '供应商', '用料名称', '单耗', '单位', '缩水率', '损耗率', '备注']
                table_data['品类'] = df_po[0]
                table_data['品号'] = df_po[1]
                table_data['用料颜色'] = df_po[2]
                table_data['规格'] = df_po[3]
                table_data['供应商'] = df_po[4]
                table_data['用料名称'] = df_po[5]
                table_data['单耗'] = df_po[6]
                table_data['单位'] = df_po[7]
                table_data['缩水率'] = ''
                table_data['损耗率'] = df_po[9]
                # 将table_data_jv中的备注列赋值给table_data中的备注列
                table_data['备注'] = df_po[16]
                # 删除品类为空的行，并删除索引
                table_data = table_data[table_data['品类'] != '']
                table_data = table_data.reset_index(drop=True)
            # table_data['单耗']列如果为''，默认为1.0
            table_data['单耗'] = table_data['单耗'].replace('', '0.0')
            table_data.fillna('', inplace=True)
            # 查找“品类”列为 "VAS" 的第一个位置
            vas_first_index = table_data[table_data['品类'] == 'VAS'].index.min()
            # 从 "VAS" 分割 DataFrame
            # 上半部分的行
            # print(vas_first_index)
            upper_half = table_data.iloc[:vas_first_index]
            upper_half['BOM类型'] = '面辅料'
            # upper_half.drop(upper_half.index[-1], inplace=True)
            # 下半部分的行
            lower_half = table_data.iloc[vas_first_index + 1:]
            lower_half['BOM类型'] = 'VAS'
            # 合并 upper_half 和 lower_half
            table_data = pd.concat([upper_half, lower_half], ignore_index=True)
            table_data['品号'] = table_data['品号'].astype(str)
            table_data['用料颜色'] = table_data['用料颜色'].astype(str)
            table_data['规格'] = table_data['规格'].astype(str)
            table_data['缩水率'] = table_data['缩水率'].astype(str)
            table_data['损耗率'] = table_data['损耗率'].astype(str)
            table_data['单耗'] = table_data['单耗'].astype(float)
            table_data['款号'] = style_no
            table_data['PO号'] = po_no
            table_data['颜色'] = sheet_name
            table_data['delKey'] = style_no + '-' + po_no + '-' + sheet_name
            table_data['rowNum'] = table_data.index + 1
            table_data['客户'] = customer
            table_data['FileName'] = str(lfile).split('.')[0]
            table_data['CreateDate'] = str(datetime.datetime.now()).split('.')[0]
            # 删除空行
            table_data.dropna(subset=['品号'], inplace=True)
            table_data_add = table_data_add.append(table_data, ignore_index=True)
        self.delList.extend(table_data_add['delKey'].tolist())
        self.table_value.append([tuple(row) for row in table_data_add.values])
        # print(self.table_value)

    def update_db(self):
        dbCol = self.add_data_title[:]
        dbCol.append('FileName')
        dbCol.append('CreateDate')
        # 建立连接并获取cursor
        conn = pymssql.connect(self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        # 组装删除的值
        del_tuple = tuple(self.delList)
        # 删除已经存在的文件
        delSql = 'delete from D_1DepIzacBom where delKey = (%s)'
        cursor.executemany(delSql, del_tuple)
        # 组装插入的值
        insertValue = []
        for tabVal in self.table_value:
            insertValue += tabVal
        insertSql = 'INSERT INTO D_1DepIzacBom VALUES ('
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
