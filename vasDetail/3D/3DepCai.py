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
        self.add_data_title = ['Saison', 'Kollektion', 'Produktionsdekade', 'Produzent', 'Auftrag', 'Liefertermin', 'Artikel', 'Artikelbezeichnung', 'Grundware', 'Type', 'Size', 'Qty', 'FileName']
        # 数字类型的字段
        self.number_item = ['Qty']
        # 根据勤哲的key匹配对应trimList中的key和value
        self.local_file = 'd:\\3DepCai'
        # 最终dataframe
        self.table_data = pd.DataFrame(data=None, columns=self.add_data_title)
        self.table_value = []
        # 删除列表
        self.delete_key = []
        # 循环文件，处理合并
        for lroot, ldirs, lfiles in os.walk(self.local_file):
            for lfile in lfiles:
                self.file_to_dataframe(os.path.join(lroot, lfile), str(lfile).split('.')[0])
        self.table_data['CreateDate'] = str(datetime.datetime.now()).split('.')[0]
        self.table_value.append([tuple(row) for row in self.table_data.values])
        # 删除项去重
        self.delete_item = list(set(self.delete_key))

        # 更新数据库
        self.update_db()
        print('已经完成操作！' + str(datetime.datetime.now()).split('.')[0])
        input('按回车退出 ')

    def file_to_dataframe(self, io, lfile):
        print('-----------------------------')
        pdf = pdfplumber.open(io)
        for page in pdf.pages:
            pdf_df = pd.DataFrame(data=None, columns=self.add_data_title)
            # 尺码信息
            tmp_size_list = []
            # PO基本信息
            Saison= ''
            Kollektion= ''
            Produktionsdekade= ''
            Produzent= ''
            Auftrag= ''
            Liefertermin= ''
            Artikelbezeichnung= ''
            Artikel= ''
            Grundware= ''
            table_num = 0
            for table in page.extract_tables():
                if table.__len__() < 3:
                    continue
                type_val = {}
                table_num += 1
                # 第一个表格是PO基本信息
                if table_num == 1:
                    table_i = 0
                    for i in range(0, len(table)):
                        table_tmp = [item for item in table[i] if item is not None]
                        if table_tmp.__len__() != 3:
                            continue
                        table_i += 1
                        if table_i == 1:
                            Saison = table_tmp[0].split('\n')[1]
                            Kollektion = table_tmp[1].split('\n')[1]
                            Produktionsdekade = table_tmp[2].split('\n')[1]
                        elif table_i == 2:
                            Produzent = table_tmp[0].split('\n')[1]
                            Auftrag = table_tmp[1].split('\n')[1]
                            Liefertermin = table_tmp[2].split('\n')[1]
                        elif table_i == 3:
                            Artikelbezeichnung = table_tmp[0].split('\n')[1]
                            Artikel = table_tmp[1].split('\n')[1]
                            Grundware = table_tmp[2].split('\n')[1]
                # 第二个表格是配码信息
                else:
                    # 第一行是配码，第二行到倒数第二行中，第一列是type，后面的数据是该配码下的数量，
                    title_count = 0
                    for row in table:
                        if title_count == 0:
                            tmp_size_list = row[1:len(row) - 1]
                        elif title_count < len(table) - 1:
                            size_list = row[1:len(row) - 1]
                            type_val[row[0]] = size_list
                        title_count += 1
            for k in type_val.keys():
                pdf_df.loc[:, self.add_data_title[10]] = tmp_size_list
                pdf_df.loc[:, self.add_data_title[9]] = k
                pdf_df.loc[:, self.add_data_title[11]] = type_val[k]
                pdf_df.loc[:, self.add_data_title[0]] = Saison
                pdf_df.loc[:, self.add_data_title[1]] = Kollektion
                pdf_df.loc[:, self.add_data_title[2]] = Produktionsdekade
                pdf_df.loc[:, self.add_data_title[3]] = Produzent
                pdf_df.loc[:, self.add_data_title[4]] = Auftrag
                pdf_df.loc[:, self.add_data_title[5]] = self.change_date(Liefertermin)
                pdf_df.loc[:, self.add_data_title[6]] = Artikelbezeichnung
                pdf_df.loc[:, self.add_data_title[7]] = Artikel
                pdf_df.loc[:, self.add_data_title[8]] = Grundware
                pdf_df.loc[:, self.add_data_title[12]] = lfile
                self.table_data = self.table_data.append(pdf_df, ignore_index=True)
            # 删除Qty为空的数据
            self.table_data['Qty'].replace('', 0, inplace=True)
            self.delete_key.append(lfile)

    def update_db(self):
        dbCol = self.add_data_title[:]
        dbCol.append('CreateDate')
        # 建立连接并获取cursor
        conn = pymssql.connect(self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        # 组装删除的值
        del_tuple = tuple(self.delete_item)
        # # 删除已经存在的文件
        delSql = 'delete from D_3DepCai where FileName = (%s)'
        cursor.executemany(delSql, del_tuple)
        # 组装插入的值
        insertValue = []
        for tabVal in self.table_value:
            insertValue += tabVal
        insertSql = 'INSERT INTO D_3DepCai (' + (",".join(str(i) for i in dbCol)) + ') VALUES ('
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
    
    def change_date(self, date):
        # 使用正则表达式提取所有数字
        digits = re.sub(r'\D', '', str(date))
        return str(digits[4:]) + '-' + str(digits[2:4]) + '-' + str(digits[:2])

def gui_start():
    VAS = VAS_GUI()
    VAS.get_files()


if __name__ == '__main__':
    gui_start()
