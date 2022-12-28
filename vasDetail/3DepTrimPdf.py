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
        self.add_data_title = ['款号', '品名', '面料号', '用料品号',
                               '供应商', '供应商品号', '数量', '规格', '物料说明', '物料颜色', '成衣颜色', 'version']
        # 数字类型的字段
        self.number_item = ['数量', 'version']
        # 根据勤哲的key匹配对应trimList中的key和value
        self.local_trim_list_file = r'\\192.168.0.3\03-业务三部共享\EXPRESS 工艺\大货 工艺书\勤哲BOM最新PDF文件'
        self.local_pdf_detail_file = 'd:\\3DepTrimlistPdfTemp'
        # 删除目录内文件
        if os.path.exists(self.local_pdf_detail_file):
            shutil.rmtree(self.local_pdf_detail_file)
        os.mkdir(self.local_pdf_detail_file)

        # copy服务器的TRIMLIST文件到本地
        for root, dirs, files in os.walk(self.local_trim_list_file):
            for file in files:
                if str(file).__contains__('.pdf') or str(file).__contains__('.PDF'):
                    shutil.copy(os.path.join(root, file),
                                self.local_pdf_detail_file)

        # 查询已存在的记录
        self.select_trim_old_value()
        # 合并p_key，找到对应的version
        self.old_all_data['p_key'] = self.old_all_data[self.add_data_title[0]] + '^*^' + \
            self.old_all_data[self.add_data_title[1]] + '^*^' + self.old_all_data[self.add_data_title[2]] + '^*^' + \
            self.old_all_data[self.add_data_title[10]]
        self.old_version = self.old_all_data.set_index("p_key")[
            "version"].to_dict()

        # 最终dataframe
        self.table_data = pd.DataFrame(data=None, columns=self.add_data_title)
        self.table_value = []
        # 删除列表
        self.delete_key = []
        # 循环文件，处理合并
        for lroot, ldirs, lfiles in os.walk(self.local_pdf_detail_file):
            for lfile in lfiles:
                self.file_to_dataframe(os.path.join(lroot, lfile), str(
                    lfile).split('.')[0])
        self.table_data['CreateDate'] = str(
            datetime.datetime.now()).split('.')[0]
        self.table_value = []
        self.table_value.append([tuple(row) for row in self.table_data.values])
        # 删除项去重
        # self.delete_item = list(set(self.delete_key))

        # 更新数据库
        self.update_db()
        print('已经完成操作！' + str(datetime.datetime.now()).split('.')[0])
        input('按回车退出 ')

    def file_to_dataframe(self, io, lfile):
        pdf = pdfplumber.open(io)
        df_title = []
        count = 0  # 页数
        # 款号, 品名, 面料号
        style_no = ''
        goods_name = ''
        fabric_no = ''
        for page in pdf.pages:
            df_values = []
            count += 1
            if count == 1:
                # page.extract_text()  # 抓取当前页的全部信息
                # 文件前面的非表格内容
                file_txt = str(page.extract_text()).split(
                    'Material Description')[0]
                style_no = self.get_value_two_word(
                    file_txt, 'Spec # ', ' Mat. One Content:').strip()[0:8].strip()
                goods_name = re.sub(r'[0-9]+', '', self.get_value_two_word(
                    file_txt, style_no, 'Wash/Fin: ').split(' - ')[2].replace('\n', '')).strip()
                fabric_no = file_txt.split('Mat. One ID#: ')[
                    1].replace('\n', '').strip()[0:len(style_no)].strip()
            for table in page.extract_tables():
                # title的行
                title_count = 0
                for row in table:
                    if title_count == 0:
                        df_title = row
                    else:
                        df_values.append(row)
                    title_count += 1

            # title颜色list
            color_list = []
            for idx in range(len(df_title)):
                if idx > 6 and idx < len(df_title) - 1 and df_title[idx] == '':
                    df_title[idx] = df_title[idx+1].replace('\n', '')
                    df_title[idx+1] = '^_^'
                    temp_color = df_title[idx].replace('\n', '')
                    if temp_color not in color_list:
                        color_list.append(temp_color)

            df = pd.DataFrame(df_values, columns=df_title)
            df.loc[:, self.add_data_title[0]] = style_no
            df.loc[:, self.add_data_title[1]] = goods_name
            df.loc[:, self.add_data_title[2]] = fabric_no
            # 删除错位的笑脸列（^_^）
            temp_df = df.drop('^_^', axis=1)
            # 修改换行的列，并生成最终dataframe
            pdf_df = pd.DataFrame(data=None, columns=self.add_data_title)
            for color in color_list:
                p_key = style_no + '^*^' + goods_name + '^*^' + fabric_no + '^*^' + color
                temp_version = 1
                if self.old_version.__contains__(p_key):
                    temp_version = int(self.old_version[p_key]) + 1
                    self.delete_key.append(p_key)
                # ['款号', '品名', '面料号', '用料品号', '供应商', '供应商品号', '数量', '规格', '物料说明', '物料颜色', '成衣颜色']
                pdf_df[self.add_data_title[0]] = temp_df[self.add_data_title[0]]
                pdf_df[self.add_data_title[1]] = temp_df[self.add_data_title[1]]
                pdf_df[self.add_data_title[2]] = temp_df[self.add_data_title[2]]
                pdf_df[self.add_data_title[3]] = temp_df['Material Description'].map(
                    lambda x: str(x).split('\n')[0])
                pdf_df[self.add_data_title[4]] = temp_df['Supplier'].map(
                    lambda x: str(x).replace('\n', ''))
                pdf_df[self.add_data_title[5]
                    ] = temp_df['Quality/Supplier#'].map(lambda x: str(x).replace('\n', ''))
                pdf_df[self.add_data_title[6]] = temp_df['Qty'].map(
                    lambda x: str(x).replace('\n', ''))
                pdf_df[self.add_data_title[7]] = temp_df['Size'].map(
                    lambda x: str(x).replace('\n', ''))
                pdf_df[self.add_data_title[8]] = temp_df['Placement'].map(
                    lambda x: str(x).replace('\n', ''))
                pdf_df[self.add_data_title[9]] = temp_df[color].map(
                    lambda x: str(x).replace('\n', '').replace('-', ''))
                pdf_df[self.add_data_title[10]] = color
                pdf_df[self.add_data_title[11]] = temp_version
                self.table_data = self.table_data.append(pdf_df, ignore_index=True)

    # 获取一个字符串中两个字母中间的值(one为None时从第一位取, two为None时取到最后)
    def get_value_two_word(self, txt_str, one, two):
        if one == None:
            return txt_str[:txt_str.find(two)]
        if two == None:
            return txt_str[txt_str.find(one) + len(one):]
        return txt_str[txt_str.find(one) + len(one):txt_str.find(two)]

    def select_trim_old_value(self):
        # 建立连接并获取PO数据
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        strCol = ",".join(str(i) for i in self.add_data_title)
        select_sql = 'select ' + strCol + ' from D_3DepTrimPdf'
        cursor.execute(select_sql)
        row = cursor.fetchall()
        self.old_all_data = pd.DataFrame(
            data=list(row), columns=self.add_data_title)
        cursor.close()
        conn.close()

    def update_db(self):
        dbCol = self.add_data_title[:]
        dbCol.append('CreateDate')
        # 建立连接并获取cursor
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        # 组装删除的值
        # del_tuple = tuple(self.delete_item)
        # # 删除已经存在的文件
        # delSql = 'delete from D_3DepTrimPdf where 款号 = (%s) and 品名 = (%s) and 面料号 = (%s) and 成衣颜色 = (%s) '
        # cursor.executemany(delSql, del_tuple)
        # 组装插入的值
        insertValue = []
        for tabVal in self.table_value:
            insertValue += tabVal
        insertSql = 'INSERT INTO D_3DepTrimPdf (' + (
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


def gui_start():
    VAS = VAS_GUI()
    VAS.get_files()


if __name__ == '__main__':
    gui_start()
