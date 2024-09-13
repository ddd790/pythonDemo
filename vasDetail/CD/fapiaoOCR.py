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
        # self.add_data_title = ['发票号码', '开票日期','购买方名称', '购买方纳税人识别号', '销售方名称', '销售方纳税人识别号', '项目名称', '规格型号', '单位', '数量', '单价', '金额',  
        #                      '税率', '税额', '价税合计', '备注', '文件名', '文件类型', '创建时间']
        self.add_data_title = ['InvoiceNo', 'InvoiceDate', 'BuyName', 'BuyNo', 'SellName', 'SellNo', 'Name', 'Size', 'Unit', 'Number', 'UnitPrice', 'Price',
                               'Rate', 'Tax', 'TotalPrice', 'Remarks', 'FileName', 'FileType', 'CreateDate']
        # self.add_data_title = ['test']
        # 数字类型的字段
        self.number_item = ['Number', 'UnitPrice', 'Price', 'Rate', 'Tax', 'TotalPrice',]
        # 服务器发票文件路径
        networked_directory = r'\\192.168.0.3\18-电子发票'
        # self.local_list_file = 'd:\\fapiaoTest'
        self.local_list_file = 'd:\\fapiao'
        self.local_list_file_j = 'd:\\fapiao\加工费和成衣'
        self.local_list_file_w = 'd:\\fapiao\物料发票'
        # 删除目录内文件
        if os.path.exists(self.local_list_file_j):
            shutil.rmtree(self.local_list_file_j, onerror=self.readonly_handler)
        os.mkdir(self.local_list_file_j)
        if os.path.exists(self.local_list_file_w):
            shutil.rmtree(self.local_list_file_w, onerror=self.readonly_handler)
        os.mkdir(self.local_list_file_w)
        # copy服务器的发票文件到本地
        self.table_value = []
        # self.table_data = pd.DataFrame(data=None, columns=['test'])
        for root, dirs, files in os.walk(networked_directory):
            for file in files:
                # print(str(root) + str(file))
                # self.table_data.append([str(str(root) + str(file)).replace('\\', '~')])
                # self.table_data.append(pd.DataFrame([tmp_str],columns=['test']), )
                if (str(file).__contains__('.pdf') or str(file).__contains__('.PDF')) and not str(file).__contains__('~'):
                    # tmp_str = str(str(root) + '*' + str(file)).replace('\\', '*')
                    # self.table_data = pd.concat([self.table_data, pd.DataFrame([tmp_str], columns=['test'])]).reset_index(drop=True)
                    if root.__contains__('加工费和成衣'):
                        shutil.copy2(os.path.join(root, file), self.local_list_file_j)
                    elif root.__contains__('物料发票'):
                        shutil.copy2(os.path.join(root, file), self.local_list_file_w)
        # 最终dataframe
        self.table_data = pd.DataFrame(data=None, columns=self.add_data_title)
        # # 查询数据库已经存在的发票号码
        self.select_invoice_old_value()
        self.table_value = []
        # 循环文件，处理合并
        for lroot, ldirs, lfiles in os.walk(self.local_list_file):
            for lfile in lfiles:
                # print(lfile)
                # 发票文件类型,目前只有【物料发票】和【加工费和成衣】
                file_type = '物料发票'
                if lroot.__contains__('加工费和成衣'):
                    file_type = '加工费和成衣'
                self.file_to_dataframe(os.path.join(lroot, lfile), str(lfile).split('.')[0], file_type)
        self.table_value.append([tuple(row) for row in self.table_data.values])
        # print(self.table_value)
        # 更新数据库
        self.update_db()
        # self.update_db_test()
        # 回车退出
        print('------------------------------------------------------------')
        print('已经完成操作！' + str(datetime.datetime.now()).split('.')[0])
        input('按回车退出 ')

    def file_to_dataframe(self, io, lfile, file_type):
        pdf_df = pd.DataFrame(data=None, columns=self.add_data_title)
        pdf = pdfplumber.open(io)
        # ['发票号码', '开票日期','购买方名称', '购买方纳税人识别号', '销售方名称', '销售方纳税人识别号', '项目名称', '规格型号', '单位', '数量', '单价', '金额', '税率', '税额', '价税合计', '备注']
        invoice_no = ''
        invoice_date = ''
        buy_name = ''
        buy_no = ''
        sell_name = ''
        sell_no = ''
        total_price = 0
        remarks = ''
        # 打开电子发票的PDF文件  
        for page in pdf.pages:
            # 提取第一页的文本内容  
            text = page.extract_text()
            invoice_no = self.get_value_two_word(text, '发票号码：', '开票日期：').strip().replace('\n', '')[:20]
            invoice_date = self.get_value_two_word(text, '开票日期：', None)[:11].replace('年', '-').replace('月', '-').replace('日', '').replace(' ', '')
            if invoice_no.__contains__('年'):
                invoice_no = self.get_value_two_word(text, '电子发票（增值税专用发票）', '发票号码：').strip().replace('\n', '')[:20]
                invoice_date = self.get_value_two_word(text, '发票号码：', '开票日期：').strip().replace('\n', '')[:11].replace('年', '-').replace('月', '-').replace('日', '').replace(' ', '')
            # 项目明细数据
            text = text.replace('税   额', '税  额').replace('税  额', '税 额')
            detail_info = self.get_value_two_word(text, '税 额\n', '合 计').strip()
            detali_info_list = detail_info.split('\n')
            # ['项目名称', '规格型号', '单位', '数量', '单价', '金额', '税率', '税额']的集合
            name_list = []
            size_list = []
            unit_list = []
            number_list = []
            unit_price_list = []
            price_list = []
            rate_list = []
            tax_list = []
            for item in detali_info_list:
                # print(item)
                item = item.replace('  ', ' ')
                detail_item_list = item.split(' ')
                temp_list = [item for item in detail_item_list if item != '']
                # 不满足条件的单行数据跳过
                if len(temp_list) < 6:
                    continue
                name_list.append(self.deleteByStar(detail_item_list[0]))
                tax_list.append(detail_item_list[-1])
                rate_list.append(detail_item_list[-2].replace('%', ''))
                price_list.append(detail_item_list[-3])
                unit_price_list.append(detail_item_list[-4])
                number_list.append(detail_item_list[-5])
                if detail_item_list[-6].__contains__('*') or len(detail_item_list[-6]) > 5:
                    unit_list.append('')
                else:
                    unit_list.append(detail_item_list[-6])
                size_list.append(self.get_value_two_word(item, detail_item_list[0], detail_item_list[-6]).strip())
            # 读取不到表格的情况
            if len(page.extract_tables()) == 0:
                buy_name = self.get_value_two_word(text, '购 名称：', '销 名称：').strip().replace('\n', '')
                sell_name = self.get_value_two_word(text,  '销 名称：', None).split(' ')[0]
                no_list = text.split('社会信用代码/纳税人识别号')
                buy_no = no_list[1].split(' ')[0]
                sell_no = no_list[2].split(' ')[0]
                total_price = text.split('（小写）¥')[1].replace('\n', ' ').split(' ')[0]
            else:
                for table in page.extract_tables():
                    # 购买双方信息
                    one_table = table[0]
                    # 另外一种发票，读取的是第二个表格的内容
                    # print(one_table)
                    if one_table[0] != None and not one_table[0].__contains__('购'):
                        continue
                    buy_name = self.get_value_two_word(one_table[1].split('\n')[0],  '名称：', None).strip().replace('\n', '')
                    buy_no = self.get_value_two_word(one_table[1],  '识别号', None).strip().replace('\n', '')
                    sell_name = self.get_value_two_word(one_table[-1].split('\n')[0],  '名称：', None).strip().replace('\n', '')
                    sell_no = self.get_value_two_word(one_table[-1],  '识别号', None).strip().replace('\n', '')
                    total_price = self.get_value_two_word(table[2][2], '¥', None).strip().replace('\n', '')
                    remarks = str(table[3][1].strip())
        # ['发票号码', '开票日期','购买方名称', '购买方纳税人识别号', '销售方名称', '销售方纳税人识别号', '项目名称', '规格型号', '单位', '数量', '单价', '金额', '税率', '税额', '价税合计', '备注']
        if invoice_no not in self.old_invoice_no_list:
            pdf_df.loc[:, self.add_data_title[6]] = name_list
            pdf_df.loc[:, self.add_data_title[7]] = size_list
            pdf_df.loc[:, self.add_data_title[8]] = unit_list
            pdf_df.loc[:, self.add_data_title[9]] = number_list
            pdf_df.loc[:, self.add_data_title[10]] = unit_price_list
            pdf_df.loc[:, self.add_data_title[11]] = price_list
            pdf_df.loc[:, self.add_data_title[12]] = rate_list
            pdf_df.loc[:, self.add_data_title[13]] = tax_list
            pdf_df.loc[:, self.add_data_title[14]] = total_price
            pdf_df.loc[:, self.add_data_title[15]] = remarks
            pdf_df.loc[:, self.add_data_title[16]] = lfile
            pdf_df.loc[:, self.add_data_title[17]] = file_type
            pdf_df.loc[:, self.add_data_title[18]] = str(datetime.datetime.now()).split('.')[0]
            pdf_df.loc[:, self.add_data_title[0]] = invoice_no
            pdf_df.loc[:, self.add_data_title[1]] = invoice_date
            pdf_df.loc[:, self.add_data_title[2]] = buy_name.replace(':', '').replace('：', '')
            pdf_df.loc[:, self.add_data_title[3]] = buy_no.replace(':', '').replace('：', '')
            pdf_df.loc[:, self.add_data_title[4]] = sell_name.replace(':', '').replace('：', '')
            pdf_df.loc[:, self.add_data_title[5]] = sell_no.replace(':', '').replace('：', '')
            self.table_data = self.table_data.append(pdf_df, ignore_index=True)
        pdf.close()

    # 获取一个字符串中两个字母中间的值(one为None时从第一位取, two为None时取到最后)
    def get_value_two_word(self, txt_str, one, two):
        if one == None:
            return txt_str[:txt_str.find(two)]
        if two == None:
            return txt_str[txt_str.find(one) + len(one):]
        return txt_str[txt_str.find(one) + len(one):txt_str.find(two)]

    # 删除文字中两个*之间的文字，包含两个*
    def deleteByStar(self, text):
        pattern = r"\*.*?\*"
        return re.sub(pattern, "", text)   

    def update_db(self):
        dbCol = self.add_data_title[:]
        # 建立连接并获取cursor
        conn = pymssql.connect(self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        # 组装插入的值
        insertValue = []
        for tabVal in self.table_value:
            insertValue += tabVal
        insertSql = 'INSERT INTO D_InvoiceOCR (' + (",".join(str(i) for i in dbCol)) + ') VALUES ('
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

    def update_db_test(self):
        dbCol = self.add_data_title[:]
        # 建立连接并获取cursor
        conn = pymssql.connect(self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        # 组装插入的值
        insertValue = []
        for tabVal in self.table_value:
            insertValue += tabVal
        insertSql = 'INSERT INTO test (' + (",".join(str(i) for i in dbCol)) + ') VALUES ('
        for colVal in dbCol:
            insertSql += '%s'
        insertSql += ')'
        # print(insertValue)
        cursor.executemany(insertSql, insertValue)
        conn.commit()
        conn.close()

    # 查询数据库已经存在的发票号码
    def select_invoice_old_value(self):
        # 建立连接并获取PO数据
        conn = pymssql.connect(self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        select_sql = 'select distinct InvoiceNo from D_InvoiceOCR'
        cursor.execute(select_sql)
        row = cursor.fetchall()
        old_invoice_no_list = pd.DataFrame(data=list(row), columns=['InvoiceNo'])
        self.old_invoice_no_list = list(set(old_invoice_no_list['InvoiceNo']))
        cursor.close()
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
