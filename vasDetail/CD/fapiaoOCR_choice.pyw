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
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import filedialog, messagebox


class VAS_GUI():
    def __init__(self, init_window_name):
        self.init_window_name = init_window_name

    def set_init_window(self):
        # 设置标题
        self.init_window_name.title('发票文件读取工具！')
        # 设置窗口大小
        self.init_window_name.geometry('400x300')
        # tab页
        tab = ttk.Notebook(self.init_window_name, height=300, width=380)
        # po
        poFrame = Frame(tab)
        self.po_form_frame(poFrame)
        tab.add(poFrame, text="客户文件读取")
        tab.pack()

    def po_form_frame(self, poFrame):
        # 显示文字框
        self.file_show_label = Text(poFrame, width=50, height=10)
        self.file_show_label.grid(sticky=W, row=1, column=1, columnspan=10)

        # 按钮
        self.commit_button = Button(poFrame, text="选择上传文件", bg="lightblue", width=18, command=self.open_file)
        self.commit_button.grid(sticky=W, row=7, column=1)
        self.commit_button = Button(poFrame, text="点击上传数据", bg="lightblue", width=18, command=self.commit_form)
        self.commit_button.grid(sticky=W, row=8, column=1)

    def open_file(self):
        self.file_path = filedialog.askopenfilenames(title=u'选择文件', initialdir=(os.path.expanduser(r'\\192.168.0.3\18-电子发票')))
        self.file_show_label.delete(1.0, tk.END)
        for path in self.file_path:
            if '//192.168.0.3/18-电子发票' not in path:
                messagebox.showerror("路径错误", f"路径错误: 请选择共享服务器下的【18-电子发票】文件夹下的文件！")
                return
            self.file_show_label.insert('insert', str(path.split('/')[-1]) + '\n')

    # 电子发票数据提取
    def commit_form(self):
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
        # 查询数据库已经存在的文件名
        self.select_fileName_old_value()
        # copy服务器的发票文件到本地
        for path in self.file_path:
            if not str(path).__contains__('~') and str(path.split('/')[-1]).replace('.pdf', '') not in self.old_fileName_no_list:
                if path.__contains__('加工费和成衣'):
                    shutil.copy2(path, self.local_list_file_j)
                elif path.__contains__('物料发票'):
                    shutil.copy2(path, self.local_list_file_w)
        # 最终dataframe
        self.table_data = pd.DataFrame(data=None, columns=self.add_data_title)
        # 查询数据库已经存在的发票号码
        self.select_invoice_old_value()
        self.table_value = []
        # 循环文件，处理合并
        for lroot, ldirs, lfiles in os.walk(self.local_list_file):
            for lfile in lfiles:
                # 发票文件类型,目前只有【物料发票】和【加工费和成衣】
                file_type = '物料发票'
                if lroot.__contains__('加工费和成衣'):
                    file_type = '加工费和成衣'
                self.file_to_dataframe(os.path.join(lroot, lfile), str(lfile).split('.')[0], file_type)
        self.table_value.append([tuple(row) for row in self.table_data.values])
        for row in self.table_data.itertuples(index=False):
            # 判断是否为数字，不是数字则输出文件名
            if not self.is_number(row.Number):
                self.file_show_label.insert('insert', '文件名：' + row.FileName + '，发票号码：' + row.InvoiceNo + '，项目明细数据有误，请检查！' + '\n')
                continue
        # 更新数据库
        try:
            self.update_db()
            messagebox.showinfo("成功", f"已经提交成功了，请到勤哲里查看结果！")
        except:
            messagebox.showerror("出错啦", f"提交错误，请联系管理员！")

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
        # for page in pdf.pages:
        page = pdf.pages[0]
        # 提取第一页的文本内容  
        text = page.extract_text()
        # print(text)
        text = text.replace('发 票 号码：', '发票号码：').replace('发 票 号 码 ：', '发票号码：').replace('开票 日期：', '开票日期：').replace('电 子 发 票 （ 增 值 税 专 用 发 票 ）', '电子发票（增值税专用发票）')
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
            if len(temp_list) < 6 or detail_item_list[0].__contains__('项目名称'):
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
        # 如果table_data的InvoiceDate列长度不等于10，则将InvoiceDate列的值设置为当前时间
        if len(self.table_data['InvoiceDate'].values[0]) != 10:
            self.table_data['InvoiceDate'] = str(datetime.datetime.now()).split(' ')[0]
        # 去掉table_data的BuyName和SellName列中的空格
        self.table_data['BuyName'] = self.table_data['BuyName'].str.replace(' ', '')
        self.table_data['SellName'] = self.table_data['SellName'].str.replace(' ', '')
        # print(self.table_data)
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
    
    # 查询数据库已经存在的发票文件名
    def select_fileName_old_value(self):
        # 建立连接并获取PO数据
        conn = pymssql.connect(self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        select_sql = 'select distinct FileName from D_InvoiceOCR'
        cursor.execute(select_sql)
        row = cursor.fetchall()
        old_fileName_no_list = pd.DataFrame(data=list(row), columns=['FileName'])
        self.old_fileName_no_list = list(set(old_fileName_no_list['FileName']))
        cursor.close()
        conn.close()
    
    # 文件只读删除的解决
    def readonly_handler(self, func, path, exc_info):
        os.chmod(path, stat.S_IWRITE)
        func(path)
    
    def is_number(self, s):
        try:
            float(s)
            return True
        except ValueError:
            pass

        try:
            import unicodedata
            unicodedata.numeric(s)
            return True
        except (TypeError, ValueError):
            pass

        return False

def gui_start():
    init_window = Tk()  # 实例化出一个父窗口
    VAS = VAS_GUI(init_window)
    VAS.set_init_window()  # 设置根窗口默认属性
    init_window.mainloop()  # 父窗口进入事件循环，可以理解为保持窗口运行，否则界面不展示
    # VAS = VAS_GUI()
    # VAS.get_files()


if __name__ == '__main__':
    gui_start()
