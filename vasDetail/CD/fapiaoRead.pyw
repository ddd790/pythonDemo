import os
import re
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
import datetime
import pymssql


class VAS_GUI():
    def __init__(self, init_window_name):
        self.init_window_name = init_window_name

    def set_init_window(self):
        # sql服务器名
        self.serverName = '192.168.0.11'
        # 登陆用户名和密码
        self.userName = 'sa'
        self.passWord = 'jiangbin@007'
        # 数据库名
        self.dbName = 'ESApp1'
        self.local_po_file = 'd:\\4DepPo'
        self.file_path = ''
        # 追加的dataFrame的title
        self.add_data_title = ['InvoiceNo', 'InvoiceDate', 'BuyName', 'BuyNo', 'SellName', 'SellNo', 'Name', 'Size', 'Unit', 'Number', 'UnitPrice', 'Price',
                               'Rate', 'Tax', 'TotalPrice', 'Remarks', 'FileName', 'FileType', 'CreateDate']
        # 数字类型的字段
        self.number_item = ['Number', 'UnitPrice', 'Price', 'Rate', 'Tax', 'TotalPrice',]
        # 设置标题
        self.init_window_name.title('读取纸质发票工具！')
        # 设置窗口大小
        self.init_window_name.geometry('400x300')
        # tab页
        tab = ttk.Notebook(self.init_window_name, height=300, width=380)
        # po
        poFrame = Frame(tab)
        self.po_form_frame(poFrame)
        tab.add(poFrame, text="纸质发票读取")
        tab.pack()

    def po_form_frame(self, poFrame):
        # 最终dataframe
        self.table_data = pd.DataFrame(data=None, columns=self.add_data_title)
        self.file_path = ''
        # 显示文字框
        self.file_show_label = Text(poFrame, width=50, height=10)
        self.file_show_label.grid(sticky=W, row=1, column=1, columnspan=10)

        # 按钮
        self.commit_button = Button(poFrame, text="选择上传文件", bg="lightblue", width=18, command=self.open_file)
        self.commit_button.grid(sticky=W, row=7, column=1)
        self.commit_button = Button(poFrame, text="点击上传数据", bg="lightblue", width=18, command=self.commit_form)
        self.commit_button.grid(sticky=W, row=8, column=1)
        
    def open_file(self):
        self.file_path = filedialog.askopenfilename(title=u'选择文件', initialdir=(os.path.expanduser('D:/')))
        self.file_show_label.insert('insert', self.file_path + '\n')

    def commit_form(self):
        # 查询数据库已经存在的发票号码
        try:
            self.select_invoice_old_value()
            self.table_value = []
            # 读取文件，处理合并
            self.file_to_dataframe(self.file_path, str(self.file_path).split('.')[0])
            self.table_value.append([tuple(row) for row in self.table_data.values])
            self.file_show_label.insert('insert', '总计： ' + str(len(self.table_data)) + ' 条记录。\n')
            self.file_show_label.insert('insert', '数据处理中......' + str(datetime.datetime.now()).split('.')[0] + '\n')
            # 更新数据库
            self.update_db()
            self.file_show_label.insert('insert', '成功：数据处理完毕。时间：' + str(datetime.datetime.now()).split('.')[0] + '\n')
        except:
            self.file_show_label.insert('insert', '失败：请联系信息部。时间：' + str(datetime.datetime.now()).split('.')[0] + '\n')

    def file_to_dataframe(self, io, fileName):
        # 读取文件
        excelData = pd.read_excel(io, header=None, keep_default_na=False)
        # csv中的title
        formartExcelTitle = []
        # csv数据
        formatExcelvalue = excelData.values
        for item in excelData.values[0]:
            formartExcelTitle.append(item)
        df = pd.DataFrame(formatExcelvalue, columns=formartExcelTitle)
        # 去掉首行
        df.drop(index=0,inplace=True)
        df.reset_index(drop=True,inplace=True)
        df['发票号码'] = df['发票号码'].astype('str')
        # 发票号码不存在的，进行insert
        for idx, row in df.iterrows():
            if row['发票号码'] in self.old_invoice_no_list:
                df.drop(idx, inplace=True)
            else:
                df.at[idx, '货物或应税劳务名称'] = self.deleteByStar(row['货物或应税劳务名称'])
        # excel列（购方名称, 销方名称, 货物或应税劳务名称, 规格型号, 单位, 数量, 单价, 金额, 税额, 税率, 发票号码, 发票代码, 开票日期, 销方税号, 销方地址及电话, 销方开户行及账户, 备注）
        # 对应的DB列 ['InvoiceNo', 'InvoiceDate', 'BuyName', 'BuyNo', 'SellName', 'SellNo', 'Name', 'Size', 'Unit', 'Number', 'UnitPrice', 'Price',
        #                       'Rate', 'Tax', 'TotalPrice', 'Remarks', 'FileName', 'FileType', 'CreateDate']
        df_dict = {
            'InvoiceNo' : '发票号码',
            'InvoiceDate' : '开票日期',
            'BuyName' : '购方名称',
            'BuyNo' : '',
            'SellName' : '销方名称',
            'SellNo' : '销方税号',
            'Name' : '货物或应税劳务名称',
            'Size' : '规格型号',
            'Unit' : '单位',
            'Number' : '数量',
            'UnitPrice' : '单价',
            'Price' : '金额',
            'Rate' : '',
            'Tax' : '税额',
            'TotalPrice' : '金额',
            'Remarks' : '备注',
            'FileName' : '',
            'FileType' : '',
            'CreateDate' : ''
        }
        for item in self.add_data_title:
            if df_dict[item] != '':
                self.table_data.loc[:, item] = list(df[df_dict[item]])
        # 固定值
        self.table_data.loc[:, 'Rate'] = 13
        self.table_data.loc[:, 'FileName'] = ''
        self.table_data.loc[:, 'FileType'] = '物料发票'
        self.table_data.loc[:, 'CreateDate'] = str(datetime.datetime.now()).split('.')[0]
        self.table_data['TotalPrice'] = self.table_data['Price'] + self.table_data['Tax']
        # 添加税号
        self.table_data.loc[self.table_data['BuyName']=='大连景泰蓝服装有限公司', ['BuyNo']] = '91210202696038997B'
        self.table_data.loc[self.table_data['BuyName']=='大连瑞忠贸易有限公司', ['BuyNo']] = '91210202MA0XNN881C'
        print(self.table_data)

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

    # 删除文字中两个*之间的文字，包含两个*
    def deleteByStar(self, text):
        pattern = r"\*.*?\*"
        return re.sub(pattern, "", text)

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


if __name__ == '__main__':
    gui_start()
