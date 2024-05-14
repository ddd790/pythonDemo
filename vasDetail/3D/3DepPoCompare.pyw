import os
import re
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
import datetime
import pymssql
import numpy as np

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
        self.file_path = ''
        # 追加的dataFrame的title
        self.add_data_title = ['客户定单部门', '到店铺月份', 'StoreTier', '款号', '款式描述', '面料号', '面料颜色', 'COLOR', '订单数量', '预计NDC月','预计NDC周', 
        '实际NDC月季节号', '实际NDC周', '供应商', '报客户工厂', 'FIBERDYED', 'MOQMCQBoth', 'UnitstoAvoidMinimumSurcharge', 'FOBBulkCost', 'BulkOffer', 
        'CostOfferComments', 'VendorTargetCost', '成衣海运RFQ', '成衣空运RFQ', '成衣海运价格', '成衣空运价格', '面料定料日', 'PUBLISHDATE', '计划PO日', 
        '报客户PO号', '面料出厂日', '计划裁剪日', '面料运输方式', '成衣运输方式', '成衣船期', '成衣实际船期', '成衣实际空运日', 'BookingDatesComments', 
        'MaterialID', 'BRYY', 'BR总码数', '面料单价_美元', '面料定单NFG号', '面料厂', '面料产地', '面料周期', 'CreateDate']
        # 数字类型的字段
        self.number_item = ['订单数量']
        # 设置标题
        self.init_window_name.title('三部PO文件读取工具！')
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
        self.table_value = []
        try:
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
        # 数字读取成文字的列
        dtype={'Spec Style' : str, 'Units to Avoid Minimum Surcharge' : str, 'Bulk Offer #' : str, 'Vendor Target Cost' : str, 'AIR OFFER #' : str, 'NFG' : str, 'FABRIC LT' : str}
        # 读取文件
        excelData = pd.read_excel(io, header=3, sheet_name=None, usecols='C:AW', dtype=dtype)
        # sheet名list
        sheetNameList = excelData.keys()
        nowDate = str(datetime.datetime.now()).split('.')[0]
        df = pd.DataFrame(data=None, columns=self.add_data_title)
        # 循环读取每个sheet数据
        for item in sheetNameList:
            formatExcelvalue = excelData[item].values
            # sheet页的dataframe
            df = pd.concat([df, pd.DataFrame(formatExcelvalue, columns=self.add_data_title)])
            df.loc[:, '客户定单部门'] = item
            df.loc[:, 'CreateDate'] = nowDate
            df.dropna(axis=0, subset = ['款号'], inplace=True)
            self.table_data = pd.concat([self.table_data, df])
        # 数字列转换成0
        self.table_data.replace(np.nan, '', inplace=True)
        for col in self.add_data_title:
            if col != 'CreateDate' and col != '订单数量':
                self.table_data[col] = self.table_data[col].astype(str)
        # print(self.table_data)

    def update_db(self):
        dbCol = self.add_data_title[:]
        # 建立连接并获取cursor
        conn = pymssql.connect(self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        cursor.execute('TRUNCATE TABLE D_3DepPOExcel')
        # 组装插入的值
        insertValue = []
        for tabVal in self.table_value:
            insertValue += tabVal
        insertSql = 'INSERT INTO D_3DepPOExcel (' + (",".join(str(i) for i in dbCol)) + ') VALUES ('
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
        # log插入
        insertSql = 'INSERT INTO D_3DepPOExcel_Log (' + (",".join(str(i) for i in dbCol)) + ') VALUES ('
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
