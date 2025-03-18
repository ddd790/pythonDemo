import os
import pandas as pd
import pymssql
import datetime
import shutil
from tkinter import *
from tkinter import ttk
import tkinter.messagebox as tmessage

class VAS_GUI():
    def __init__(self, init_window_name):
        self.init_window_name = init_window_name

    def set_init_window(self):
        # print('数据操作进行中......' + str(datetime.datetime.now()).split('.')[0])
        # sql服务器名
        self.serverName = '192.168.0.11'
        # 登陆用户名和密码
        self.userName = 'sa'
        self.passWord = 'jiangbin@007'
        # 数据库名
        self.dbName = 'ESApp1'
        # 追加的dataFrame的title
        self.add_data_title = ['用料', '用料名称', '实测幅宽', '实测缩率', '报价单耗', 
                               '订料单耗', '客户', '款号','类别', '面料成份', '单耗负责人', '数量', '面料号', '下单日期', '数据提供', '发放日期', 'V_Key', 'FileName', 'Version']
        # self.digel_add_title = ['供应商品号']
        # 数字类型的字段
        self.number_item = ['Version', '报价单耗', '订料单耗']
        self.local_excel_detail_file = 'd:\\大货面辅料单耗用量表'
        self.bak_file =  'd:\\单耗用量表_bak'
        self.init_window_name.title('三部文件操作工具！')
        self.init_window_name.geometry('400x300')
        tab = ttk.Notebook(self.init_window_name, height=300, width=380)
        poFrame = Frame(tab)
        self.po_form_frame(poFrame)
        tab.add(poFrame, text="三部单耗文件读取")
        tab.pack()

    def po_form_frame(self, poFrame):
        self.costmer_option = ["DIGEL", "EXPRESS", "MAC"]
        self.type_label = Label(poFrame, text="客户：")
        self.type_label.grid(sticky=W, row=1, column=1)
        self.radio_val = IntVar()
        self.radio_val_1 = Radiobutton(poFrame, text=self.costmer_option[0], variable=self.radio_val, value=0).grid(sticky=W, row=2, column=1)
        self.radio_val_2 = Radiobutton(poFrame, text=self.costmer_option[1], variable=self.radio_val, value=1).grid(sticky=W, row=2, column=2)
        self.radio_val_3 = Radiobutton(poFrame, text=self.costmer_option[2], variable=self.radio_val, value=2).grid(sticky=W, row=2, column=3)
        # 按钮
        self.commit_button = Button(poFrame, text="读取单耗", bg="lightblue", width=16, command=self.get_files)
        self.commit_button.grid(sticky=W, row=8, column=1)
        # 显示文字框
        self.file_show_label = Label(poFrame, text="请在D盘建立【大货面辅料单耗用量表】文件夹，放入文件后点击按钮。", wraplength=400)
        self.file_show_label.grid(sticky=W, row=9, column=1, columnspan=10)
        self.file_show_label = Label(poFrame, text="※ 注意选择好【客户】。", wraplength=400)
        self.file_show_label.config(fg="red")
        self.file_show_label.grid(sticky=W, row=10, column=1, columnspan=10)
    # 批量获取服务器数据，读取大货面辅料单耗用量
    def get_files(self):
        # 查询已存在的记录
        self.select_old_value()
        # 根据V_Key，找到对应的version
        self.old_version = self.old_all_data.set_index("V_Key")["Version"].to_dict()
        # 最终dataframe
        self.table_data = pd.DataFrame(data=None, columns=self.add_data_title)
        self.table_value = []
        # 循环文件，处理合并
        try:
            for lroot, ldirs, lfiles in os.walk(self.local_excel_detail_file):
                if len(lfiles) == 0:
                    tmessage.showerror('错误', '没有任何文件呀！')
                    return None
                for lfile in lfiles:
                    if not str(lfile).__contains__('~'):
                        # print('文件名：' + str(lfile).split('.')[0])
                        self.file_to_dataframe(os.path.join(lroot, lfile), str(lfile).split('.')[0], self.radio_val.get())
            # 更新数据库
            self.update_db()
            # 删除目录内文件
            self.delete_files_in_folder(self.local_excel_detail_file)
            tmessage.showinfo('成功', '恭喜操作成功，请到勤哲系统中查看结果吧！')
        except:
            tmessage.showerror('出错啦', '错误在所难免，不要着急，请联系信息部人员进行解决！')

    def file_to_dataframe(self, io, lfile, radio_val):
        # 读取excel
        # 读取 Excel 文件中的所有工作表名称
        xls = pd.ExcelFile(io)
        sheet_names = xls.sheet_names
        read_sheet_name = sheet_names[0]
        # DIGEL获取最后一个工作表的名称
        if radio_val == 0:
            read_sheet_name = sheet_names[-1]
        df = pd.read_excel(io, sheet_name=read_sheet_name, nrows=100)
        # print(df)
        # '客户', '款号','类别'(也叫款式描述), '面料成份', '单耗负责人', '数量', '面料号', '下单日期', '数据提供', '发放日期'
        if radio_val == 0:
            customer, style, category, composition, user, qty, fNum, orderDate, dataSupport, sendDate = df.iloc[3, 3], df.iloc[3, 5], df.iloc[3, 8], df.iloc[3, 13], df.iloc[3, 16], df.iloc[4, 8], df.iloc[4, 13], df.iloc[4, 16], df.iloc[5, 1], df.iloc[5, 16]
        elif radio_val == 1:
            customer, style, category, composition, user, qty, fNum, orderDate, dataSupport, sendDate = df.iloc[3, 2], df.iloc[3, 4], df.iloc[3, 8], df.iloc[3, 13], df.iloc[3, 15], df.iloc[4, 4], df.iloc[4, 13], df.iloc[4, 15], '', df.iloc[5, 15]
        else:
            customer, style, category, composition, user, qty, fNum, orderDate, dataSupport, sendDate = df.iloc[2, 1], df.iloc[0, 1], df.iloc[1, 1], '', df.iloc[3, 2], 0, '', '1900-01-01', '', '1900-01-01'
        if pd.isna(composition) :
            composition = ''
        # v_key记录版本的key
        v_key = str(style) + '_' + str(category)
        # 取出用料中含有面料的行（10-21行中）
        if radio_val == 0:
            rows_of_interest = df.iloc[9:21, 0]
        elif radio_val == 1:
            rows_of_interest = df.iloc[9:21, 0]
        else:
            rows_of_interest = df.iloc[6:20, 1]
        rows_of_interest = rows_of_interest.fillna('')
        # 包含面料或者fabric的行的索引
        contains_shell_fabric = rows_of_interest.str.contains('fabric', case=False)
        # 判断contains_shell_fabric中是否有True
        if contains_shell_fabric.any():
            contains_fabric = rows_of_interest.str.contains('fabric', case=False)
        else:
            contains_fabric = rows_of_interest.str.contains('面料', case=False)
        row_indices_with_fabric = contains_fabric[contains_fabric].index.tolist()
        # 用料相关信息
        if radio_val == 0:
            table_data = pd.DataFrame(df.iloc[row_indices_with_fabric, [0,1,5,7,11,13,1,1,1,1,1,11,2,3]])
            # table_data.columns = self.add_data_title[0:13]
            new_array = self.add_data_title[0:13] + ["供应商品号"]
            table_data.columns = new_array
        elif radio_val == 1:
            table_data = pd.DataFrame(df.iloc[row_indices_with_fabric, [0,1,4,6,10,12]])
            table_data.columns = self.add_data_title[0:6]
        else:
            table_data = pd.DataFrame(df.iloc[row_indices_with_fabric, [1,2,8,10,15,17,15,15,15,5]])
            table_data.columns = self.add_data_title[0:10]
        table_data.fillna(0, inplace=True)
        # 订单相关信息
        table_data['实测幅宽'] = table_data['实测幅宽'].astype(str)
        table_data['实测缩率'] = table_data['实测缩率'].astype(str)
        table_data['报价单耗'] = table_data['报价单耗'].astype(str)
        table_data['订料单耗'] = table_data['订料单耗'].astype(str)
        table_data['客户'] = customer
        table_data['款号'] = style
        table_data['类别'] = category
        if radio_val != 0:
            table_data['面料成份'] = composition
        else:
            table_data['面料成份'] =table_data['面料成份'].astype(str)
        table_data['单耗负责人'] = user
        table_data['数量'] = qty
        if radio_val == 1:
            table_data['面料号'] = fNum
        elif radio_val == 2:
            table_data['面料号'] = table_data['用料名称']
        else:
            table_data['面料号'] =table_data['面料号'].astype(str) + '~' + table_data['供应商品号'].astype(str)
        table_data['下单日期'] = orderDate
        table_data['数据提供'] = dataSupport
        table_data['发放日期'] = sendDate
        table_data['V_Key'] = v_key
        table_data['FileName'] = str(lfile).split('.')[0]
        table_data['Version'] = 0 if self.old_version.get(v_key) == None else int(self.old_version.get(v_key)) + 1
        table_data['CreateDate'] = str(datetime.datetime.now()).split('.')[0]
        table_data.fillna('', inplace=True)
        if radio_val == 0:
            # 删除table_data的"供应商品号"列
            table_data.drop('供应商品号', axis=1, inplace=True)
        # 筛选用料名称不等于0的数据行
        filtered_rows = [tuple(row) for _, row in table_data[table_data['用料名称'] != 0].iterrows()]
        self.table_value.append(filtered_rows)

    def select_old_value(self):
        # 建立连接并获取PO数据
        conn = pymssql.connect(self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        strCol = ",".join(str(i) for i in self.add_data_title)
        select_sql = 'select ' + strCol + ' from D_三部大货面辅料单耗用量表'
        cursor.execute(select_sql)
        row = cursor.fetchall()
        self.old_all_data = pd.DataFrame(data=list(row), columns=self.add_data_title)
        cursor.close()
        conn.close()

    def update_db(self):
        dbCol = self.add_data_title[:]
        dbCol.append('CreateDate')
        # 建立连接并获取cursor
        conn = pymssql.connect(self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        # 组装插入的值
        insertValue = []
        for tabVal in self.table_value:
            insertValue += tabVal
        insertSql = 'INSERT INTO D_三部大货面辅料单耗用量表 (' + (",".join(str(i) for i in dbCol)) + ') VALUES ('
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
                    # 备份文件
                    if not os.path.exists(self.bak_file):
                        os.mkdir(self.bak_file)
                    shutil.copy(file_path, self.bak_file + '\\' + str(filename))
                    # 删除文件
                    os.remove(file_path)
                    # print(f"Deleted file: {file_path}")
                except Exception as e:
                    print(f"Failed to delete {file_path}. Error: {e}")
    
def gui_start():
    init_window = Tk()
    VAS = VAS_GUI(init_window)
    VAS.set_init_window()
    init_window.mainloop()

if __name__ == '__main__':
    gui_start()
