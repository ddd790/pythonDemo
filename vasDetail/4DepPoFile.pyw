import os
import shutil
import re
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import tkinter.messagebox as tmessage
import pandas as pd
import datetime
import pymssql
from dateutil import parser
import time


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
        # 设置标题
        self.init_window_name.title('四部文件操作工具！')
        # 设置窗口大小
        self.init_window_name.geometry('400x300')
        # tab页
        tab = ttk.Notebook(self.init_window_name, height=300, width=380)
        # po
        poFrame = Frame(tab)
        self.po_form_frame(poFrame)
        tab.add(poFrame, text="四部PO文件读取")
        tab.pack()

    def po_form_frame(self, poFrame):
        # 客户列表
        self.costmer_option = ["NEXT", "SLATERS", "DEVRED", "BS", "ITX"]
        # 标签
        self.type_label = Label(poFrame, text="客户：")
        self.type_label.grid(sticky=W, row=1, column=1)

        # radiobox 客户
        self.radio_val = IntVar()
        self.radio_val_1 = Radiobutton(
            poFrame, text=self.costmer_option[0], variable=self.radio_val, value=0).grid(sticky=W, row=2, column=1)
        self.radio_val_2 = Radiobutton(
            poFrame, text=self.costmer_option[1], variable=self.radio_val, value=1).grid(sticky=W, row=2, column=2)
        self.radio_val_3 = Radiobutton(
            poFrame, text=self.costmer_option[2], variable=self.radio_val, value=2).grid(sticky=W, row=3, column=1)
        self.radio_val_4 = Radiobutton(
            poFrame, text=self.costmer_option[3], variable=self.radio_val, value=3).grid(sticky=W, row=3, column=2)
        self.radio_val_5 = Radiobutton(
            poFrame, text=self.costmer_option[4], variable=self.radio_val, value=4).grid(sticky=W, row=4, column=1)

        # 标签
        self.type_label = Label(poFrame, text="操作类型：")
        self.type_label.grid(sticky=W, row=5, column=1)
        # 操作列表
        self.option_type = ["新PO", "修改PO"]
        self.radio_type = IntVar()
        self.radio_type_1 = Radiobutton(
            poFrame, text=self.option_type[0], variable=self.radio_type, value=0).grid(sticky=W, row=6, column=1)
        # self.radio_type_2 = Radiobutton(
        #     poFrame, text=self.option_type[1], variable=self.radio_type, value=1).grid(sticky=W, row=6, column=2)

        # 按钮
        self.commit_button = Button(poFrame, text="点击读取PO",
                                    bg="lightblue", width=18, command=self.commit_form)
        self.commit_button.grid(sticky=W, row=7, column=1)

        # 显示文字框
        self.file_show_label = Label(
            poFrame, text="请在D盘建立【4DepPo】文件夹，放入文件后点击按钮。", wraplength=400)
        self.file_show_label.grid(sticky=W, row=8, column=1, columnspan=10)
        self.file_show_label = Label(
            poFrame, text="※ 注意选择好【客户】和【操作类型】。", wraplength=400)
        self.file_show_label.config(fg="red")
        self.file_show_label.grid(sticky=W, row=9, column=1, columnspan=10)

    def commit_form(self):
        # 追加的dataFrame的title
        self.add_data_title = ['TYPE', 'PO号', '款号', '英文款名', '订单数量', '客人船期',
                               '目的港', '贸易方式', '走货方式', '商标', 'version', '面料颜色', '结汇币种', '季节号']
        self.add_data_title_size = ['TYPE', 'PO号', '款号', '英文款名', '订单数量', '客人船期',
                                    '目的港', '贸易方式', '走货方式', '商标', 'version', '面料颜色', '结汇币种', '季节号', 'option', 'size', 'quantity']
        # 数字类型的字段
        self.number_item = ['订单数量', 'quantity', 'version']
        # 录入DB的POdataframe
        self.table_data = pd.DataFrame(
            data=None, columns=self.add_data_title)
        # next客户的size
        self.table_data_next_size = pd.DataFrame(
            data=None, columns=self.add_data_title_size)
        self.table_value = []
        # 删除列表
        self.delete_key = []
        # 字段排除列表
        self.delete_str_key = ['NEXT', 'Product', 'Contact', 'Assign',
                               'Item', 'See', 'Terms', 'Refer', 'Manuals', 'Refurb', 'Description',
                               'IMPORTANT', ')', 'Type', 'Fabric', 'Trans', 'Contract', 'Del', 'Ex-Fact',
                               'Total', 'Supplier', 'Booking']
        # NEXT订单的po号集合，用于删除重复的新建记录
        self.nextPoList = []
        # 循环文件，处理合并
        for lroot, ldirs, lfiles in os.walk(self.local_po_file):
            if len(lfiles) == 0:
                tmessage.showerror('错误', '没有找到任何文件！')
                return
            for lfile in lfiles:
                ctime = parser.parse(time.ctime(os.path.getctime(
                    os.path.join(lroot, lfile))))
                create_time = ctime.strftime('%Y-%m-%d %H:%M:%S')
                self.file_to_dataframe(os.path.join(lroot, lfile), str(
                    lfile).split('.')[0], self.radio_val.get(), create_time)
        # 客户("NEXT", "SLATERS", "DEVRED", "BS", "ITX")
        self.table_data['客户'] = str(
            self.costmer_option[self.radio_val.get()])
        self.table_data['CreateDate'] = str(
            datetime.datetime.now()).split('.')[0]
        # NEXT size的dataframe
        self.table_data_next_size['CreateDate'] = str(
            datetime.datetime.now()).split('.')[0]
        self.table_value = []
        self.table_value.append([tuple(row)
                                for row in self.table_data.values])
        # NEXT的 size的dataframe
        self.table_value_next_size = []
        self.table_value_next_size.append(
            [tuple(row) for row in self.table_data_next_size.values])
        # 更新PO表
        self.update_db()
        if self.radio_val.get() == 0:
            # 更新NEXT的size表(0为新增，1为更新)
            self.update_next_size_db(self.radio_type.get())
        try:
            tmessage.showinfo('成功', '恭喜操作成功，请到勤哲系统中查看结果吧！')
        except:
            tmessage.showerror('错误', '人生苦短,程序出错了,请联系信息部孙适老师！')

    def file_to_dataframe(self, io, lfile, radioType, fileDate):
        # NEXT客户
        if radioType == 0:
            if self.radio_type.get() == 0:
                self.excel_to_dataframe_next_add(io, fileDate)
            else:
                self.excel_to_dataframe_next_update(io, fileDate)

    # NEXT客户的新增操作
    def excel_to_dataframe_next_add(self, io, fileDate):
        # 读取文件
        excel_header = []
        excelData = pd.read_excel(
            io, header=None, keep_default_na=False)
        excelCol = excelData.shape[1] + 1
        for h_idx in range(1, excelCol):
            excel_header.append('列' + str(h_idx))
        df = pd.DataFrame(excelData.values, columns=excel_header)
        df.dropna(axis=0, how='all')
        # 款号
        temp_style_no_list = self.check_str_key(df['列1'])
        style_no_list = []
        # 英文款名和type
        en_name_type_list = self.check_str_key(df['列2'])
        en_name_list = []
        type_list = []
        style_idx = -1
        temp_en_name = ''
        for en_name_type_val in en_name_type_list:
            if len(en_name_type_val) > 3:
                style_idx = style_idx + 1
                temp_en_name = en_name_type_val
            else:
                en_name_list.append(temp_en_name)
                style_no_list.append(temp_style_no_list[style_idx])
                type_list.append(en_name_type_val)
                # 走货方式
        trans_list = self.check_str_key(df['列5'])
        # PO号
        contract_no_list = self.check_str_key(df['列6'])
        contract_del_list = self.check_str_key(df['列7'])
        po_list = []
        for po_idx in range(0, len(contract_no_list)):
            po_list.append(
                str(contract_no_list[po_idx]) + '-' + str(contract_del_list[po_idx]))
        # 客人船期
        shipping_list = []
        ex_fact_list = self.check_str_key(df['列8'])
        c_year = str(fileDate).split('-')[0]
        c_month = str(fileDate).split('-')[1]
        for ex_fact in ex_fact_list:
            year = c_year
            if int(str(ex_fact).split('/')[1]) <= int(c_month):
                year = int(c_year) + 1
            shipping_list.append(
                str(year) + '-' + str(ex_fact).split('/')[1] + '-' + str(ex_fact).split('/')[0])
        # 订单数量
        qty_list = self.check_str_key(df['列9'])
        # 颜色，尺码，配码数量的list
        next_size_num_list = self.get_color_num_next(df, excelCol)
        # 组装数据进行存储
        po_df = pd.DataFrame(data=None, columns=self.add_data_title)
        po_df[self.add_data_title[0]] = type_list
        po_df[self.add_data_title[1]] = po_list
        po_df[self.add_data_title[2]] = style_no_list
        po_df[self.add_data_title[3]] = en_name_list
        po_df[self.add_data_title[4]] = qty_list
        po_df[self.add_data_title[5]] = shipping_list
        po_df[self.add_data_title[6]] = ''
        po_df[self.add_data_title[7]] = ''
        po_df[self.add_data_title[8]] = trans_list
        po_df[self.add_data_title[9]] = ''
        po_df[self.add_data_title[10]] = 1
        po_df[self.add_data_title[11]] = ''
        po_df[self.add_data_title[12]] = ''
        po_df[self.add_data_title[13]] = ''
        self.table_data = self.table_data.append(po_df, ignore_index=True)
        # NEXT的size数据的组装
        for n_idx in range(len(po_list)):
            po_df_size = pd.DataFrame(
                data=None, columns=self.add_data_title_size)
            po_df_size[self.add_data_title_size[14]
                       ] = next_size_num_list[n_idx]['sizeNo']
            po_df_size[self.add_data_title_size[15]
                       ] = next_size_num_list[n_idx]['size']
            po_df_size[self.add_data_title_size[16]
                       ] = next_size_num_list[n_idx]['num']
            po_df_size[self.add_data_title_size[0]] = type_list[n_idx]
            po_df_size[self.add_data_title_size[1]] = po_list[n_idx]
            po_df_size[self.add_data_title_size[2]] = style_no_list[n_idx]
            po_df_size[self.add_data_title_size[3]] = en_name_list[n_idx]
            po_df_size[self.add_data_title_size[4]] = qty_list[n_idx]
            po_df_size[self.add_data_title_size[5]] = shipping_list[n_idx]
            po_df_size[self.add_data_title_size[6]] = ''
            po_df_size[self.add_data_title_size[7]] = ''
            po_df_size[self.add_data_title_size[8]] = trans_list[n_idx]
            po_df_size[self.add_data_title_size[9]] = ''
            po_df_size[self.add_data_title_size[10]] = 1
            po_df_size[self.add_data_title_size[11]] = ''
            po_df_size[self.add_data_title_size[12]] = ''
            po_df_size[self.add_data_title_size[13]] = ''
            self.table_data_next_size = self.table_data_next_size.append(
                po_df_size, ignore_index=True)
        self.nextPoList.extend(po_list)

    # NEXT客户的修改操作
    def excel_to_dataframe_next_update(self, io, fileDate):
        excelData = pd.read_excel(
            io, header=None, keep_default_na=False)
        # 整理表格数据
        arr_excel_val = []
        # 跳过Item
        jump_flag = False
        for e_v in excelData.values:
            if str(e_v[0]) == 'Item':
                jump_flag = True
                continue
            if jump_flag and (str(e_v[0]) != '' or str(e_v[1]) != ''):
                arr_excel_val.append(e_v.tolist())
        # 组装PO信息和size的信息
        # 款号
        style_no_list = []
        # 英文款名和type
        en_name_list = []
        type_list = []
        # 走货方式
        trans_list = []
        # PO号
        po_list = []
        temp_style_no = ''
        temp_en_name_list = ''
        for po_info_val in arr_excel_val:
            if str(po_info_val[0]) != '':
                temp_style_no = po_info_val[0]
                temp_en_name_list = po_info_val[1]
        print(arr_excel_val)

    # 删除带有关键字的字段
    def check_str_key(self, df_val):
        res_list = []
        for val in df_val:
            val_add_flag = True
            for del_str in self.delete_str_key:
                if str(val).strip() == '' or del_str in str(val):
                    val_add_flag = False
                    break
            if val_add_flag:
                res_list.append(str(val).replace(':', ''))
        return res_list

    # 提取NEXT配色配码信息
    def get_color_num_next(self, next_df, excelCol):
        # 配色配码的列
        cols = []
        for h_idx in range(10, excelCol - 2):
            cols.append(h_idx)
        res_df = next_df[next_df.columns[cols]]
        # 按照Click here.....（列12）进行分组
        color_size_num_list = []
        color_list = []
        size_list = []
        num_list = []
        step_idx = 1
        for res_idx in range(len(res_df)):
            if 'Assign' in str(res_df['列14'][res_idx]):
                step_idx = 1
                if len(color_list) > 0:
                    color_size_num_list.append(
                        self.set_color_size_num_df(color_list, size_list, num_list))
                color_list = []
                size_list = []
                num_list = []
                continue
            if str(res_df['列11'][res_idx]) != '':
                res_df_list = res_df.iloc[res_idx].values.tolist()
                res_df_list = [i for i in res_df_list if i != '']
                if step_idx > 3:
                    step_idx = 1
                if step_idx == 1:
                    color_list.extend(res_df_list)
                elif step_idx == 2:
                    size_list.extend(res_df_list)
                elif step_idx == 3:
                    num_list.extend(res_df_list)
                step_idx = step_idx + 1
            # 最后一组需要追加上
            if res_idx == len(res_df) - 1:
                # 配色配码的dataframe
                color_size_num_list.append(
                    self.set_color_size_num_df(color_list, size_list, num_list))
        return color_size_num_list

    # 设置NEXT的配色配码
    def set_color_size_num_df(self, color_list, size_list, num_list):
        color_size_num_df = pd.DataFrame(
            columns=['sizeNo', 'size', 'num'])
        color_size_num_df['sizeNo'] = color_list
        color_size_num_df['size'] = size_list
        color_size_num_df['num'] = num_list
        return color_size_num_df

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

    def update_db(self):
        dbCol = self.add_data_title[:]
        dbCol.append('客户')
        dbCol.append('CreateDate')
        print(dbCol)
        # 建立连接并获取cursor
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        if len(self.nextPoList) > 0:
            # 组装删除的值
            del_tuple = tuple(self.nextPoList)
            # 删除已经存在的文件
            delSql = 'delete from D_4DepPoInfo where version = 1 and PO号 = (%s)'
            cursor.executemany(delSql, del_tuple)
        # 组装插入的值
        insertValue = []
        for tabVal in self.table_value:
            insertValue += tabVal
        insertSql = 'INSERT INTO D_4DepPoInfo VALUES ('
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

    def update_next_size_db(self, insertType):
        dbCol = self.add_data_title_size[:]
        dbCol.append('CreateDate')
        print(dbCol)
        # 建立连接并获取cursor
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        if len(self.nextPoList) > 0 and insertType == 0:
            # 组装删除的值
            del_tuple = tuple(self.nextPoList)
            # 删除已经存在的文件
            delSql = 'delete from D_4DepNEXTSize where version = 1 and PO号 = (%s)'
            cursor.executemany(delSql, del_tuple)
        # 组装插入的值
        insertValue = []
        for tabVal in self.table_value_next_size:
            insertValue += tabVal
        insertSql = ''
        if insertType == 0:
            insertSql = 'INSERT INTO D_4DepNEXTSize VALUES ('
            for colVal in dbCol:
                if colVal == 'CreateDate':
                    insertSql += '%s'
                elif colVal in self.number_item:
                    insertSql += '%d, '
                else:
                    insertSql += '%s, '
            insertSql += ')'
        print(insertSql)
        cursor.executemany(insertSql, insertValue)
        conn.commit()
        conn.close()


def gui_start():
    init_window = Tk()  # 实例化出一个父窗口
    VAS = VAS_GUI(init_window)
    VAS.set_init_window()  # 设置根窗口默认属性

    init_window.mainloop()  # 父窗口进入事件循环，可以理解为保持窗口运行，否则界面不展示


if __name__ == '__main__':
    gui_start()
