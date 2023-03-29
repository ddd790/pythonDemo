from cgi import print_arguments
import os
from tkinter import *
from tkinter import ttk
import tkinter.messagebox as tmessage
import pandas as pd
import datetime
import pymssql
from dateutil import parser
import time
import pdfplumber
from aip import AipOcr
import numpy as np


class VAS_GUI:

    def __init__(self, init_window_name):
        self.init_window_name = init_window_name

    def set_init_window(self):
        APP_ID = '25101742'
        API_KEY = 'Z5qy26GRDUdDKlBRHGT21XZt'
        SECRET_KEY = 'p6BCz0xxGXSTbDR3MfWAfViBRbFilaAu'
        self.client = AipOcr(APP_ID, API_KEY, SECRET_KEY)
        self.serverName = '192.168.0.11'
        self.userName = 'sa'
        self.passWord = 'jiangbin@007'
        self.dbName = 'ESApp1'
        self.local_po_file = 'd:\\4DepPo'
        self.init_window_name.title('四部文件操作工具！')
        self.init_window_name.geometry('400x300')
        tab = ttk.Notebook(self.init_window_name, height=300, width=380)
        poFrame = Frame(tab)
        self.po_form_frame(poFrame)
        tab.add(poFrame, text="四部PO文件读取")
        tab.pack()

    def po_form_frame(self, poFrame):
        self.costmer_option = ["NEXT", "SLATERS", "DEVRED", "BS", "ITX", "CELIO", "IZAC"]
        self.type_label = Label(poFrame, text="客户：")
        self.type_label.grid(sticky=W, row=1, column=1)
        self.radio_val = IntVar()
        self.radio_val_1 = Radiobutton(poFrame, text=self.costmer_option[0], variable=self.radio_val, value=0).grid(sticky=W, row=2, column=1)
        self.radio_val_2 = Radiobutton(poFrame, text=self.costmer_option[1], variable=self.radio_val, value=1).grid(sticky=W, row=2, column=2)
        self.radio_val_3 = Radiobutton(poFrame, text=self.costmer_option[2], variable=self.radio_val, value=2).grid(sticky=W, row=3, column=1)
        self.radio_val_4 = Radiobutton(poFrame, text=self.costmer_option[3], variable=self.radio_val, value=3).grid(sticky=W, row=3, column=2)
        self.radio_val_5 = Radiobutton(poFrame, text=self.costmer_option[4], variable=self.radio_val, value=4).grid(sticky=W, row=4, column=1)
        self.radio_val_6 = Radiobutton(poFrame, text=self.costmer_option[5], variable=self.radio_val, value=5).grid(sticky=W, row=4, column=2)
        self.radio_val_7 = Radiobutton(poFrame, text=self.costmer_option[6], variable=self.radio_val, value=6).grid(sticky=W, row=5, column=1)
        self.type_label = Label(poFrame, text="操作类型：")
        self.type_label.grid(sticky=W, row=6, column=1)
        self.option_type = ["新PO", "修改PO"]
        self.radio_type = IntVar()
        self.radio_type_1 = Radiobutton(poFrame, text=self.option_type[0], variable=self.radio_type, value=0).grid(sticky=W, row=7, column=1)
        # 按钮
        self.commit_button = Button(poFrame, text="点击读取PO", bg="lightblue", width=18, command=self.commit_form)
        self.commit_button.grid(sticky=W, row=8, column=1)
        # 显示文字框
        self.file_show_label = Label(poFrame, text="请在D盘建立【4DepPo】文件夹，放入文件后点击按钮。", wraplength=400)
        self.file_show_label.grid(sticky=W, row=9, column=1, columnspan=10)
        self.file_show_label = Label(poFrame, text="※ 注意选择好【客户】和【操作类型】。", wraplength=400)
        self.file_show_label.config(fg="red")
        self.file_show_label.grid(sticky=W, row=10, column=1, columnspan=10)

    def commit_form(self):
        self.add_data_title = ['TYPE', 'PO号', '款号', '英文款名', '订单数量', '客人船期',
                               '目的港', '贸易方式', '走货方式', '商标', 'version', '面料颜色', '结汇币种', '季节号', '来单日期']
        self.add_data_title_size_next = self.add_data_title + ['option', 'size', 'quantity']
        self.add_data_title_size_zara = self.add_data_title + ['size', 'quantity']
        self.number_item = ['订单数量', 'quantity', 'version']
        self.table_data = pd.DataFrame(data=None, columns=self.add_data_title)
        self.table_data_next_size = pd.DataFrame(data=None, columns=self.add_data_title_size_next)
        self.zara_base_size = ['XXS', 'XS', 'S', 'M', 'L', 'XL', 'XXL', '3XL', '4XL']
        self.table_data_zara_size = pd.DataFrame(data=None, columns=self.add_data_title_size_zara)
        self.table_value = []
        self.delete_key = []
        self.delete_str_key = ['NEXT', 'Product', 'Contact', 'Assign', 'Item', 'See', 'Terms', 'Refer', 'Manuals', 'Refurb', 'Description', 'IMPORTANT',
                               ')', 'Type', 'Fabric', 'Trans', 'Contract', 'Del', 'Ex-Fact', 'Total', 'Supplier', 'Booking']
        self.nextPoList = []

        # 选择【BS】客户报错
        if self.radio_val.get() == 3:
            tmessage.showerror('错误', '人生苦短,目前还没有该客户的读取功能,施主请慢走！')
            return None
        try:
            for lroot, ldirs, lfiles in os.walk(self.local_po_file):
                if len(lfiles) == 0:
                    tmessage.showerror('错误', '没有找到任何文件！')
                    return None
                for lfile in lfiles:
                    temp_file_root = os.path.join(lroot, lfile)
                    # print(temp_file_root)
                    ctime = parser.parse(time.ctime(
                        os.path.getctime(temp_file_root)))
                    create_time = ctime.strftime('%Y-%m-%d %H:%M:%S')
                    self.file_to_dataframe(temp_file_root, str(lfile).split('.')[0], self.radio_val.get(), create_time)
            # 客户("NEXT", "SLATERS", "DEVRED", "BS", "ITX", "CELIO", "IZAC")
            self.table_data['客户'] = str(self.costmer_option[self.radio_val.get()])
            self.table_value = self.set_table_dataframe(self.table_data)
            # print('----------------------------------')
            # print(self.table_data)
            # print(self.table_data_next_size)
            self.update_db()
            self.update_po_size_db()
            tmessage.showinfo('成功', '恭喜操作成功，请到勤哲系统中查看结果吧！')
        except:
            tmessage.showerror('错误', '人生苦短,程序出错了,请联系信息部孙适老师！')

    def update_po_size_db(self):
        if self.radio_val.get() == 0:
            self.table_value_next_size = self.set_table_dataframe(self.table_data_next_size)
            self.update_size_db(self.radio_type.get(), self.add_data_title_size_next, self.nextPoList, self.table_value_next_size, 'D_4DepNEXTSize')
        elif self.radio_val.get() == 4:
            self.table_value_zara_size = self.set_table_dataframe(self.table_data_zara_size)
            self.update_size_db(self.radio_type.get(), self.add_data_title_size_zara, self.nextPoList, self.table_value_zara_size, 'D_4DepZARASize')
        elif self.radio_val.get() == 2:
            # DEVRED 跟zara一样
            self.table_value_zara_size = self.set_table_dataframe(self.table_data_zara_size)
            self.update_size_db(self.radio_type.get(), self.add_data_title_size_zara, self.nextPoList, self.table_value_zara_size, 'D_4DepDevredSize')
        elif self.radio_val.get() == 5:
            # CELIO 跟zara一样
            self.table_value_zara_size = self.set_table_dataframe(self.table_data_zara_size)
            self.update_size_db(self.radio_type.get(), self.add_data_title_size_zara, self.nextPoList, self.table_value_zara_size, 'D_4DepCelioSize')
        elif self.radio_val.get() == 6:
            # IZAC 跟zara一样
            self.table_value_zara_size = self.set_table_dataframe(self.table_data_zara_size)
            self.update_size_db(self.radio_type.get(), self.add_data_title_size_zara, self.nextPoList, self.table_value_zara_size, 'D_4DepIZACSize')
        # elif self.radio_val.get() == 1:
        #     self.table_value_slaters_size = self.set_table_dataframe(self.table_data_zara_size)
        #     # print(self.table_value_slaters_size)
        #     self.update_size_db(self.radio_type.get(), self.add_data_title_size_zara,
        #                         self.nextPoList, self.table_value_slaters_size, 'D_4DepSLATERSSize')

    def set_table_dataframe(self, table_data):
        table_value = []
        table_data['CreateDate'] = str(datetime.datetime.now()).split('.')[0]
        table_value.append([tuple(row) for row in table_data.values])
        return table_value

    def file_to_dataframe(self, io, lfile, radioType, fileDate):
        if radioType == 0:
            if self.radio_type.get() == 0:
                self.excel_to_dataframe_next_add(io, fileDate)
            else:
                self.excel_to_dataframe_next_update(io, fileDate)
        elif radioType == 4:
            self.pdf_to_dataframe_itx_add(io, fileDate)
        elif radioType == 1:
            self.excel_to_dataframe_slaters_add(io, fileDate)
        elif radioType == 2:
            self.pdf_to_dataframe_devred_add(io, fileDate)
        elif radioType == 5:
            self.pdf_to_dataframe_celio_add(io, fileDate)
        elif radioType == 6:
            self.pdf_to_dataframe_izac_add(io, fileDate)

    def pdf_to_dataframe_izac_add(self, io, fileDate):
        pdf = pdfplumber.open(io)
        po_no = ''
        style_no = ''
        en_style_name = ''
        destination = ''  # 目的港
        trade_mode = ''  # 贸易方式
        shipping_mode = ''  # 走货方式
        fabric_list = []
        currency = ''
        season = ''
        come_date = ''
        sum_qty = 0
        hs_code = ''
        unit_price = ''
        table1 = []
        for page in pdf.pages:
            file_txt = str(page.extract_text())
            delivery_date = ''
            if file_txt.__contains__('SHIPMENT DATE:'):
                delivery_date = self.format_shipping_date(self.get_value_two_word(file_txt, 'SHIPMENT DATE:', None).strip()[:10])
            if file_txt.__contains__('Incoterm :'):
                trade_mode = self.get_value_two_word(file_txt, 'Incoterm :', 'Transport :').strip()
                shipping_mode = self.get_value_two_word(file_txt, 'Transport :', 'Transporteur :').strip()
            if page.page_number == 1:
                po_no_tmp = self.get_value_two_word(file_txt, 'Order Number :', 'Invoicing :').strip()
                po_no = "".join(list(filter(str.isdigit, po_no_tmp[:6])))
                destination = self.get_value_two_word(file_txt, 'Consigne', 'DALIAN CHINA').strip().split(' ')[-1]
                en_style_name = self.get_value_two_word(file_txt, 'Classification : ', 'Désignation :').strip()
                style_no = self.get_value_two_word(file_txt, 'Article Code :', 'Nomenclature douanière :').strip().split(' ')[0]
                hs_code = self.get_value_two_word(file_txt, 'Nomenclature douanière :', 'Collection :').strip()
                season = self.get_value_two_word(file_txt, 'Collection :', 'Thème :').strip()
                trade_mode = self.get_value_two_word(file_txt, 'Incoterm :', 'Transport :').strip()
                shipping_mode = self.get_value_two_word(file_txt, 'Transport :', 'Transporteur :').strip()
                come_date = self.format_shipping_date(self.get_value_two_word(po_no_tmp, 'Date :', 'Invoicing :').strip()[:10])
                unit_price = self.get_value_two_word(file_txt, 'Net Buying Price ex VAT :', 'Discount :').strip()
                # 获取pdf第一个表格
                table = page.extract_tables()[0]
                table1 = page.extract_tables()[1]
                tmp_qty_info = 'Total Quantity :' + self.get_value_two_word(file_txt, 'Total Quantity :', None)
                tmp_qty = self.get_value_two_word(tmp_qty_info, 'Total Quantity :', style_no).strip().replace(' ', '')
                sum_qty = int(str(tmp_qty))
                currency = self.get_value_two_word(tmp_qty_info, 'CURRENCY :', '\n').strip()
                for row in table:
                    row.pop()
                    tmp_size_type = row[0]
                    if tmp_size_type is not None and tmp_size_type != '' and tmp_size_type != 'Color' and tmp_size_type != 'Total':
                        del (row[0])
                        fabric_list = tmp_size_type.split('\n')
        # 获取size和颜色的列表
        tmp_size_table = table1[1]
        tmp_color_list = tmp_size_table[0].split('\n')
        tmp_size_color = tmp_color_list[0]
        tmp_size_list = tmp_size_table[1].split('\n')
        tmp_qty_list = tmp_size_table[3].split('\n')
        size_idx = 0
        # 尺码和数量的list
        size_list = {}
        size_detail_list = {}
        size_list_val = []
        size_detail_list_val = []
        for color_idx in range(len(tmp_color_list)):
            if tmp_color_list[color_idx] != tmp_size_color:
                tmp_size_color = tmp_color_list[color_idx]
                size_idx = size_idx + 1
                size_list_val = []
                size_detail_list_val = []
            size_list_val.append(tmp_size_list[color_idx])
            size_detail_list_val.append(tmp_qty_list[color_idx])
            size_list[tmp_color_list[color_idx]] = size_list_val
            size_detail_list[tmp_color_list[color_idx]] = size_detail_list_val

        for n_idx in range(len(fabric_list)):
            # 'TYPE', 'PO号', '款号', '英文款名', '订单数量', '客人船期', '目的港', '贸易方式', '走货方式', '商标', 'version', '面料颜色', '结汇币种', '季节号', '来单日期'
            # 由于其他的客户没有HS编码，【izac】这个客户的【TYPE】列存放HS编码。
            # 由于其他的客户没有结汇单价，【izac】这个客户的【商标】列存放结汇单价。
            input_color = fabric_list[n_idx]
            temp_po_detail = [hs_code, po_no, style_no, en_style_name, sum_qty, delivery_date,
                              destination, trade_mode, shipping_mode, unit_price, 1, input_color, currency, season, come_date]
            po_detail = pd.Series(temp_po_detail, index=self.add_data_title)
            self.table_data = self.table_data.append(po_detail, ignore_index=True)
            for s_idx in range(len(size_list[input_color])):
                if int(str(size_detail_list[input_color][s_idx])) > 0:
                    temp_size_detail = temp_po_detail + \
                        [self.change_izac_size(size_list[input_color][s_idx]), int(str(size_detail_list[input_color][s_idx]))]
                    zara_size_detail = pd.Series(temp_size_detail, index=self.add_data_title_size_zara)
                    self.table_data_zara_size = self.table_data_zara_size.append(zara_size_detail, ignore_index=True)
        # print(self.table_data)
        # print(self.table_data_zara_size)
        self.nextPoList.append(po_no)

    def change_izac_size(self, size):
        if str(size).__contains__(' ('):
            size_change_list = {'42-44 (': 'XS', '46-48 (': 'S', '50-52 (': 'M', '54-56 (': 'L', '58-60 (': 'XL', '62-64 (': 'XXL', '66-68 (': 'XXXL'}
            return size_change_list[size]
        return size

    def pdf_to_dataframe_celio_add(self, io, fileDate):
        pdf = pdfplumber.open(io)
        packaging_type = ''
        po_no = ''
        style_no = ''
        en_style_name = ''
        destination = ''  # 目的港
        trade_mode = ''  # 贸易方式
        shipping_mode = ''  # 走货方式
        fabric = ''
        currency = ''
        season = ''
        come_date = ''
        sum_qty = 0
        size_list = []
        size_detail_list = []
        unit_price = ''
        for page in pdf.pages:
            if page.page_number == 1:
                file_txt = str(page.extract_text())
                po_no = self.get_value_two_word(file_txt, 'Purchase Order Nr:', 'Order type:').strip()
                destination = self.get_value_two_word(file_txt, 'Country: ', 'Tel:').strip()
                trade_destination = self.get_value_two_word(file_txt, 'Incoterm:', 'Origin of Fabric:').strip()
                trade_mode = trade_destination.split(' ')[0]
                shipping_mode = self.get_value_two_word(file_txt, 'Shipping mode:', 'Agent:').strip()
                season_tmp = self.get_value_two_word(file_txt, 'Season:', 'Trading address:').strip()
                season = season_tmp.split(' ')[0][0] + season_tmp.split(' ')[1]
                come_date = self.format_shipping_date(self.get_value_two_word(file_txt, 'initial order date:', 'Purchase Order Nr:').strip())
                delivery_date = self.format_shipping_date(self.get_value_two_word(
                    file_txt, 'Initial delivery date to forwarder:', 'Delayed delivery date to forwarder:').strip())
                # 获取pdf表格中的内容
                for table in page.extract_tables():
                    for row in table:
                        if row[0] is not None and row[0] != 'PACKAGING\nTYPE':
                            packaging_type = row[0]
                            style_no = row[1]
                            fabric = row[5]
                            unit_price = str(row[12]).split(' ')[0].replace(',', '.')
                            currency = str(row[12]).split(' ')[1]
                            size_list.append(str(row[6]))
                            size_detail_list.append(row[10])
                            sum_qty = sum_qty + int(row[10])
        # 'TYPE', 'PO号', '款号', '英文款名', '订单数量', '客人船期', '目的港', '贸易方式', '走货方式', '商标', 'version', '面料颜色', '结汇币种', '季节号', '来单日期'
        # 由于其他的客户没有美金单价，【celio】这个客户的【商标】列存放美金单价。
        temp_po_detail = [packaging_type, po_no, style_no, en_style_name, sum_qty, delivery_date,
                          destination, trade_mode, shipping_mode, unit_price, 1, fabric, currency, season, come_date]
        po_detail = pd.Series(temp_po_detail, index=self.add_data_title)
        self.table_data = self.table_data.append(po_detail, ignore_index=True)
        for s_idx in range(len(size_list)):
            if int(str(size_detail_list[s_idx])) > 0:
                temp_size_detail = temp_po_detail + [size_list[s_idx], int(str(size_detail_list[s_idx]))]
                zara_size_detail = pd.Series(temp_size_detail, index=self.add_data_title_size_zara)
                self.table_data_zara_size = self.table_data_zara_size.append(zara_size_detail, ignore_index=True)
        self.nextPoList.append(po_no)

    def pdf_to_dataframe_devred_add(self, io, fileDate):
        pdf = pdfplumber.open(io)
        style_no = ''
        en_style_name = ''
        destination = ''  # 目的港
        trade_mode = ''  # 贸易方式
        fabric = ''
        currency = ''
        season = ''
        come_date = ''
        qty_list = []
        deliv_list = []
        first_size = ''
        for page in pdf.pages:
            file_txt = str(page.extract_text())
            style_no = self.get_value_two_word(file_txt, 'SAMPLE REFERENCE', 'COLOR').strip()
            en_style_name = self.get_value_two_word(file_txt, 'NAME', 'MODEL').strip()
            trade_destination = self.get_value_two_word(file_txt, 'INCOTERM', 'TOTAL QUANTITY').strip()
            destination = trade_destination.split(' ')[1]
            trade_mode = trade_destination.split(' ')[0]
            fabric = self.get_value_two_word(file_txt, 'COLOR', '\nDESCRIPTION').strip()
            currency = self.get_value_two_word(file_txt, 'CURRENCY', 'INCOTERM').strip()
            season = self.get_value_two_word(file_txt, 'SEASON :', 'Date:').strip()
            come_date = self.format_shipping_date(self.get_value_two_word(file_txt, 'Date:', 'SFAM:').strip())
            if file_txt.__contains__('TXXS'):
                first_size = 'TXXS'
            elif file_txt.__contains__('TXS'):
                first_size = 'TXS'
            elif file_txt.__contains__('TS'):
                first_size = 'TS'
            else:
                first_size = 'TM'
            find_first_size = '\n' + first_size + ' '
            delivery_date = self.get_value_two_word(file_txt, 'QTY', find_first_size).strip()
            deliv_list = delivery_date.split(' ')
            for i in range(len(deliv_list)):
                deliv_list[i] = self.format_shipping_date(deliv_list[i])
            # 先取数量明细字段
            num_detail_txt = self.get_value_two_word(file_txt, 'SIZE TOTAL QTY', 'REMARKS').strip()
            # 数量取得
            num_txt = self.get_value_two_word(num_detail_txt, 'TOTAL', 'PACKING REQUIREMENTS').strip()
            qty_list = np.split(num_txt.split(' '), [1, len(deliv_list) + 1])[1]
        size_list = []
        size_content_flag = False
        size_detail_list = [[] for i in range(len(deliv_list))]
        # 获取pdf表格中的内容
        for table in page.extract_tables():
            for row in table:
                if row[0] == first_size:
                    size_content_flag = True
                if row[0] == 'TOTAL':
                    size_content_flag = False
                if size_content_flag:
                    size_list.append(row[0].replace('T', ''))
                    for n_idx in range(len(deliv_list)):
                        size_detail_list[n_idx].append(row[n_idx + 2])

        for n_idx in range(len(deliv_list)):
            # 'TYPE', 'PO号', '款号', '英文款名', '订单数量', '客人船期', '目的港', '贸易方式', '走货方式', '商标', 'version', '面料颜色', '结汇币种', '季节号', '来单日期'
            temp_po_detail = ['', '', style_no, en_style_name, qty_list[n_idx], deliv_list[n_idx],
                              destination, trade_mode, 'SHIP', 'DEVRED', 1, fabric, currency, season, come_date]
            po_detail = pd.Series(temp_po_detail, index=self.add_data_title)
            self.table_data = self.table_data.append(po_detail, ignore_index=True)
            for s_idx in range(len(size_list)):
                if int(str(size_detail_list[n_idx][s_idx])) > 0:
                    temp_size_detail = temp_po_detail + [size_list[s_idx], int(str(size_detail_list[n_idx][s_idx]))]
                    zara_size_detail = pd.Series(temp_size_detail, index=self.add_data_title_size_zara)
                    self.table_data_zara_size = self.table_data_zara_size.append(zara_size_detail, ignore_index=True)
        self.nextPoList.append(style_no)

    def excel_to_dataframe_slaters_add(self, io, fileDate):
        excelData = pd.read_excel(io, header=None, keep_default_na=False, sheet_name='ORDER')
        nrows = excelData.shape[0]
        po = excelData.iloc[9, 1]
        style_no = excelData.iloc[9, 1]
        en_style_name = excelData.iloc[18, 3]
        # 循环QTY列
        qty_list = []
        qty_flag = False
        for row in range(nrows):
            qty_value = excelData.iloc[row, 3]
            if str(qty_value) == 'QTY':
                qty_flag = True
                continue
            if str(qty_value) == str(en_style_name):
                break
            if qty_flag and str(qty_value).strip() != '':
                qty_list.append(qty_value)
        # 商标
        label = excelData.iloc[6, 3]
        # 面料颜色
        fabric = str(excelData.iloc[12, 1]) + '-' + str(excelData.iloc[13, 1]) + '-' + str(excelData.iloc[14, 1])
        # 取季节号中的年
        season_year = str(excelData.iloc[10, 1])[-4:]
        # 季节号
        season = str(excelData.iloc[10, 1]).replace('/', '').replace('20', '')
        # 循环DELIV列
        deliv_list = []
        deliv_flag = False
        for row in range(nrows):
            deliv_value = excelData.iloc[row, 4]
            if str(deliv_value) == 'DELIV':
                deliv_flag = True
                continue
            if deliv_flag and str(deliv_value).strip() != '':
                deliv_list.append(self.change_shipping_date(deliv_value, season_year))
        # 来单日期
        come_date = excelData.iloc[8, 1]
        po_list = []
        type_list = []
        po_idx = 0
        if len(qty_list) == 1:
            po_list.append(po)
            type_list.append(en_style_name)
        else:
            for qty in qty_list:
                po_idx += 1
                po_list.append(po + '/' + str(po_idx))
                type_list.append(en_style_name)
        if en_style_name == '3PC' or en_style_name == '2PC':
            en_style_name = 'MENS SUIT'
        # 'TYPE', 'PO号', '款号', '英文款名', '订单数量', '客人船期', '目的港', '贸易方式', '走货方式', '商标', 'version', '面料颜色', '结汇币种', '季节号', '来单日期'
        po_df = pd.DataFrame(data=None, columns=self.add_data_title)
        po_df[self.add_data_title[0]] = type_list
        po_df[self.add_data_title[1]] = po_list
        po_df[self.add_data_title[2]] = style_no
        po_df[self.add_data_title[3]] = en_style_name
        po_df[self.add_data_title[4]] = qty_list
        po_df[self.add_data_title[5]] = deliv_list
        po_df[self.add_data_title[6]] = '英国'
        po_df[self.add_data_title[7]] = 'DDP'
        po_df[self.add_data_title[8]] = 'SHIP'
        po_df[self.add_data_title[9]] = label
        po_df[self.add_data_title[10]] = 1
        po_df[self.add_data_title[11]] = fabric
        po_df[self.add_data_title[12]] = 'USD'
        po_df[self.add_data_title[13]] = season
        po_df[self.add_data_title[14]] = come_date
        self.table_data = self.table_data.append(po_df, ignore_index=True)

    def pdf_to_dataframe_itx_add(self, io, fileDate):
        pdf_file = self.get_file_content(io)
        options = {}
        options['detect_direction'] = 'true'
        res_pdf = self.client.basicAccuratePdf(pdf_file, options)
        ocr_msg = ''
        for i in res_pdf.get('words_result'):
            ocr_msg = ocr_msg + '{}\n'.format(i.get('words'))
        is_green = ''
        if 'JOIN' in ocr_msg and 'LIFE' in ocr_msg:
            is_green = '环保订单'
        pdf = pdfplumber.open(io)
        count = 0
        send_to = '江苏'
        po_list = []
        po = ''
        season = ''
        come_date_list = []
        come_date = ''
        handover_date_list = []
        handover_date = ''
        incoterm_list = []
        incoterm = ''
        transport_mode_list = []
        transport_mode = ''
        style_no = ''
        en_style_name = ''
        currency = 'CNY'
        first_page_color_list = []
        first_page_size_num_list = []
        first_page_po_num_list = []
        size_list = []
        one_page_flag = False
        style_no_flag = False
        for page in pdf.pages:
            count += 1
            if count == 1:
                file_txt = str(page.extract_text()).split('TOTAL ORDER')[0]
                send_to_txt = self.get_value_two_word(
                    file_txt, 'SEND TO', 'ORDER NR').strip()
                if send_to_txt.__contains__('ESPAÑA'):
                    send_to = '西班牙'
                    currency = 'USD'
                elif send_to_txt.__contains__('PAÍSES'):
                    send_to = '巴黎'
                    currency = 'USD'
                info_txt = self.get_value_two_word(
                    file_txt, 'ORDER NR', 'TOTAL ORDER').strip()
                file_txt_list = info_txt.split('\n')
                if send_to == '西班牙' or send_to == '巴黎':
                    # market_of_origin_list = ['MAINLAND'] 原产地市场
                    for temp_des in file_txt_list:
                        if str(temp_des).__contains__('MAINLAND'):
                            style_no = temp_des.split(' ')[0]
                            en_style_name = self.get_value_two_word(
                                temp_des, style_no, 'MAINLAND').strip().replace('[', '').replace(']', '')
                            break
                        elif str(temp_des).__contains__('CAMBODIA'):
                            style_no = temp_des.split(' ')[0]
                            en_style_name = self.get_value_two_word(
                                temp_des, style_no, 'CAMBODIA').strip().replace('[', '').replace(']', '')
                            break
                else:
                    for temp_des in file_txt_list:
                        if str(temp_des).__contains__('市场'):
                            style_no_flag = True
                            continue
                        if style_no_flag:
                            style_no = temp_des.split(' ')[0]
                            en_style_name = self.get_value_two_word(temp_des, style_no, None).strip()
                            style_no_flag = False
                order_flag = False
                season_flag = False
                other_info_flag = False
                po_detail_info_flag = False
                for table in page.extract_tables():
                    for row1 in table:
                        row = [self.replace_exist_word(i) for i in row1]
                        if str(row[0]).__contains__('ORDER NR'):
                            order_flag = True
                            continue
                        if order_flag and row[0] != '':
                            po = str(row[0]).replace('PRE', '').strip()
                            come_date = self.format_shipping_date(str(row[2]))
                            order_flag = False
                        if str(row[0]).__contains__('SEASON'):
                            season_flag = True
                            continue
                        if season_flag and row[0] != '':
                            season = str(row[0]).replace(' ', '')
                            season_flag = False
                        if str(row[0]).strip() == 'FROM' or str(row[0]).strip() == 'TO' or str(row[0]).strip() == 'TO / 交货地点':
                            other_info_flag = True
                            continue
                        if other_info_flag:
                            handover_date = self.format_shipping_date(str(row[2]))
                            incoterm = str(row[4])
                            transport_mode = str(row[6])
                            if send_to == '江苏':
                                incoterm = str(row[5])
                                transport_mode = str(row[7])
                            other_info_flag = False
                        if str(row[0]).strip() == 'COLOUR' or str(row[0]).strip() == 'COLOUR / 颜色':
                            size_list = row[1:][:-1]
                            po_detail_info_flag = True
                            continue
                        if str(row[0]).strip() == 'TOTAL' or str(row[0]).strip() == 'TOTAL / 总数':
                            po_detail_info_flag = False
                            continue
                        if po_detail_info_flag and row[0] != '':
                            po_list.append(po)
                            come_date_list.append(come_date)
                            handover_date_list.append(handover_date)
                            incoterm_list.append(incoterm)
                            transport_mode_list.append(transport_mode)
                            first_page_color_list.append(row[0])
                            first_page_size_num_list.append(row[1:][:-1])
                            first_page_po_num_list.append(int(str(row[-1]).replace(',', '')))
                            continue
                        if file_txt.__contains__('INCOTERM') and len(first_page_color_list) > 0:
                            pass
                        elif file_txt.__contains__('INCOTERM') and len(first_page_color_list) == 0:
                            one_page_flag = True
                        else:
                            po_list = []
                            come_date_list = []
                            handover_date_list = []
                            incoterm_list = []
                            transport_mode_list = []
                            first_page_color_list = []
                            first_page_size_num_list = []
                            first_page_po_num_list = []
            else:
                for table in page.extract_tables():
                    for row in table:
                        row = [self.replace_exist_word(i) for i in row1]
                        if row[0] is None or str(row[0]) == '':
                            continue
                        row = list(filter(None, row))
                        if str(row[0]).__contains__('LOGISTIC ORDER'):
                            order_flag = True
                            continue
                        if order_flag and row[0] != '':
                            po = str(row[0]).replace('PRE', '').strip()
                            incoterm = str(row[2])
                            handover_date = self.format_shipping_date(str(row[4]))
                            transport_mode = str(row[5])
                            order_flag = False
                        if str(row[0]).strip() == 'COLOUR' or str(row[0]).strip() == 'COLOUR / 颜色':
                            if one_page_flag:
                                size_list = row[1:][:-1]
                                one_page_flag = False
                            po_detail_info_flag = True
                            continue
                        if str(row[0]).strip() == 'TOTAL' or str(row[0]).strip() == 'TOTAL / 总数':
                            po_detail_info_flag = False
                            continue
                        if po_detail_info_flag and row[0] != '':
                            po_list.append(po)
                            incoterm_list.append(incoterm)
                            come_date_list.append(come_date)
                            handover_date_list.append(handover_date)
                            transport_mode_list.append(transport_mode)
                            first_page_color_list.append(row[0])
                            first_page_size_num_list.append(row[1:][:-1])
                            first_page_po_num_list.append(int(str(row[-1]).replace(',', '')))
                            continue
            for n_idx in range(len(first_page_color_list)):
                temp_po_detail = [
                    is_green,
                    po_list[n_idx],
                    style_no,
                    en_style_name,
                    first_page_po_num_list[n_idx],
                    handover_date_list[n_idx],
                    send_to,
                    incoterm_list[n_idx],
                    transport_mode_list[n_idx],
                    '',
                    1,
                    first_page_color_list[n_idx],
                    currency,
                    season,
                    come_date_list[n_idx]]
                po_detail = pd.Series(temp_po_detail, index=self.add_data_title)
                self.table_data = self.table_data.append(po_detail, ignore_index=True)
                for s_idx in range(len(size_list)):
                    temp_size_detail = temp_po_detail + [size_list[s_idx], int(str(first_page_size_num_list[n_idx][s_idx]).replace(',', ''))]
                    zara_size_detail = pd.Series(temp_size_detail, index=self.add_data_title_size_zara)
                    self.table_data_zara_size = self.table_data_zara_size.append(zara_size_detail, ignore_index=True)
            # print(self.table_data)
            # print(self.table_data_zara_size)
            self.nextPoList.extend(po_list)

    def excel_to_dataframe_next_add(self, io, fileDate):
        excel_header = []
        excelData = pd.read_excel(io, header=None, keep_default_na=False)
        excelCol = excelData.shape[1] + 1
        for h_idx in range(1, excelCol):
            excel_header.append('列' + str(h_idx))
        df = pd.DataFrame(excelData.values, columns=excel_header)
        df.dropna(axis=0, how='all')
        temp_style_no_list = self.check_str_key(df['列1'])
        style_no_list = []
        en_name_type_list = self.check_str_key(df['列2'])
        en_name_list = []
        type_list = []
        style_idx = -1
        temp_en_name = ''
        fabric_color_list = []
        for en_name_type_val in en_name_type_list:
            if len(en_name_type_val) > 3:
                style_idx = style_idx + 1
                temp_en_name = en_name_type_val
                continue
            en_name_list.append(temp_en_name.split('Style')[0] + 'Style')
            fabric_color_list.append(temp_en_name.split('Style')[1])
            style_no_list.append(temp_style_no_list[style_idx])
            type_list.append(en_name_type_val)
        # 走货方式
        trans_list = self.check_str_key(df['列5'])
        # PO号
        contract_no_list = self.check_str_key(df['列6'])
        contract_del_list = self.check_str_key(df['列7'])
        po_list = []
        for po_idx in range(0, len(contract_no_list)):
            po_list.append(str(contract_no_list[po_idx]) + '-' + str(contract_del_list[po_idx]))
        shipping_list = []
        ex_fact_list = self.check_str_key(df['列8'])
        c_year = str(fileDate).split('-')[0]
        c_month = str(fileDate).split('-')[1]
        for ex_fact in ex_fact_list:
            year = c_year
            if int(str(ex_fact).split('/')[1]) <= int(c_month):
                year = int(c_year) + 1
            shipping_list.append(str(year) + '-' + str(ex_fact).split('/')[1] + '-' + str(ex_fact).split('/')[0])
        qty_list = self.check_str_key(df['列9'])
        next_size_num_list = self.get_color_num_next(df, excelCol)
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
        po_df[self.add_data_title[11]] = fabric_color_list
        po_df[self.add_data_title[12]] = 'USD'
        po_df[self.add_data_title[13]] = ''
        po_df[self.add_data_title[14]] = str(datetime.datetime.now()).split('.')[0]
        self.table_data = self.table_data.append(po_df, ignore_index=True)
        for n_idx in range(len(po_list)):
            po_df_size = pd.DataFrame(data=None, columns=self.add_data_title_size_next)
            po_df_size[self.add_data_title_size_next[15]] = next_size_num_list[n_idx]['sizeNo']
            po_df_size[self.add_data_title_size_next[16]] = next_size_num_list[n_idx]['size']
            po_df_size[self.add_data_title_size_next[17]] = next_size_num_list[n_idx]['num']
            po_df_size[self.add_data_title_size_next[0]] = type_list[n_idx]
            po_df_size[self.add_data_title_size_next[1]] = po_list[n_idx]
            po_df_size[self.add_data_title_size_next[2]] = style_no_list[n_idx]
            po_df_size[self.add_data_title_size_next[3]] = en_name_list[n_idx]
            po_df_size[self.add_data_title_size_next[4]] = qty_list[n_idx]
            po_df_size[self.add_data_title_size_next[5]] = shipping_list[n_idx]
            po_df_size[self.add_data_title_size_next[6]] = ''
            po_df_size[self.add_data_title_size_next[7]] = ''
            po_df_size[self.add_data_title_size_next[8]] = trans_list[n_idx]
            po_df_size[self.add_data_title_size_next[9]] = ''
            po_df_size[self.add_data_title_size_next[10]] = 1
            po_df_size[self.add_data_title_size_next[11]] = fabric_color_list[n_idx]
            po_df_size[self.add_data_title_size_next[12]] = 'USD'
            po_df_size[self.add_data_title_size_next[13]] = ''
            po_df_size[self.add_data_title_size_next[14]] = str(datetime.datetime.now()).split('.')[0]
            self.table_data_next_size = self.table_data_next_size.append(po_df_size, ignore_index=True)
        self.nextPoList.extend(po_list)

    def excel_to_dataframe_next_update(self, io, fileDate):
        excelData = pd.read_excel(io, header=None, keep_default_na=False)
        arr_excel_val = []
        jump_flag = False
        for e_v in excelData.values:
            if str(e_v[0]) == 'Item':
                jump_flag = True
                continue
            if not jump_flag or str(e_v[0]) != '':
                if str(e_v[1]) != '':
                    arr_excel_val.append(e_v.tolist())
                    continue
        style_no_list = []
        en_name_list = []
        type_list = []
        trans_list = []
        po_list = []
        temp_style_no = ''
        temp_en_name_list = ''
        for po_info_val in arr_excel_val:
            if str(po_info_val[0]) != '':
                temp_style_no = po_info_val[0]
                temp_en_name_list = po_info_val[1]
        print(arr_excel_val)

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

    def get_color_num_next(self, next_df, excelCol):
        cols = []
        for h_idx in range(10, excelCol - 1):
            cols.append(h_idx)
        res_df = next_df[next_df.columns[cols]]
        color_size_num_list = []
        color_list = []
        size_list = []
        num_list = []
        step_idx = 1
        for res_idx in range(len(res_df)):
            # print(res_df['列14'][res_idx])
            if 'Assign site / factory.' in str(res_df['列14'][res_idx]):
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
            if res_idx == len(res_df) - 1:
                color_size_num_list.append(
                    self.set_color_size_num_df(color_list, size_list, num_list))
        return color_size_num_list

    def set_color_size_num_df(self, color_list, size_list, num_list):
        color_size_num_df = pd.DataFrame(columns=['sizeNo', 'size', 'num'])
        color_size_num_df['sizeNo'] = color_list
        color_size_num_df['size'] = size_list
        color_size_num_df['num'] = num_list
        return color_size_num_df

    def update_db(self):
        dbCol = self.add_data_title[:]
        dbCol.append('客户')
        dbCol.append('CreateDate')
        conn = pymssql.connect(self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        if len(self.nextPoList) > 0:
            # del_tuple = tuple(self.nextPoList)
            del_tuple = []
            for tuple_po in self.nextPoList:
                del_tuple.append((tuple_po, tuple_po))
            delSql = 'delete from D_4DepPoInfo where version = 1 and (PO号 = (%s) OR 款号 = (%s))'
            cursor.executemany(delSql, del_tuple)
        insertValue = []
        for tabVal in self.table_value:
            insertValue += tabVal
        insertSql = 'INSERT INTO D_4DepPoInfo VALUES ('
        for colVal in dbCol:
            if colVal == 'CreateDate':
                insertSql += '%s'
                continue
            if colVal in self.number_item:
                insertSql += '%d, '
                continue
            insertSql += '%s, '
        insertSql += ')'
        cursor.executemany(insertSql, insertValue)
        conn.commit()
        conn.close()

    def update_size_db(self, insertType, dbColVal, poList, insertItem, tableName):
        dbCol = dbColVal[:]
        dbCol.append('CreateDate')
        conn = pymssql.connect(self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        if len(poList) > 0 and insertType == 0:
            # del_tuple = tuple(poList)
            del_tuple = []
            for tuple_po in poList:
                del_tuple.append((tuple_po, tuple_po))
            delSql = 'delete from ' + tableName + ' where version = 1 and (PO号 = (%s) OR 款号 = (%s))'
            cursor.executemany(delSql, del_tuple)
        insertValue = []
        for tabVal in insertItem:
            insertValue += tabVal
        insertSql = ''
        if insertType == 0:
            insertSql = 'INSERT INTO ' + tableName + ' VALUES ('
            for colVal in dbCol:
                if colVal == 'CreateDate':
                    insertSql += '%s'
                    continue
                if colVal in self.number_item:
                    insertSql += '%d, '
                    continue
                insertSql += '%s, '
            insertSql += ')'
        cursor.executemany(insertSql, insertValue)
        conn.commit()
        conn.close()

    def get_value_two_word(self, txt_str, one, two):
        if one == None:
            return txt_str[:txt_str.find(two)]
        if two == None:
            return txt_str[txt_str.find(one) + len(one):]
        return txt_str[txt_str.find(one) + len(one):txt_str.find(two)]

    def zara_size_num_split(self, size_list, num_list):
        res_size_list = []
        res_num_list = []
        temp_num_list = []
        for n_idx in range(len(num_list)):
            for s_idx in range(len(size_list)):
                temp_size_list = str(size_list[s_idx]).split('-')
                res_size_list.append(temp_size_list[0])
                res_size_list.append(temp_size_list[1])
                temp_num_list.append(num_list[n_idx][s_idx])
                temp_num_list.append(num_list[n_idx][s_idx])
                res_num_list.append(temp_num_list)
        return {'size_list': res_size_list, 'num_list': res_num_list}

    def format_shipping_date(self, temp_str):
        if temp_str == '':
            return ''
        if temp_str.__contains__('DEADLINE'):
            temp_str = str(temp_str).replace('DEADLINE', '')
        if temp_str.__contains__('.'):
            temp_str = str(temp_str).replace('.', '/')
        temp_str = str(temp_str).strip()
        t_handover_date = temp_str.split('/')
        return t_handover_date[2] + '-' + t_handover_date[1] + '-' + t_handover_date[0]

    def get_file_content(self, filePath):
        with open(filePath, "rb") as fp:
            return fp.read()

    def replace_exist_word(self, value):
        word = ['D\n', 'R\n', 'A\n', 'F\n', 'T\n']
        for item in word:
            if str(value).__contains__(item):
                value = value.replace(item, '')
                break
        return value

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

    def change_shipping_date(self, ship_date, year):
        date_list = ship_date.split(' ')
        key_list = {
            'JAN': '01',
            'FEB': '02',
            'MAR': '03',
            'APR': '04',
            'MAY': '05',
            'JUNE': '06',
            'JULY': '07',
            'AUG': '08',
            'SEPT': '09',
            'OCT': '10',
            'NOV': '11',
            'DEC': '12'
        }
        month = key_list[date_list[0]]
        day = date_list[1][:-2]
        if len(day) == 1:
            day = '0' + day
        return year + '-' + month + '-' + day + ' 00:00:00'


def gui_start():
    init_window = Tk()
    VAS = VAS_GUI(init_window)
    VAS.set_init_window()
    init_window.mainloop()


if __name__ == '__main__':
    gui_start()
