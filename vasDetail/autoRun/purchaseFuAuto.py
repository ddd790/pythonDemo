import pandas as pd
import datetime
import pymssql
import numpy as np
from tkinter import *
from datetime import date, timedelta


class VAS_GUI():
    # 批量获取服务器数据，进行累加操作
    def get_files(self):
        print('数据操作进行中，请耐心等待......')
        # sql服务器名
        self.serverName = '192.168.0.11'
        # 登陆用户名和密码
        self.userName = 'sa'
        self.passWord = 'jiangbin@007'
        # 数据库名
        self.dbName = 'ESApp1'
        # 追加的dataFrame的title
        self.add_data_title = ['自动编号', 'P_KEY', '表单状态', '变更时间', '季节号', '下单日期', '进仓日期', '预计到料时间', '发货日期', '客户需要的出厂日', '工厂FACTORY',
                               '合同号', '订单PO号', '客户款式', '款式', '品名', '颜色', '内部辅料档', '辅料包类型', '面料', '供应商', '辅料名称', '辅料品号', '规格',
                               '转换比率', '采购颜色', '采购单耗', '成衣数量', '应采数量', '实际采购数量', '单价', '金额', '到货日期', '到料数量', '大货单耗', '走货数量',
                               '使用数量', '损耗数量', '调出数量', '调入数量', '余料调拨说明', '理论余料', '实际余料', '备注', '变更说明', 'PO编号', 'ITEM']

        self.po_to_purchase_fu = {'自动编号': 'PO编号', '表单状态': '表单状态', '季节号': '季节号', '变更说明': '变更说明', '变更时间': '变更时间', '订单PO号': '订单PO号',
                                  '面料': '面料', '数量': '成衣数量', 'ITEM': 'ITEM', '客户款式': '客户款式', '款式': '款式', '品名': '品名', '颜色': '颜色',
                                  '客户需要的出厂日': '客户需要的出厂日', '工厂FACTORY': '工厂FACTORY', '预计到料时间': '预计到料时间', '下单日期': '下单日期', '内部辅料档': '内部辅料档'}

        # 取【一部采购大表辅料组测试_2】中已经存在的信息
        self.select_fu_old_value()

        # 取【一部PO大表订单信息_明细】中的信息
        self.select_po_col = ['自动编号', '表单状态', '季节号', '变更说明', '变更时间', '订单PO号', '面料', '数量', 'ITEM', '客户款式', '款式', '品名',
                              '颜色', '客户需要的出厂日', '工厂FACTORY', '预计到料时间', '下单日期', '内部辅料档']
        self.select_po_new_value()

        # 数字列
        self.int_val = ['自动编号', '转换比率', '采购单耗', '成衣数量', '应采数量', '实际采购数量', '单价', '金额', '到料数量', '大货单耗', '走货数量', '使用数量',
                        '损耗数量', '调出数量', '调入数量', '理论余料', '实际余料', 'PO编号']
        # 日期列
        self.datetime_val = ['变更时间', '下单日期', '进仓日期', '预计到料时间', '发货日期',
                             '客户需要的出厂日', '到货日期']
        # 处理PO中新增和修改的数据，与里料中的原数据进行合并
        self.table_value = []
        self.merge_old_new_value()

        # 更新数据库
        print('开始更新数据！' + str(datetime.datetime.now()).split('.')[0])
        self.update_db()
        print('已经完成操作！' + str(datetime.datetime.now()).split('.')[0])

    def merge_old_new_value(self):
        # 将po中的新增和更新的数据，找到里料类型有值的（1条变N条）
        dic_list = []
        for row in self.po_new_data.itertuples():
            temp_row = {}
            for temp_po_col in self.select_po_col:
                temp_col_name = self.po_to_purchase_fu[temp_po_col]
                temp_row[temp_col_name] = getattr(row, temp_po_col)
                temp_row_add = temp_row.copy()
            dic_list.append(temp_row_add)

        new_fu_df = pd.DataFrame(dic_list, columns=self.add_data_title)

        # 如果查询的PO编号，old里面有，就更新，如果没有就新增
        updata_item = ['表单状态', '季节号', '变更说明', '变更时间', '订单PO号', '面料', '成衣数量', 'ITEM', '客户款式', '款式', '品名',
                       '颜色', '客户需要的出厂日', '工厂FACTORY', '预计到料时间', '下单日期', '内部辅料档']
        table_data = pd.DataFrame(columns=self.add_data_title)
        add_dic_list_new = []
        # 转为字典
        dict_new_fu_array = new_fu_df.to_dict(orient='records')
        if self.old_all_data.empty:
            auto_no = 0
            for dict_new_fu in dict_new_fu_array:
                dict_new_fu['自动编号'] = auto_no
                add_dic_list_new.append(dict_new_fu)
        else:
            po_num_old = list(set(self.old_all_data['PO编号']))

            dict_old_fu_array = self.old_all_data.to_dict(orient='records')
            # 旧数据全部追加
            for dict_old_fu in dict_old_fu_array:
                add_dic_list_new.append(dict_old_fu)

            # 新数据追加到后面
            update_val = {}
            for dict_new_fu in dict_new_fu_array:
                if dict_new_fu['PO编号'] not in po_num_old:
                    dict_new_fu['自动编号'] = 0
                    add_dic_list_new.append(dict_new_fu)
                else:
                    searchKey = str(dict_new_fu['PO编号'])
                    update_val[searchKey] = dict_new_fu

            # 更新已有数据
            for add_dic in add_dic_list_new:
                update_po = str(add_dic['PO编号'])
                if update_po in update_val.keys():
                    updata_val_dict = update_val[update_po]
                    for item_val in updata_item:
                        add_dic[item_val] = updata_val_dict[item_val]

        table_data = pd.DataFrame(
            add_dic_list_new, columns=self.add_data_title)

        table_data.sort_values(by='订单PO号')
        table_data['CreateDate'] = str(datetime.datetime.now()).split('.')[0]

        # 去除空的数据
        for int_val in self.add_data_title:
            if int_val in self.int_val:
                table_data[int_val].fillna(0, inplace=True)
            else:
                table_data[int_val].fillna('', inplace=True)
        # 去除日期型的NAT数据
        for time_row in self.datetime_val:
            table_data[time_row] = pd.to_datetime(
                table_data[time_row], format='%Y-%m-%d %H:%M:%S')
            table_data[time_row] = np.where(
                table_data[time_row].notnull(), table_data[time_row].dt.strftime('%Y-%m-%d %H:%M:%S'), '')
        self.table_value.append([tuple(row) for row in table_data.values])

    def select_fu_old_value(self):
        # 建立连接并获取辅料填写的数据（采购表）
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        strCol = ",".join(str(i) for i in self.add_data_title)
        select_sql = 'select ' + strCol + ' from 一部采购大表辅料组测试_2'
        cursor.execute(select_sql)
        row = cursor.fetchall()
        self.old_all_data = pd.DataFrame(
            data=list(row), columns=self.add_data_title)
        cursor.close()
        conn.close()

    def select_po_new_value(self):
        start_time = '2022-01-01'
        tomorrow = (date.today() + timedelta(days=1)).strftime("%Y-%m-%d")
        # 建立连接并获取订单信息中的数据
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        strCol = ",".join(str(i) for i in self.select_po_col)
        select_sql = "select " + strCol + \
            " from 一部PO大表订单信息_明细 where 下单日期 BETWEEN '" + \
            start_time + "' AND '" + tomorrow + "'"
        cursor.execute(select_sql)
        row = cursor.fetchall()
        self.po_new_data = pd.DataFrame(
            data=list(row), columns=self.select_po_col)
        cursor.close()
        conn.close()

    def update_db(self):
        dbCol = self.add_data_title[:]
        dbCol.append('CreateDate')
        sqlCol = ",".join(str(i) for i in dbCol)
        # 建立连接并获取cursor
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        cursor.execute('TRUNCATE TABLE D_PurchaseFuInfo')
        # 组装插入的值
        insertValue = []
        for tabVal in self.table_value:
            insertValue += tabVal
        insertSql = 'INSERT INTO D_PurchaseFuInfo (' + sqlCol + ') VALUES ('
        for colVal in dbCol:
            if colVal == 'CreateDate':
                insertSql += '%s'
            elif colVal in self.int_val:
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
