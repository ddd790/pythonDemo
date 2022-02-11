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
        self.serverName = '192.168.0.6'
        # 登陆用户名和密码
        self.userName = 'sa'
        self.passWord = 'MS_guanli09'
        # 数据库名
        self.dbName = 'ESApp1'
        # 追加的dataFrame的title
        self.add_data_title = ['自动编号', '表单状态', '季节号', '变更说明', '变更时间', '订单PO号', 'PO号批注', '面料', '成衣数量', 'ITEM', '款式', '品名', '颜色',
                               '客户需要的出厂日', '客户需要的船期', '工厂FACTORY', '里料类型', '品色号', '供应商', '品号', 'PO编号', '单耗', '需求数量', '采购数量',
                               '系统合同', '到料时间', '是否发货', '发货状态', '入库数量', '币种', '单价', '金额', '备注1', '备注2', '实际单耗', '走货数量', '样品数量',
                               '理论用量', '余量', '调拨描述', '使用数量', '实际余料', '调入', '调出', '损耗数量']

        self.po_to_purchase_li = {'自动编号': 'PO编号', '表单状态': '表单状态', '季节号': '季节号', '变更说明': '变更说明', '变更时间': '变更时间', '订单PO号': '订单PO号',
                                  'PO号批注': 'PO号批注', '面料': '面料', '数量': '成衣数量', 'ITEM': 'ITEM', '款式': '款式', '品名': '品名', '颜色': '颜色',
                                  '客户需要的出厂日': '客户需要的出厂日', '客户需要的船期': '客户需要的船期', '工厂FACTORY': '工厂FACTORY', '预计到料时间': '到料时间'}

        # 里料类型
        # self.li_type = ['前身里料品色号', '袖里料品色号', '第三种里料品色号',
        #                 '特殊用料', '裤膝', '裤子兜布', '腰里明细', '马甲前身里', '马甲后背里', '马甲后背面']
        self.li_type = ['前身里料品色号', '袖里料品色号', '第三种里料品色号',
                        '特殊用料', '裤膝', '裤子兜布', '腰里明细', '上衣口袋布成份']
        # 取【一部采购大表里料组测试_2】中已经存在的信息
        self.select_li_old_value()

        # 取【一部PO大表订单信息_明细】中表单状态是新增和修改的信息
        # self.select_po_col = ['自动编号', '表单状态', '季节号', '变更说明', '变更时间', '订单PO号', 'PO号批注', '面料', '成衣数量', 'ITEM', '款式', '品名', '颜色',
        #                       '客户需要的出厂日', '客户需要的船期', '工厂', '前身里料品色号', '袖里料品色号', '第三种里料品色号', '特殊用料', '裤膝', '裤子兜布', '腰里明细',
        #                       '马甲前身里', '马甲后背里', '马甲后背面']
        self.select_po_col = ['自动编号', '表单状态', '季节号', '变更说明', '变更时间', '订单PO号', 'PO号批注', '面料', '数量', 'ITEM', '款式', '品名', '颜色', '客户需要的出厂日',
                              '客户需要的船期', '工厂FACTORY', '预计到料时间', '前身里料品色号', '袖里料品色号', '第三种里料品色号', '特殊用料', '裤膝', '裤子兜布', '腰里明细', '上衣口袋布成份']
        self.select_po_new_value()

        # 数字列
        self.int_val = ['成衣数量', '单耗', '采购数量', '单价', '金额', '入库数量', '实际单耗', '走货数量', '理论用量', '余量', '需求数量', '样品数量', '使用数量',
                        '实际余料', '调入', '调出', '损耗数量', 'PO编号']
        # 日期列
        self.datetime_val = ['到料时间', '变更时间', '客户需要的出厂日', '客户需要的船期']
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
                if temp_po_col not in self.li_type:
                    temp_col_name = self.po_to_purchase_li[temp_po_col]
                    temp_row[temp_col_name] = getattr(row, temp_po_col)
                if temp_po_col in self.li_type and str(getattr(row, temp_po_col)).strip() != '' and str(getattr(row, temp_po_col)).strip() != ',':
                    temp_row['里料类型'] = temp_po_col
                    temp_row['品色号'] = getattr(row, temp_po_col)
                    temp_row_add = temp_row.copy()
                    dic_list.append(temp_row_add)

        new_li_df = pd.DataFrame(dic_list, columns=self.add_data_title)
        # new_li_df = pd.DataFrame(columns=self.add_data_title)
        # new_li_df = pd.concat([new_li_df, pd.DataFrame(dic_list)])

        # 循环裂变后的数据，判断是否具体是追加还是变更
        # 注：更新的数据，有个注意的地方，如果供应商或者品号变更了的，需要追加新的，旧的删除掉PO编号
        updata_item = ['表单状态', '季节号', '变更说明', '变更时间', '订单PO号', 'PO号批注', '面料', '成衣数量', 'ITEM', '款式', '品名', '颜色',
                       '客户需要的出厂日', '客户需要的船期', '工厂FACTORY', '里料类型', '品色号']
        table_data = pd.DataFrame(columns=self.add_data_title)
        add_dic_list_new = []
        # 转为字典
        dict_new_li_array = new_li_df.to_dict(orient='records')
        if self.old_all_data.empty:
            auto_no = 0
            for dict_new_fu in dict_new_li_array:
                dict_new_fu['自动编号'] = auto_no
                add_dic_list_new.append(dict_new_fu)
        else:
            po_num_new = list(set(new_li_df['PO编号']))

            # 转为字典
            dict_old_li_array = self.old_all_data.to_dict(orient='records')
            # 旧数据中有的，新数据中没有的，需要添加到最终结果中
            for dict_old_li in dict_old_li_array:
                if dict_old_li['PO编号'] not in po_num_new:
                    add_dic_list_new.append(dict_old_li)

            # 循环判断是否有需要追加和删除的数据
            for dict_new_li in dict_new_li_array:
                change_flag = True
                for dict_old_li in dict_old_li_array:
                    row_li_type = str(dict_new_li['里料类型']).strip()
                    o_row_li_type = str(dict_old_li['里料类型']).strip()
                    row_po_number = dict_new_li['PO编号']
                    o_row_po_number = dict_old_li['PO编号']
                    row_pin_number = str(dict_new_li['品色号']).strip()
                    o_row_pin_number = str(dict_old_li['品色号']).strip()
                    # 变更的数据中，品色号相同的替换，不同的追加
                    if o_row_li_type == row_li_type and o_row_po_number == row_po_number:
                        change_flag = False
                        if o_row_pin_number == row_pin_number:
                            for temp_updata_item in updata_item:
                                var_val = dict_new_li[temp_updata_item]
                                dict_old_li[temp_updata_item] = var_val
                            add_dic_list_new.append(dict_old_li)
                        else:
                            dict_new_li['自动编号'] = 0
                            add_dic_list_new.append(dict_new_li)
                            dict_old_li['PO编号'] = 0
                            add_dic_list_new.append(dict_old_li)
                # 将新增的元数据追加到table_data中
                if change_flag:
                    dict_new_li['自动编号'] = 0
                    add_dic_list_new.append(dict_new_li)

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
            table_data[time_row] = np.where(
                table_data[time_row].notnull(), table_data[time_row].dt.strftime('%Y-%m-%d %H:%M:%S'), '')
        self.table_value.append([tuple(row) for row in table_data.values])

    def select_li_old_value(self):
        # 建立连接并获取里料填写的数据（采购表）
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        strCol = ",".join(str(i) for i in self.add_data_title)
        select_sql = 'select ' + strCol + ' from 一部采购大表里料组测试_2'
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
        cursor.execute('TRUNCATE TABLE D_PurchaseLiInfo')
        # 组装插入的值
        insertValue = []
        for tabVal in self.table_value:
            insertValue += tabVal
        insertSql = 'INSERT INTO D_PurchaseLiInfo (' + sqlCol + ') VALUES ('
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
