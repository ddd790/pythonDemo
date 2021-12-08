import os
import pandas as pd
import datetime
import pymssql
from tkinter import *


class VAS_GUI():
    # 批量获取服务器数据，进行累加操作
    def commit_batch(self):
        print('数据操作进行中......')
        # sql服务器名
        self.serverName = '192.168.0.6'
        # 登陆用户名和密码
        self.userName = 'sa'
        self.passWord = 'MS_guanli09'

        # 数据库的列
        self.dbCol = ['PurchasingSeas', 'PurchasingDoc', 'Vendor', 'VendorName', 'SupplVendor', 'SupplierName', 'LabelType', 'TLMat', 'VendorMatNo', 'Description', 'VASUs',
                      'PurchasingMethod', 'Vasby', 'VASVendor', 'VASVendorName', 'TrackingNbr', 'RequiredQty', 'RemainingQty', 'Exfactorydate', 'Changedon', 'Cr', 'shippingDate', 'PPRType', 'deleteFlag', 'CreateDate']

        # 循环文件，处理合并，并存入数据库
        self.local_vas_detail_file = r'\\192.168.0.3\01-业务一部资料\Ada\VAS-PPR'
        self.df_data = pd.DataFrame(columns=self.dbCol)

        for lroot, ldirs, lfiles in os.walk(self.local_vas_detail_file):
            for lfile in lfiles:
                if str(lroot).__contains__('CHINA'):
                    self.folder_name = 'CHINA'
                elif str(lroot).__contains__('Linde'):
                    self.folder_name = 'Linde'
                elif str(lroot).__contains__('MK'):
                    self.folder_name = 'MK'
                elif str(lroot).__contains__('Sunshine'):
                    self.folder_name = 'Sunshine'
                if str(lfile).__contains__('VAS_PPR'):
                    self.arrange_excel_data(os.path.join(lroot, lfile))

        # 追加数据
        self.batch_update_db(self.df_data, 0)
        # 查询目前数据库所有数据
        self.select_all_data()
        # 去掉创建时间
        tempCol = []
        for temp_db_col in self.dbCol:
            if temp_db_col != 'CreateDate':
                tempCol.append(temp_db_col)
        # 去掉重复的新数据
        self.old_all_data.drop_duplicates(
            subset=tempCol, keep='first', inplace=True)
        # 将去重的数据重新放入数据库中
        self.batch_update_db(self.old_all_data, 1)
        print('已经完成数据操作！')

    def batch_update_db(self, temp_data, deleteFlag):
        self.table_value = []
        # 将新的数据追加到旧数据中
        self.table_value.append([tuple(row)
                                for row in temp_data.values])
        # 更新数据库
        self.update_db(deleteFlag)

    def arrange_excel_data(self, io):
        self.dataItem = ['Purchasing Seas', 'Purchasing Doc.', 'Vendor', 'Vendor Name', 'Suppl. Vendor', 'Supplier Name', 'Label Type', 'T/L Mat.', 'Vendor Mat. No.', 'Description',
                         'VAS$', 'Purchasing Method', 'Vas by', 'VAS Vendor', 'VAS Vendor Name', 'Tracking Nbr', 'Required Qty', 'Remaining Qty', 'Ex factory date', 'Changed on', 'Cr', 'shippingDate']
        # 不满足格式条件的excel，需要转成csv，然后转成DataFrame
        new_data = self.file_to_dataframe(io)
        formartTitle = list(new_data)

        # 有的excel没有对应的列，需要将没有的赋值为空，找到对应的index
        arrangeIndex = []
        for iIdx, iVal in enumerate(self.dataItem):
            i_f = ''
            for fIdx, fTitle in enumerate(formartTitle):
                if iVal == fTitle:
                    i_f = fIdx
                    break
            arrangeIndex.append(i_f)

        # 新的dataFrame的数据
        new_df_value = []
        for newIdx, newVal in new_data.iterrows():
            detail_val = []
            for arrIndex in arrangeIndex:
                if arrIndex != '':
                    detail_val.append(str(newVal[arrIndex]).strip())
                else:
                    detail_val.append('')
            new_df_value.append(detail_val)
        new_df = pd.DataFrame(new_df_value, columns=self.dataItem)
        new_df.drop_duplicates(keep='last', inplace=True)
        new_df['PPRType'] = self.folder_name
        new_df['deleteFlag'] = 0
        new_df['CreateDate'] = str(datetime.datetime.now()).split('.')[0]
        new_df.columns = self.dbCol
        self.df_data = self.df_data.append(new_df, ignore_index=True)

    def file_to_dataframe(self, io):
        file_name = self.excel_csv_change(io, 1)
        formatCsvData = self.csv_to_dataframe(file_name)
        change_file_name = self.excel_csv_change(file_name, 2)
        return formatCsvData

    def excel_csv_change(self, io, flag):
        # 原文件后缀名
        suffix_name = '.xls' if flag == 1 else '.csv'
        # 新文件后缀名
        new_suffix_name = '.csv' if flag == 1 else '.xls'
        # flag = 1为excel2csv, flag = 2为csv2excel
        index = io.find(suffix_name)
        new_file_name = io[:index]+new_suffix_name
        os.replace(io, new_file_name)
        return new_file_name

    def csv_to_dataframe(self, io):
        for decode in ('gbk', 'utf-8', 'gb18030'):
            try:
                csv = pd.read_csv(
                    io, encoding=decode, skip_blank_lines=True, delimiter="@", header=None)
                break
            except:
                pass
        # csv中的title, 并去掉空格
        formartCsvTitle = []
        # csv数据, 并去掉空格
        formatCsv = []

        for csvIdx in range(0, len(csv)):
            tempCsvVal = str(csv.iloc[csvIdx].values[0]).replace(
                '\t', '@').split('@')
            for tempIdx in range(0, len(tempCsvVal)):
                tempCsvVal[tempIdx] = str(tempCsvVal[tempIdx]).strip()
                if tempCsvVal[tempIdx] != '0':
                    tempCsvVal[tempIdx] = tempCsvVal[tempIdx].lstrip('0')
            if csvIdx == 0:
                # csv的title数组
                formartCsvTitle = tempCsvVal
            else:
                # csv的数据的数组
                formatCsv.append(tempCsvVal)

        df = pd.DataFrame(formatCsv, columns=formartCsvTitle)
        return df

    def update_db(self, deleteFlag):
        # 建立连接并获取cursor
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, "ESApp1")
        cursor = conn.cursor()
        # 是否删除原数据
        if deleteFlag == 1:
            cursor.execute('TRUNCATE TABLE D_VasPPRInfo')
        # 组装插入的值
        insertValue = []
        for tabVal in self.table_value:
            insertValue += tabVal
        insertSql = 'INSERT INTO D_VasPPRInfo VALUES ('
        for colVal in self.dbCol:
            if colVal == 'CreateDate':
                insertSql += '%s'
            elif colVal == 'RequiredQty' or colVal == 'RemainingQty' or colVal == 'VASUs' or colVal == 'deleteFlag':
                insertSql += '%d, '
            else:
                insertSql += '%s, '
        insertSql += ')'
        cursor.executemany(insertSql, insertValue)
        conn.commit()
        conn.close()

    def select_all_data(self):
        # 建立连接并获取cursor
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, "ESApp1")
        cursor = conn.cursor()
        strCol = ",".join(str(i) for i in self.dbCol)
        select_sql = 'select ' + strCol + ' from D_VasPPRInfo'
        cursor.execute(select_sql)
        row = cursor.fetchall()
        self.old_all_data = pd.DataFrame(data=list(row), columns=self.dbCol)
        cursor.close()
        conn.close()


def gui_start():
    VAS = VAS_GUI()
    VAS.commit_batch()


if __name__ == '__main__':
    gui_start()
