import os
import re
import pandas as pd
import datetime
import pymssql
from tkinter import *


class VAS_GUI():
    # 批量获取服务器数据，进行累加操作
    def commit_batch(self):
        print('数据操作进行中......' + str(datetime.datetime.now()).split('.')[0])
        self.local_vas_detail_file = 'd:\excelVasDetailFile'
        # 循环本地临时文件，处理合并，并存入数据库
        self.table_value = []
        for lroot, ldirs, lfiles in os.walk(self.local_vas_detail_file):
            for lfile in lfiles:
                self.read_excel(os.path.join(lroot, lfile))

        # 更新数据库
        self.update_db()
        print('已经完成数据操作！' + str(datetime.datetime.now()).split('.')[0])

    def compare_xls_file(self):
        # 遍历目录，留下最新的文件
        fileNameList = []
        tempDelMap = {}
        for eroot, edirs, efiles in os.walk(self.local_vas_detail_file):
            for name in efiles:
                fileName = os.path.splitext(name)[0]
                nameList = fileName.split('_')
                nameKey = nameList[3]
                if nameKey not in fileNameList:
                    fileNameList.append(nameKey)
                    tempDelMap[nameKey] = name
                else:
                    tempDelFile = tempDelMap[nameKey]
                    tempDelFileNameList = os.path.splitext(tempDelFile)[
                        0].split('_')
                    if int(nameList[4][0:8]) > int(tempDelFileNameList[4][0:8]):
                        os.remove(os.path.join(eroot, tempDelFile))
                        tempDelMap[nameKey] = name
                    else:
                        if len(nameList) > 5:
                            if int(tempDelFileNameList[5].split('(')[0]) > int(nameList[5].split('(')[0]):
                                os.remove(os.path.join(eroot, name))
                            else:
                                os.remove(os.path.join(eroot, tempDelFile))
                                tempDelMap[nameKey] = name
                        else:
                            # 有括号留括号，删除没有括号的
                            if str(tempDelFile).__contains__('(') and not str(name).__contains__('('):
                                os.remove(os.path.join(eroot, name))
                            elif str(name).__contains__('(') and not str(tempDelFile).__contains__('('):
                                os.remove(os.path.join(eroot, tempDelFile))
                                tempDelMap[nameKey] = name
                            elif str(name).__contains__('(') and str(tempDelFile).__contains__('('):
                                reName = re.findall(r'[(](.*?)[)]', name)[0]
                                reTempDelFile = re.findall(
                                    r'[(](.*?)[)]', tempDelFile)[0]
                                if int(reName) > int(reTempDelFile):
                                    os.remove(os.path.join(eroot, tempDelFile))
                                    tempDelMap[nameKey] = name
                                else:
                                    os.remove(os.path.join(eroot, name))
                            else:
                                if int(nameList[4]) > int(tempDelFileNameList[4]):
                                    os.remove(os.path.join(eroot, tempDelFile))
                                    tempDelMap[nameKey] = name
                                else:
                                    os.remove(os.path.join(eroot, name))

    def read_excel(self, io):
        self.dataItem = ['Purchasing Document', 'Item', 'Ex factory date', 'Split from',
                         'PO ship date', 'Stock Category', 'Material', 'Grid Value', 'Quantity', 'Priority', 'Package ID', 'BG', 'CT', 'CL', 'CS', 'HB', 'HG', 'HT', 'SL', 'SB', 'SH', 'IB', 'OB', 'OI', 'OO', 'CI', 'PT', 'PB', 'SF', 'SZ', 'ST', 'TK', 'UC', 'NL', 'LS', 'PU', 'MO', 'MH', 'MI', 'TS', 'PF', 'PI', 'PL', 'PP', 'TP', 'FSCP', 'FSEP', 'FSSP']
        # 对excel读取的数据进行整理，整理成符合要求的格式（按照dataItem中的列进行排列）
        data = self.arrange_excel_data(io, self.dataItem)
        self.table_value.append([tuple(row) for row in data.values])

    def update_db(self):
        dbCol = ['PO', 'Item', 'ExFactoryDate', 'SplitFrom', 'POShipDate', 'StockCategory', 'Material', 'GridValue', 'Quantity',
                 'Priority', 'PackageID', 'BG', 'CT', 'CL', 'CS', 'HB', 'HG', 'HT', 'SL', 'SB', 'SH', 'IB', 'OB', 'OI', 'OO', 'CI', 'PT', 'PB', 'SF', 'SZ', 'ST', 'TK', 'UC', 'NL', 'LS', 'PU', 'MO', 'MH', 'MI', 'TS', 'PF', 'PI', 'PL', 'PP', 'TP', 'FSCP', 'FSEP', 'FSSP', 'CreateDate']
        # sql服务器名
        serverName = '192.168.0.11'
        # 登陆用户名和密码
        userName = 'sa'
        passWord = 'jiangbin@007'
        # 建立连接并获取cursor
        conn = pymssql.connect(serverName, userName, passWord, "ESApp1")
        cursor = conn.cursor()
        cursor.execute('TRUNCATE TABLE D_VasDetailInfo')
        # 组装插入的值
        insertValue = []
        for tabVal in self.table_value:
            insertValue += tabVal
        insertSql = 'INSERT INTO D_VasDetailInfo VALUES ('
        for colVal in dbCol:
            if colVal == 'CreateDate':
                insertSql += '%s'
            elif colVal == 'Quantity':
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

    def arrange_excel_data(self, io, dataItem):
        # 不满足格式条件的excel，需要转成csv，然后转成DataFrame
        new_data = self.file_to_dataframe(io)
        formartTitle = list(new_data)

        # 有的excel没有对应的列，需要将没有的赋值为空，找到对应的index
        arrangeIndex = []
        for iIdx, iVal in enumerate(dataItem):
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
                    detail_val.append(newVal[arrIndex])
                else:
                    detail_val.append('')
            new_df_value.append(detail_val)
        new_df = pd.DataFrame(new_df_value, columns=dataItem)
        new_df['CreateDate'] = str(datetime.datetime.now()).split('.')[0]
        return new_df

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
        csv = ''
        for decode in ('gbk', 'utf-8', 'gb18030'):
            try:
                csv = pd.read_csv(
                    io, encoding=decode, skip_blank_lines=True, delimiter=";", header=None)
                break
            except:
                pass
        # csv中的title, 并去掉空格
        formartCsvTitle = []
        # csv数据, 并去掉空格
        formatCsv = []

        for csvIdx in range(0, len(csv)):
            tempCsvVal = str(csv.iloc[csvIdx].values[0]).replace(
                '\t', ';').split(';')
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


def gui_start():
    VAS = VAS_GUI()
    VAS.commit_batch()


if __name__ == '__main__':
    gui_start()
