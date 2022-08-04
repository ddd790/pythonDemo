from dataclasses import replace
import os
import pandas as pd
import datetime
import pymssql
import math
from tkinter import *


class VAS_GUI():
    # 批量获取服务器数据，进行累加操作
    def get_files(self):
        print('文件操作进行中......')
        # sql服务器名
        self.serverName = '192.168.0.6'
        # 登陆用户名和密码
        self.userName = 'sa'
        self.passWord = 'MS_guanli09'
        # 数据库名
        self.dbName = 'ESApp1'
        # 追加的dataFrame的title
        self.add_data_title = ['文件名', '品名', 'title',
                               'value', 'color', 'FinalOrder', 'SheetType', 'fileName', 'shippingDate']
        # 不需要追加的字段
        self.other_str = ['STANDARD', 'TALL', 'nan', 'Standard', 'Tall']
        networked_directory = r'\\192.168.0.6\02-业务二部资料\业务2部\2022大货\SUITSHOP\新郎装原始裁单'
        self.local_cai_detail_file = 'd:\caidan'
        # 删除目录内文件
        # if os.path.exists(self.local_cai_detail_file):
        #     shutil.rmtree(self.local_cai_detail_file)
        # os.mkdir(self.local_cai_detail_file)
        # # copy服务器的Vas_details文件到本地
        # for root, dirs, files in os.walk(networked_directory):
        #     for file in files:
        #         if str(file).__contains__('.xls') or str(file).__contains__('.xlsx'):
        #             shutil.copy(os.path.join(root, file),
        #                         self.local_cai_detail_file)

        # 循环文件，处理合并
        self.table_value = []
        # 文件名的list
        self.fileNameList = []
        # error flag
        self.errorFlag = True
        # error列表
        self.errorList = {
            'sku': '第三行的title不正确。',
            'space': '数据列中间只能有一列空白。'
        }
        # error信息
        self.errorMsg = []
        for lroot, ldirs, lfiles in os.walk(self.local_cai_detail_file):
            for lfile in lfiles:
                print('文件名：' + str(lfile).split('.')[0])
                if not lfile.__contains__('&'):
                    print('文件名错误，不包含&和时间')
                self.fileNameList.append(str(lfile).split('.')[0])
                df = pd.read_excel(os.path.join(lroot, lfile),
                                   sheet_name=None, nrows=1000, skiprows=[0, 1])
                # 读取数据转化成dataframe
                self.file_to_dataframe(df, df.keys(), str(lfile).split('.')[0])

        # 更新数据库，删除文件
        if self.errorFlag:
            self.update_db()
            for lroot, ldirs, lfiles in os.walk(self.local_cai_detail_file):
                for lfile in lfiles:
                    os.remove(self.local_cai_detail_file + '\\' + lfile)
            print('已经完成操作！' + str(datetime.datetime.now()).split('.')[0])
        else:
            for msg in self.errorMsg:
                print(msg)
        input('按回车退出 ')

    def file_to_dataframe(self, df, sheetList, fileName):
        table_data = pd.DataFrame(columns=self.add_data_title)
        for sheetName in sheetList:
            sheet_title_flag = True
            sheet_title_error_content = []
            temp_date = fileName.split('&')[1]
            shippingDate = temp_date[:4] + '-' + temp_date[4:6] + \
                '-' + temp_date[6:8] + ' 00:00:00.00'
            # sheet页名字中有0的，需要读取3列，其他情况读取2列
            if sheetName.__contains__('0'):
                for i in range(0, 100, 4):
                    try:
                        # 第0,4,8,12列的值
                        df_first_col = df[sheetName].iloc[:, i]
                        # 第1,5,9,13列的值
                        df_second_col = df[sheetName].iloc[:, i + 1]
                        # 第2,6,10,14列的值
                        df_three_col = df[sheetName].iloc[:, i + 2]
                        # title
                        title = str(df_first_col.name)
                        second_title = str(df_second_col.name)
                        three_title = str(df_three_col.name)
                        col_first_title = title.split('.')[0].split(':')[0]
                        col_second_title = second_title.split('.')[
                            0].split(':')[0]
                        col_three_title = three_title.split('.')[
                            0].split(':')[0]
                        # 第一列为空，第二列不为空,报错
                        if col_first_title == 'Unnamed' and col_second_title != 'Unnamed' and col_three_title != 'Unnamed':
                            self.errorMsg.append(
                                fileName + ':' + sheetName + '-' + self.errorList['space'])
                            break
                        for d in range(df_first_col.size):
                            # 第一列不为空，第二列不为空或者0
                            if self.is_number(df_three_col[d]) and not math.isnan(df_three_col[d]) and int(df_three_col[d]) != 0 and str(df_first_col[d]) not in self.other_str and not str(df_first_col[d]).__contains__('Total') and str(df_second_col[d]).strip() != 'nan':
                                if col_first_title.upper() != 'COLOR':
                                    self.errorFlag = False
                                    sheet_title_flag = False
                                    sheet_title_error_content.append(i + 1)
                                if col_second_title.upper() != 'SIZE':
                                    self.errorFlag = False
                                    sheet_title_flag = False
                                    sheet_title_error_content.append(i + 2)
                                if col_three_title.upper() != 'ORDER' and col_three_title.upper() != 'FINAL ORDER':
                                    self.errorFlag = False
                                    sheet_title_flag = False
                                    sheet_title_error_content.append(i + 3)
                                temp_dict = {}
                                temp_dict[self.add_data_title[0]] = fileName.split('&')[
                                    0]
                                temp_dict[self.add_data_title[1]
                                          ] = sheetName[:-1]
                                temp_dict[self.add_data_title[2]
                                          ] = col_second_title
                                temp_dict[self.add_data_title[3]
                                          ] = str(df_second_col[d]).strip()
                                temp_dict[self.add_data_title[4]
                                          ] = df_first_col[d]
                                temp_dict[self.add_data_title[5]
                                          ] = df_three_col[d]
                                temp_dict[self.add_data_title[6]] = 4 if sheetName.__contains__(
                                    'Tuxedo') else 3
                                temp_dict[self.add_data_title[7]] = fileName
                                temp_dict[self.add_data_title[8]
                                          ] = shippingDate
                                table_data = table_data.append(
                                    temp_dict, ignore_index=True)
                    except:
                        break
            else:
                for i in range(0, 100, 3):
                    try:
                        # 第0,3,6,9列的值
                        df_first_col = df[sheetName].iloc[:, i]
                        # 第1,4,7,10列的值
                        df_second_col = df[sheetName].iloc[:, i + 1]
                        # title
                        title = str(df_first_col.name)
                        second_title = str(df_second_col.name)
                        col_first_title = title.split('.')[0].split(':')[0]
                        col_second_title = second_title.split('.')[
                            0].split(':')[0]
                        # 第一列为空，第二列不为空,报错
                        if col_first_title == 'Unnamed' and col_second_title != 'Unnamed':
                            self.errorMsg.append(
                                fileName + ':' + sheetName + '-' + self.errorList['space'])
                            break
                        for d in range(df_first_col.size):
                            # 第一列不为空，第二列不为空或者0
                            if self.is_number(df_second_col[d]) and not math.isnan(df_second_col[d]) and int(df_second_col[d]) != 0 and str(df_first_col[d]) not in self.other_str and not str(df_first_col[d]).__contains__('Total'):
                                if col_first_title.upper() != 'SKU':
                                    self.errorFlag = False
                                    sheet_title_flag = False
                                    sheet_title_error_content.append(i + 1)
                                if col_second_title.upper() != 'ORDER' and col_second_title.upper() != 'FINAL ORDER':
                                    self.errorFlag = False
                                    sheet_title_flag = False
                                    sheet_title_error_content.append(i + 2)
                                temp_dict = {}
                                temp_dict[self.add_data_title[0]] = fileName.split('&')[
                                    0]
                                temp_dict[self.add_data_title[1]] = sheetName
                                temp_dict[self.add_data_title[2]] = title.split('.')[
                                    0].split(':')[0] if 'Unnamed' not in title.split('.')[
                                    0].split(':')[0] else 'SKU'
                                temp_dict[self.add_data_title[3]
                                          ] = str(df_first_col[d]).strip()
                                temp_dict[self.add_data_title[4]] = ''
                                temp_dict[self.add_data_title[5]
                                          ] = df_second_col[d]
                                temp_dict[self.add_data_title[6]] = 2 if sheetName.__contains__(
                                    'Tuxedo') else 1
                                temp_dict[self.add_data_title[7]] = fileName
                                temp_dict[self.add_data_title[8]
                                          ] = shippingDate
                                table_data = table_data.append(
                                    temp_dict, ignore_index=True)
                    except:
                        break
            if not sheet_title_flag:
                # 去重
                error_column = list(
                    set(sheet_title_error_content))
                # 排序
                error_column.sort()
                self.errorMsg.append(
                    fileName + ':' + sheetName + '-' + self.errorList['sku'])
        table_data['CreateDate'] = str(
            datetime.datetime.now()).split('.')[0]
        self.table_value.append([tuple(row) for row in table_data.values])

    def update_db(self):
        dbCol = self.add_data_title[:]
        dbCol.append('CreateDate')
        # 建立连接并获取cursor
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        # 组装删除的值
        del_tuple = tuple(self.fileNameList)
        # 删除已经存在的文件
        delSql = 'delete from D_TwoDepCai where fileName = (%s)'
        cursor.executemany(delSql, del_tuple)
        # 组装插入的值
        insertValue = []
        for tabVal in self.table_value:
            insertValue += tabVal
        insertSql = 'INSERT INTO D_TwoDepCai VALUES ('
        for colVal in dbCol:
            if colVal == 'CreateDate':
                insertSql += '%s'
            elif colVal == 'FinalOrder' or colVal == 'SheetType':
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
    VAS = VAS_GUI()
    VAS.get_files()


if __name__ == '__main__':
    gui_start()
