import os
import pandas as pd
import datetime
import pymssql
import numpy as np
import pdfplumber
import shutil
import time
from dateutil import parser

from tkinter import *


class VAS_GUI():
    # 批量获取服务器数据，进行累加操作
    def get_files(self):
        print('文件操作进行中......')
        # sql服务器名
        self.serverName = '192.168.0.11'
        # 登陆用户名和密码
        self.userName = 'sa'
        self.passWord = 'jiangbin@007'
        # 数据库名
        self.dbName = 'ESApp1'
        # 追加的dataFrame的title
        self.add_data_title = ['FileName', 'FileNameJ', 'Version', 'Item', 'PO', 'Material', 'MaterialNo', 'Description', 'Component', 'FabricNo', 'Qty', 'ZROH',
                               'FabricDis', 'ExfactDate', 'ShipDate', 'Season', 'Price', 'Brand', 'District', 'Via', 'Style', 'ContanctPerson', 'Note',
                               'Rmk', 'HSCode', 'CreateDate', 'DeliverDate']
        # 数字类型的字段
        self.number_item = ['Qty', 'Price']
        # 备注的公司名
        self.company_rmk = ['MOTIVES INTERNATIONAL LIMITED',
                            'MOTIVES INTERNATIONAL HK LIMITED', 'MOTIVES CHINA LIMITED']
        # keyword集合
        self.keyword = {'cloth_content': 'Cloth Content: ', 'zroh': 'ZROH: ',
                        'color_des': ' Color Description: ', 'hscode': 'HS Code: ', 'season': 'Season: ',
                        'currency': 'Currency ', 'fax': 'FAX # :', 'person_tel': 'person/Telephone', 'ext': 'EXT:', 'meter': 'Meter',
                        'first_word': 'Please deliver to:',
                        'middle_word': 'We require an order acknowledgement for the following items:',
                        'end_word': 'It requires the following components:'}
        # 运输方式
        self.Via = ['SEA', 'AIRV', 'AIRP']
        networked_directory = r'\\192.168.0.3\01-业务一部资料\=14785212\PEERLESS\国内埃塞柬埔寨订单信息\临时'
        self.local_pdf_detail_file = 'd:\peerlessPdf'

        # 循环文件，处理合并
        self.division_word_one = '''___________________________________________________________________________________________________________________
Item Material No.             Description    Exfact.date  PO Ship date      Delivr. date
Order Qty / Unit             Price/unit   Net value Via   Weight per unit
___________________________________________________________________________________________________________________'''
        self.division_word_two = '___________________________________________________________________________________________________________________'
        self.division_word_three = '  division_word_three  '

        # 删除目录内文件
        if os.path.exists(self.local_pdf_detail_file):
            shutil.rmtree(self.local_pdf_detail_file)
        os.mkdir(self.local_pdf_detail_file)
        # copy服务器的TRIMLIST文件到本地
        for root, dirs, files in os.walk(networked_directory):
            if root.__contains__('20'):
                for file in files:
                    if str(file).__contains__('PO-') and (str(file).__contains__('.pdf') or str(file).__contains__('.PDF')):
                        shutil.copy2(os.path.join(root, file), self.local_pdf_detail_file)
        # 保留相同文件中最大的记录

        # 查询已存在的记录
        self.select_po_old_value()
        # 已经存在的文件名(去重)
        file_name_old = list(set(self.old_all_data['FileName']))

        self.pdf_data_val = []
        # 文件名的list
        self.fileNameList = []
        for lroot, ldirs, lfiles in os.walk(self.local_pdf_detail_file):
            for file in lfiles:
                file_name = str(file).split('.')[0]
                # print('文件名：' + str(file).split('-V')[0])
                self.fileNameList.append(str(file).split('.')[0])
                if (str(file).__contains__('.pdf') or str(file).__contains__('.PDF')) and file_name not in file_name_old:
                    mtime = parser.parse(time.ctime(os.path.getmtime(os.path.join(lroot, file))))
                    # ctime = time.ctime(os.path.getctime(
                    #     os.path.join(lroot, lfile)))
                    create_time = mtime.strftime('%Y-%m-%d %H:%M:%S')
                    self.file_to_dataframe_pdfplumber(file, create_time)
                    # try:
                    # except:
                    #     continue
        # 将pdf结果转成dataFrame
        table_data = pd.DataFrame(self.pdf_data_val, columns=self.add_data_title)
        # table_data['CreateDate'] = str(datetime.datetime.now()).split('.')[0]
        # table_data['CreateDate'] = '2022-01-01 00:00:00'
        self.table_value = []
        self.table_value.append([tuple(row) for row in table_data.values])

        # 更新数据库
        self.update_db()
        print('已经完成操作！' + str(datetime.datetime.now()).split('.')[0])

    def compare_pdf_file(self):
        # 遍历目录，留下最新的文件
        fileNameList = []
        tempDelMap = {}
        for eroot, edirs, efiles in os.walk(self.local_pdf_detail_file):
            for name in efiles:
                fileName = os.path.splitext(name)[0]
                nameList = fileName.split('-')
                nameKey = "-".join(str(i) for i in nameList[:-1])
                if nameKey not in fileNameList:
                    fileNameList.append(nameKey)
                    tempDelMap[nameKey] = name
                else:
                    tempDelFile = tempDelMap[nameKey]
                    tempDelFileNameList = os.path.splitext(tempDelFile)[
                        0].split('-')
                    # 判断是否有括号
                    if str(self.get_value_two_word(nameList[-1], 'V', None)).__contains__('(') and not str(self.get_value_two_word(tempDelFileNameList[-1], 'V', None)).__contains__('('):
                        os.remove(os.path.join(eroot, tempDelFile))
                        tempDelMap[nameKey] = name
                    elif str(self.get_value_two_word(nameList[-1], 'V', None)).__contains__('(') and str(self.get_value_two_word(tempDelFileNameList[-1], 'V', None)).__contains__('('):
                        if int(self.get_value_two_word(nameList[-1], '(', ')')) > int(self.get_value_two_word(tempDelFileNameList[-1], '(', ')')):
                            os.remove(os.path.join(eroot, tempDelFile))
                            tempDelMap[nameKey] = name
                        else:
                            os.remove(os.path.join(eroot, name))
                    elif not str(self.get_value_two_word(nameList[-1], 'V', None)).__contains__('(') and str(self.get_value_two_word(tempDelFileNameList[-1], 'V', None)).__contains__('('):
                        os.remove(os.path.join(eroot, name))
                    # 版本号判断
                    elif int(self.get_value_two_word(nameList[-1], 'V', None)) > int(self.get_value_two_word(tempDelFileNameList[-1], 'V', None)):
                        os.remove(os.path.join(eroot, tempDelFile))
                        tempDelMap[nameKey] = name
                    else:
                        os.remove(os.path.join(eroot, name))

    def file_to_dataframe_pdfplumber(self, fileName, create_time):
        # 文件名
        tx_fileName = str(fileName).split('.')[0]
        # print('新增文件名：' + tx_fileName)
        # PO
        tx_po = tx_fileName.split('-')[1]
        pdfreader = pdfplumber.open(
            self.local_pdf_detail_file + '\\' + fileName)
        tx_val = []
        # 循环读取pdf内容
        for index in range(len(pdfreader.pages)):
            pageReader = pdfreader.pages[index]
            pageObj = pageReader.extract_text()  # 获取内容
            tx_val.append(pageObj)
        # 每页中的元素拼接
        tx = '\n'.join(tx_val)
        # 品牌
        tx_pinpai = self.get_value_two_word(
            tx, self.keyword['currency'], self.keyword['middle_word'])
        tx_pinpai = tx_pinpai.replace('\n', '|')
        pinpai_list = tx_pinpai.split('|')
        tx_pinpai = pinpai_list[1]
        # 批注
        tx_cc = ''
        if (len(pinpai_list) > 3):
            tx_cc = pinpai_list[2].strip()
            tx_cc = tx_cc.replace(self.division_word_two, '')
        # 目的地
        tx_des = 'USA'
        tx_des_info = self.get_value_two_word(
            tx, self.keyword['first_word'], self.keyword['fax'])
        if tx_des_info.strip() == '':
            tx_des_info = self.get_value_two_word(
                tx, self.keyword['first_word'], self.keyword['person_tel'])
        if str(tx_des_info).__contains__('8888'):
            tx_des = 'CDN'
        elif str(tx_des_info).__contains__('MEXICO'):
            tx_des = 'MEXICO'
        # 联系人
        tx_person = self.get_value_two_word(
            tx, self.keyword['person_tel'], self.keyword['ext'])
        # print('-----------------------------------------------------------')
        tx_person_list = tx_person.split('/')[-2]
        tx_person_list = tx_person_list.split(' ')
        tx_person = tx_person_list[-2] + ' ' + tx_person_list[-1]
        # Rmk
        tx_Rmk = ''
        if tx.find(self.company_rmk[0]) >= 0:
            tx_Rmk = self.company_rmk[0]
        elif tx.find(self.company_rmk[1]) >= 0:
            tx_Rmk = self.company_rmk[1]
        else:
            tx_Rmk = self.company_rmk[2]
        # 去掉无用的数据
        tx = tx.replace(self.division_word_one, '')
        tx = tx.replace('\n', '|')
        tx = tx.replace('  ', '|')
        tx = self.deduplicate(tx, '|')
        detail_count = tx.count(self.keyword['end_word'])
        detail_info_list = []
        if detail_count > 0:
            for i in range(detail_count):
                # 取txt_str文字中,两个字符串中的字符
                temp_zroh = tx[self.findSubStrIndex(self.keyword['end_word'], tx, i + 1) + len(
                    self.keyword['end_word']):self.findSubStrIndex(self.keyword['meter'], tx, i + 1)]
                full_txt = tx[self.findSubStrIndex(self.division_word_two, tx, i + 1) + len(
                    self.division_word_two):self.findSubStrIndex(self.keyword['end_word'], tx, i + 1)]
                full_txt = full_txt.replace('| ', '|')
                full_txt = self.deduplicate(full_txt, '|')
                full_txt = full_txt + '|ZROH: ' + temp_zroh.split('|')[1] + ' Color Description: ' + temp_zroh.split('|')[2]
                detail_info_list.append(full_txt.strip())
        else:
            detail_other_count = tx.count(self.division_word_two)
            for i in range(detail_other_count):
                full_txt = tx[self.findSubStrIndex(self.division_word_two, tx, i + 1) + len(
                    self.division_word_two):]
                full_txt = full_txt.replace('| ', '|')
                full_txt = self.deduplicate(full_txt, '|')
                detail_info_list.append(full_txt.strip())
        if len(detail_info_list) == 0:
            print(fileName)
        if len(detail_info_list) > 0:
            for i in range(len(detail_info_list)):
                temp_info_list = detail_info_list[i].split('|')
                detail_info = []
                # 文件名
                detail_info.append(tx_fileName)
                # 文件名简写
                detail_info.append(str(tx_fileName).split('-V')[0])
                # 版本
                detail_info.append(str(tx_fileName).split('-V')[1])
                # 追加描述(描述占2个元素)
                if temp_info_list[3].strip()[0] == '2':
                    temp_info_list.insert(3, '')
                if temp_info_list[5].count('.') != 2:
                    temp_info_list.insert(5, temp_info_list[4].strip())
                # 追加Delivr. date
                deliver_date = str(temp_info_list[5]).strip()
                if temp_info_list[6].count('.') == 2:
                    # 追加Delivr. date
                    deliver_date = str(temp_info_list[6]).strip()
                    del temp_info_list[6]
                else:
                    temp_info_list[5] = temp_info_list[4]
                if not str(temp_info_list[7]).__contains__(','):
                    temp_info_list.pop(7)
                if not str(temp_info_list[11]).__contains__('ZROH'):
                    temp_info_list.insert(11, temp_info_list[-1])
                # HScode
                if not str(temp_info_list[12]).__contains__('HS Code'):
                    temp_info_list[11] = temp_info_list[11] + \
                        ' ' + temp_info_list[12]
                    temp_info_list.pop(12)
                # season
                if not str(temp_info_list[13]).__contains__('Season: '):
                    temp_info_list.pop(13)
                    temp_info_list[13] = 'Season: ' + temp_info_list[13]
                # item (如果是00000，则取Season: 后面的)
                if str(temp_info_list[0]) == '00000':
                    if str(temp_info_list[14])[0] != '0':
                        tx_cc = temp_info_list[14]
                        detail_info.append(str(temp_info_list[15]))
                    else:
                        if str(temp_info_list[14])[-1] == ',':
                            detail_info.append(
                                str(temp_info_list[14]) + str(temp_info_list[15]))
                        else:
                            detail_info.append(str(temp_info_list[14]))
                else:
                    detail_info.append(temp_info_list[0])
                # PO
                detail_info.append(tx_po)
                # 完整款
                detail_info.append(self.get_value_two_word(
                    temp_info_list[2].strip(), None, ' '))
                # 款号
                detail_info.append(self.get_value_two_word(
                    temp_info_list[1], None, '-'))
                # Description
                detail_info.append(temp_info_list[3].strip())
                # 成分
                detail_info.append(self.get_value_two_word(
                    temp_info_list[10], self.keyword['cloth_content'], None))
                # 面料号
                detail_info.append(self.get_value_two_word(
                    temp_info_list[1], '-', None))
                # 数量
                detail_info.append(
                    int("".join(list(filter(str.isdigit, temp_info_list[6])))))
                # ZROH
                detail_info.append(self.get_value_two_word(
                    temp_info_list[11], self.keyword['zroh'], self.keyword['color_des']))
                # 面料描述
                detail_info.append(self.get_value_two_word(
                    temp_info_list[11], self.keyword['color_des'], None))
                # 排产日
                detail_info.append(self.str2datatime(
                    temp_info_list[4].strip()))
                # 船期
                detail_info.append(self.str2datatime(
                    temp_info_list[5].strip()))
                # 季节
                detail_info.append(self.get_value_two_word(
                    temp_info_list[13], self.keyword['season'], None)[:3])
                # 单价
                detail_info.append(
                    float(temp_info_list[7].replace(',', '.').strip()))
                # 品牌
                detail_info.append(tx_pinpai)
                # 目的地
                detail_info.append(tx_des)
                # 运输方式
                tx_via = ''
                if temp_info_list[8].strip().__contains__(self.Via[1]):
                    tx_via = self.Via[1]
                elif temp_info_list[8].strip().__contains__(self.Via[2]):
                    tx_via = self.Via[2]
                else:
                    tx_via = self.Via[0]
                detail_info.append(tx_via)
                # 款式类型
                detail_info.append(
                    temp_info_list[2].strip() + temp_info_list[3])
                # 联系人
                detail_info.append(tx_person.strip())
                # 批注
                detail_info.append(tx_cc)
                # Rmk
                detail_info.append(tx_Rmk)
                # HSCode
                detail_info.append(self.get_value_two_word(
                    temp_info_list[12], self.keyword['hscode'], None))
                # CreateDate
                detail_info.append(create_time)
                # deliver_date
                detail_info.append(self.str2datatime(deliver_date))
                self.pdf_data_val.append(detail_info)

    # 获取一个字符串中两个字母中间的值(one为None时从第一位取, two为None时取到最后)
    def get_value_two_word(self, txt_str, one, two):
        if one == None:
            return txt_str[:txt_str.find(two)]
        if two == None:
            return txt_str[txt_str.find(one) + len(one):]
        return txt_str[txt_str.find(one) + len(one):txt_str.find(two)]

    # 删除重复字符
    def deduplicate(self, string, char):
        return char.join([substring for substring in string.strip().split(char) if substring])

    # 找字符串substr在str中第time次出现的位置
    def findSubStrIndex(self, substr, str, time):
        times = str.count(substr)
        if (times == 0) or (times < time):
            pass
        else:
            i = 0
            index = -1
            while i < time:
                index = str.find(substr, index + 1)
                i += 1
            return index

    # 字符串转日期的字符串
    def str2datatime(self, str_word):
        return str(str_word).replace('.', '-') + ' 00:00:00'

    def update_db(self):
        dbCol = self.add_data_title[:]
        # 建立连接并获取cursor
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        # 组装删除的值
        # del_tuple = tuple(self.fileNameList)
        # 删除已经存在的文件
        # delSql = 'delete from D_Peerless_Order where 文件名 = (%s)'
        # cursor.executemany(delSql, del_tuple)
        delSql = 'TRUNCATE TABLE D_Peerless_Order'
        cursor.execute(delSql)
        # 组装插入的值
        insertValue = []
        for tabVal in self.table_value:
            insertValue += tabVal
        insertSql = 'INSERT INTO D_Peerless_Order (' + (
            ",".join(str(i) for i in dbCol)) + ') VALUES ('
        for colVal in dbCol:
            if colVal == 'DeliverDate':
                insertSql += '%s'
            elif colVal in self.number_item:
                insertSql += '%d, '
            else:
                insertSql += '%s, '
        insertSql += ')'
        cursor.executemany(insertSql, insertValue)
        conn.commit()
        conn.close()

    def select_po_old_value(self):
        # 建立连接并获取PO数据
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        strCol = ",".join(str(i) for i in self.add_data_title)
        select_sql = 'select ' + strCol + ' from D_Peerless_Order'
        cursor.execute(select_sql)
        row = cursor.fetchall()
        self.old_all_data = pd.DataFrame(
            data=list(row), columns=self.add_data_title)
        cursor.close()
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
