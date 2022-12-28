import os
import shutil
import pandas as pd
import datetime
import pymssql
import numpy as np
from tkinter import *


class VAS_GUI():
    # 批量获取服务器数据，进行累加操作
    def get_files(self):
        print('数据操作进行中......')
        # 追加的dataFrame的title
        self.add_data_title = ['version', '订单PO号', '款式缩写', '面料', '英文品名', '辅料表面料描述', '是否半里子', '前身里代号', '前身里料品色号', '袖里代号', '袖里料品色号',
                               '后袖笼拼接料代码', '后袖笼拼接料品色号', '第三种里料代码', '第三种里料品色号', '扣代号', '国外扣供应商品号色号', '内扣或两种以上扣代号', '内扣或两种以上扣型号',
                               '内部辅料档', '胸衬', '领呢', '领座', '上衣口袋布成份', '肩垫', '袖笼条', '防抻条', '第二种以上小面料', '特殊用料', '订单COMMENTS', '钎子', '拉链', '裤膝',
                               '裤子兜布', '腰里代码', '腰里明细', '腰衬', '腰面夹牙腰面包条', '马甲前身里', '马甲后背里', '马甲后背面', '三角牌扣', '三角牌扣供应商品号色号']
        # 特殊用料的排除列
        self.spe_data = ['前身里料品色号', '袖里料品色号', '后袖笼拼接料品色号',
                         '第三种里料品色号', '国外扣供应商品号色号', '内扣或两种以上扣型号', '上衣口袋布成份']
        # 根据勤哲的key匹配对应trimList中的key和value
        self.trimList_key_to_qizhe_key()
        networked_directory = r'\\192.168.0.3\01-业务一部资料\=14785212\PEERLESS\国内埃塞柬埔寨订单信息\临时'
        self.local_vas_detail_file = 'd:\excelTrimListFile'
        # 删除目录内文件
        if os.path.exists(self.local_vas_detail_file):
            shutil.rmtree(self.local_vas_detail_file)
        os.mkdir(self.local_vas_detail_file)
        # copy服务器的TRIMLIST文件到本地
        for root, dirs, files in os.walk(networked_directory):
            for file in files:
                if str(file).__contains__('TRIMLIST') and (str(file).__contains__('.xls') or str(file).__contains__('.xlsx')):
                    shutil.copy(os.path.join(root, file), self.local_vas_detail_file)
        # 保留相同文件中最大的记录
        self.compare_xls_file()

        # 循环本地临时文件，处理合并
        self.table_value = []
        self.delPoList = []
        self.allDataKeys = []
        for lroot, ldirs, lfiles in os.walk(self.local_vas_detail_file):
            for lfile in lfiles:
                self.file_to_dataframe(os.path.join(lroot, lfile), str(lfile).split('-')[2].split('.')[0])
        # 更新数据库
        self.update_db()
        print('已经完成计算操作！')

    def compare_xls_file(self):
        # 遍历目录，留下最新的文件
        fileNameList = []
        tempDelMap = {}
        for eroot, edirs, efiles in os.walk(self.local_vas_detail_file):
            for name in efiles:
                fileName = os.path.splitext(name)[0]
                nameList = fileName.split('-')
                nameKey = nameList[1]
                if nameKey not in fileNameList:
                    fileNameList.append(nameKey)
                    tempDelMap[nameKey] = name
                else:
                    tempDelFile = tempDelMap[nameKey]
                    tempDelFileNameList = os.path.splitext(tempDelFile)[0].split('-')
                    if int(nameList[2][1:]) > int(tempDelFileNameList[2][1:]):
                        os.remove(os.path.join(eroot, tempDelFile))
                        tempDelMap[nameKey] = name
                    else:
                        os.remove(os.path.join(eroot, name))

    def file_to_dataframe(self, io, version):
        # 读取文件版本最大的
        excelData = pd.read_excel(io, header=None, keep_default_na=False)
        # csv中的title
        formartExcelTitle = []
        # csv数据
        formatExcelvalue = excelData.values
        for csvIdx in range(0, len(excelData.values[0])):
            formartExcelTitle.append(csvIdx)
        df = pd.DataFrame(formatExcelvalue, columns=formartExcelTitle)
        # title
        excelTitle = df[0]
        # 正常数据，偶数列的数据
        dataVal = []
        # 描述数据，基数列的数据
        disVal = []
        for tempIndex in formartExcelTitle:
            str_arr = df[tempIndex].values
            for arr_i in range(len(str_arr)):
                str_arr[arr_i] = str(str_arr[arr_i]).replace(
                    '=', '').replace('"', '')
            if tempIndex != 0 and tempIndex % 2 == 0:
                disVal.append(str_arr)
            elif tempIndex % 2 != 0:
                dataVal.append(str_arr)

        valueDf = pd.DataFrame(dataVal, columns=excelTitle)
        valueDf['version'] = version
        disDf = pd.DataFrame(disVal, columns=excelTitle)
        disDf['Purchasing Document'] = valueDf['Purchasing Document'][0]
        allDataKey = valueDf['Purchasing Document'][0] + '^_^' + version
        # 所有最大版本的key数据（PO+版本）
        if allDataKey not in self.allDataKeys:
            self.allDataKeys.append(allDataKey)
        # 删除前一版本的记录(正常是前一版本的，但是有不连续的断更，按时先删除所有该PO的信息后，再插入新的)
        # 在数据库操作时实现，这里先注释掉。
        # if version != 'V1':
        #     delVersion = 'V' + str((int(version.replace('V', '')) - 1))
        #     delStr = valueDf['Purchasing Document'][0] + '^_^' + delVersion
        #     if delStr not in self.delPoList:
        #         self.delPoList.append(delStr)
        disDf['version'] = version

        # 对excel读取的数据进行整理，整理成符合要求的格式
        self.arrange_excel_data(valueDf, disDf)

    def arrange_excel_data(self, valueDf, disDf):
        # 根据勤哲的key匹配对应trimList中的key和value
        arrangeVal = []
        for idx, value in valueDf.iterrows():
            itemVal = []
            # 前身里代号，袖里代号是否为空的flag
            frontFlag = False
            sleeveFlag = False
            # 特殊用料，需要判断前面是否出现过（self.spe_data）
            tempSpeData = []
            for qinzheTitle in self.add_data_title:
                tempDic = self.arrange_qinzhe_key[qinzheTitle]
                # 临时数组存放对应的值，并后期进行去重及逗号处理
                tempValArray = []
                for key, val in tempDic.items():
                    try:
                        setVal = str(value[key]).strip()
                        setDisVal = str(disDf.loc[idx][key]).strip()
                        if val == 1:
                            if setVal != 'NONE':
                                tempValArray.append(setVal)
                        elif val == 2:
                            if setDisVal != 'NONE':
                                tempValArray.append(setDisVal)
                        elif val == 3:
                            if setVal != '' or setDisVal != '':
                                tempSetVal = setVal + ',' + setDisVal
                                tempSetVal = tempSetVal.strip(',')
                                tempValArray.append(tempSetVal)
                        elif val == 4:
                            if setDisVal != '':
                                tempValArray.append(setDisVal)
                            elif setVal != '':
                                tempValArray.append(setVal)
                    except:
                        tempValArray.append('')
                # 去掉空格和重复
                valArray = [i for i in tempValArray if i != '']
                valArraySet = list(set(valArray))
                valArraySet.sort(key=valArray.index)
                # 前身里代号，袖里代号如果有空的，第三种里料代码选第二个
                if qinzheTitle == '前身里代号' and len(valArraySet) == 0:
                    frontFlag = True
                if qinzheTitle == '袖里代号' and len(valArraySet) == 0:
                    sleeveFlag = True
                # 第三种里料，大于2个的才取
                if qinzheTitle == '第三种里料代码' or qinzheTitle == '第三种里料品色号':
                    if frontFlag and sleeveFlag and len(valArraySet) > 0:
                        valArraySet = valArraySet[0:]
                    elif (frontFlag or sleeveFlag) and len(valArraySet) > 1:
                        valArraySet = valArraySet[1:]
                    elif len(valArraySet) > 2:
                        valArraySet = valArraySet[2:]
                    else:
                        valArraySet = []
                # 特殊用料，需要判断前面是否出现过（self.spe_data）
                if qinzheTitle in self.spe_data:
                    tempSpeData = np.append(tempSpeData, valArraySet)
                if qinzheTitle == '特殊用料':
                    for temp in tempSpeData:
                        if temp in valArraySet:
                            valArraySet.remove(temp)
                # 逗号分割值
                tempValString = ",".join(str(i) for i in valArraySet)
                # 内部辅料档去掉*号
                if qinzheTitle == '内部辅料档':
                    tempValString = tempValString.replace('*', '')
                itemVal.append(tempValString)
            arrangeVal.append(itemVal)

        table_data = pd.DataFrame(arrangeVal, columns=self.add_data_title)
        table_data['CreateDate'] = str(datetime.datetime.now()).split('.')[0]
        self.table_value.append([tuple(row) for row in table_data.values])

    def trimList_key_to_qizhe_key(self):
        # trimlist里的列，对应需要取信息的值，还是信息的描述。
        # 1为取信息的值，2为取描述的值，3为两者都取，4为取描述的值，如果描述没有，则取信息的值
        trim_list_title_list = []
        trim_list_title_list.append({'version': 1})
        trim_list_title_list.append({'Purchasing Document': 1})
        trim_list_title_list.append({'Style': 1})
        trim_list_title_list.append({'Fabric': 1})
        trim_list_title_list.append({'Fabric': 2})
        trim_list_title_list.append({'Fabric Description': 1})
        trim_list_title_list.append({'LINING PATTERN TYPE': 1})
        trim_list_title_list.append(
            {'FRONT LINING': 1, 'FRONT TOP LINING': 1, 'FRONT BOTTOM LINING': 1})
        trim_list_title_list.append(
            {'FRONT LINING': 2, 'FRONT TOP LINING': 2, 'FRONT BOTTOM LINING': 2})
        trim_list_title_list.append({'SLEEVE LINING': 1})
        trim_list_title_list.append({'SLEEVE LINING': 2})
        trim_list_title_list.append({'GUSSET': 1})
        trim_list_title_list.append({'GUSSET': 2})
        trim_list_title_list.append(
            {'FRONT LINING': 1, 'BACK LINING': 1, 'SIDE BODY LINING': 1, 'SLEEVE LINING': 1,
             'TAB LINING': 1, 'INSIDE BESOM LINING': 1, 'INSIDE BESOM FACING LINING': 1, 'INSIDE CELL LINING': 1,
             'INSIDE CELL FACING LINING': 1, 'INSIDE PEN LINING': 1, 'INSIDE PEN FACING LINING': 1,
             'FLAP LINING COAT': 1, 'INSIDE FLAP LIN  ON PATCH': 1, 'HANGER LOOP': 1, 'PIPING ON FACING': 1})
        trim_list_title_list.append(
            {'FRONT LINING': 2, 'BACK LINING': 2, 'SIDE BODY LINING': 2, 'SLEEVE LINING': 2,
             'TAB LINING': 2, 'INSIDE BESOM LINING': 2, 'INSIDE BESOM FACING LINING': 2, 'INSIDE CELL LINING': 2,
             'INSIDE CELL FACING LINING': 2, 'INSIDE PEN LINING': 2, 'INSIDE PEN FACING LINING': 2,
             'FLAP LINING COAT': 2, 'INSIDE FLAP LIN  ON PATCH': 2, 'HANGER LOOP': 2, 'PIPING ON FACING': 2})
        trim_list_title_list.append(
            {'FRONT BUTTON': 1, 'SLEEVE BUTTON': 1, 'LINING TAB BUTTON': 1, 'INSIDE PATCH BUTTON': 1,
             'INSIDE DOUBLE BREAST BUTTON': 1, 'PANT BUTTON': 1, 'VEST BUTTON': 1, 'BREAST FLAP BUTTON': 1,
             'LOWER FLAP BUTTON': 1, 'INSIDE BREAST BESOM BUTTON': 1, 'INSIDE CELL BUTTON': 1,
             'INSIDE FLAP BUTTON': 1, 'STORM TAB BUTTON': 1, 'BACK BELT BUTTON': 1,
             'SHOULDER TAB BUTTON': 1, 'VENT BUTTON': 1})
        trim_list_title_list.append(
            {'FRONT BUTTON': 2, 'SLEEVE BUTTON': 2, 'LINING TAB BUTTON': 2, 'INSIDE PATCH BUTTON': 2,
             'INSIDE DOUBLE BREAST BUTTON': 2, 'PANT BUTTON': 2, 'VEST BUTTON': 2, 'BREAST FLAP BUTTON': 2,
             'LOWER FLAP BUTTON': 2, 'INSIDE BREAST BESOM BUTTON': 2, 'INSIDE CELL BUTTON': 2,
             'INSIDE FLAP BUTTON': 2, 'STORM TAB BUTTON': 2, 'BACK BELT BUTTON': 2,
             'SHOULDER TAB BUTTON': 2, 'VENT BUTTON': 2})
        trim_list_title_list.append({'INSIDE PANT BUTTON': 1})
        trim_list_title_list.append({'INSIDE PANT BUTTON': 2})
        trim_list_title_list.append({'FUSIBLE': 1})
        trim_list_title_list.append({'CHEST PIECE': 1})
        trim_list_title_list.append({'UNDER COLLAR': 3})
        trim_list_title_list.append({'UNDER COLLAR STAND': 3})
        trim_list_title_list.append(
            {'VEST POCKETING': 4, 'VEST POCKETING OUTSIDE': 2, 'COAT POCKETING': 4, 'COAT POCKETING OUTSIDE': 2})
        trim_list_title_list.append({'SHOULDER PAD': 1})
        trim_list_title_list.append({'SLEEVE HEAD': 2})
        trim_list_title_list.append({'SEAM SLIPPAGE': 1})
        trim_list_title_list.append(
            {'ZROH LAPEL': 1, 'ZROH UPPER POCKET': 1, 'ZROH LOWER POCKET': 1,
             'ZROH LOWER POCKET FACING': 1, 'ZROH BAND': 1, 'TUX SATIN': 1, 'TUX FUSE SATIN': 1})
        trim_list_title_list.append(
            {'INSIDE PATCH': 2, 'INSIDE PIPING': 2, 'OUTSIDE PIPING': 2,
             'HANGER LOOP': 2, 'PIPING ON FACING': 2, 'TUX SATIN PANT': 3, 'SPECIAL FEATURE': 1})
        trim_list_title_list.append(
            {'Comments 1': 1, 'Comments 2': 1, 'Comments 3': 1, 'Comments 4': 1})
        trim_list_title_list.append({'VEST BUCKLE': 2})
        trim_list_title_list.append(
            {'ZIPPER': 2, 'ZIP INS BREAST BESOM PKT': 2})
        trim_list_title_list.append({'PANT LINING': 3})
        trim_list_title_list.append({'PANT POCKETING': 2})
        trim_list_title_list.append({'WAISTBAND': 1})
        trim_list_title_list.append({'WAISTBAND': 2})
        trim_list_title_list.append({'WAISTBAND FUSIBLE': 2})
        trim_list_title_list.append({'WAISTBAND PIPING': 2})
        trim_list_title_list.append({'VEST INSIDE FRONT LINING': 2})
        trim_list_title_list.append({'VEST INSIDE BACK LINING': 2})
        trim_list_title_list.append({'VEST IOUTSIDE BACK LINING': 2})
        trim_list_title_list.append({'LINING TAB BUTTON': 1})
        trim_list_title_list.append({'LINING TAB BUTTON': 2})
        # 根据勤哲数据库中的字段，进行对照整理。整理形式例：{'钎子':[{'VEST BUCKLE':2},{'ZIP INS BREAST BESOM PKT':2}]}
        self.arrange_qinzhe_key = {}
        for qinIdx in range(len(self.add_data_title)):
            self.arrange_qinzhe_key[self.add_data_title[qinIdx]] = trim_list_title_list[qinIdx]

    def update_db(self):
        dbCol = self.add_data_title[:]
        dbCol.append('CreateDate')
        # sql服务器名
        serverName = '192.168.0.11'
        # 登陆用户名和密码
        userName = 'sa'
        passWord = 'jiangbin@007'
        # 建立连接并获取cursor
        conn = pymssql.connect(serverName, userName, passWord, "ESApp1")
        cursor = conn.cursor()
        # 查询存在的KEY
        noInsertPoList = []
        for dataKey in self.allDataKeys:
            selectSql = '''select TLID from D_TrimListInfo where concat(订单PO号, '^_^', version) = \'''' + dataKey + '\''
            cursor.execute(selectSql)
            fatchResults = cursor.fetchone()
            if fatchResults:
                noInsertPoList.append(dataKey.split('^_^')[0])
            else:
                # 需要删除的PO列表
                self.delPoList.append(dataKey.split('^_^')[0])
        # 组装插入的值
        allInsertVal = []
        for tabVal in self.table_value:
            allInsertVal += tabVal
        insertValue = []
        for insertTabVal in allInsertVal:
            if insertTabVal[1] not in noInsertPoList:
                insertValue.append(insertTabVal)
        if len(self.delPoList) > 0:
            # 组装删除的值
            del_tuple = tuple(self.delPoList)
            # 删除已经存在的文件
            delSql = '''delete from D_TrimListInfo where 订单PO号 = (%s)'''
            cursor.executemany(delSql, del_tuple)
        if len(insertValue) > 0:
            insertSql = 'INSERT INTO D_TrimListInfo VALUES ('
            for colVal in dbCol:
                if colVal == 'CreateDate':
                    insertSql += '%s'
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
