import os
import shutil
import re
import pandas as pd
import pyodbc
import datetime
from tkinter import *


class VAS_GUI():
    # 批量获取服务器数据，进行累加操作
    def commit_batch(self):
        print('数据操作进行中......')
        self.select_db_po_info()
        self.get_files()

    def get_files(self):
        networked_directory = r'\\192.168.0.3\01-业务一部资料\=14785212\PEERLESS\F21'
        self.local_vas_detail_file = 'd:\excelVasDetailFile'
        # 删除目录内文件
        if os.path.exists(self.local_vas_detail_file):
            shutil.rmtree(self.local_vas_detail_file)
        os.mkdir(self.local_vas_detail_file)
        # copy服务器的Vas_details文件到本地
        for root, dirs, files in os.walk(networked_directory):
            for file in files:
                if str(file).__contains__('Vas_details') and (str(file).__contains__('.xls') or str(file).__contains__('.xlsx')):
                    shutil.copy(os.path.join(root, file),
                                self.local_vas_detail_file)
        # 保留相同文件中最大的记录
        self.compare_xls_file()

        # 循环本地临时文件，处理合并，并存入数据库
        self.table_value = []
        for lroot, ldirs, lfiles in os.walk(self.local_vas_detail_file):
            for lfile in lfiles:
                self.read_excel(os.path.join(lroot, lfile))

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
                            if int(tempDelFileNameList[5]) > int(nameList[5]):
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
        dataItem = ['Purchasing Document', 'Item', 'Material',
                    'Grid Value', 'Quantity', 'HG', 'PU', 'SZ']
        smallSize = ['XXS', 'XS', 'S', 'M', 'L', 'XL']
        # 对excel读取的数据进行整理，整理成符合要求的格式（按照dataItem中的列进行排列）
        data = self.arrange_excel_data(io, dataItem)
        # 筛选符合条件的数据
        filterData = []
        itemValMap = {}
        for val in range(0, len(data)):
            sizeVal = data.iloc[val, 3]
            filterVal = {}
            tempSizeVal = str(sizeVal)[0:2]
            pk = str(data.iloc[val, 0]) + '-' + str(data.iloc[val, 2])
            styleType = self.type_by_po_info(pk)

            # HG列的组合(按照Item来进行分)
            itemVal = data.iloc[val, 1]
            hgVal = data.iloc[val, 5]
            if itemVal not in itemValMap.keys():
                hgValList = []
            else:
                hgValList = itemValMap[itemVal]
            if str(hgVal).strip() != '' and hgVal not in hgValList:
                hgValList.append(hgVal)
            itemValMap[itemVal] = hgValList
            if styleType == 0:
                # 上装尺码是>=48 or 尺码！=S,M,L,XL
                if (self.is_number(tempSizeVal) and int(tempSizeVal) >= 48) or (self.is_number(tempSizeVal) == False and str(sizeVal).strip() not in smallSize):
                    for itemIndex in range(0, len(dataItem)):
                        tempVal = int(data.iloc[val, itemIndex]) if type(
                            data.iloc[val, itemIndex]) is int else str(data.iloc[val, itemIndex]).strip()
                        filterVal[dataItem[itemIndex]] = tempVal
                    filterVal['HG'] = '\\'.join(itemValMap[itemVal])
                    filterData.append(filterVal)
            elif styleType == 1:
                # 裤子尺码是>=44
                if self.is_number(tempSizeVal) and int(tempSizeVal) >= 44:
                    for itemIndex in range(0, len(dataItem)):
                        tempVal = int(data.iloc[val, itemIndex]) if type(
                            data.iloc[val, itemIndex]) is int else str(data.iloc[val, itemIndex]).strip()
                        filterVal[dataItem[itemIndex]] = tempVal
                    filterVal['HG'] = '\\'.join(itemValMap[itemVal])
                    filterData.append(filterVal)

        if len(filterData) <= 0:
            return
        # 循环筛选后的结果,item相同的，进行itemSum的计算
        sumTitle = ['PO', 'Item', 'Material', 'ItemSum',
                    'Sumxxl', 'Sum48', 'Sum44', 'HG', 'PU', 'SZ', 'Grid Value']
        po = filterData[0][dataItem[0]]
        item = filterData[0][dataItem[1]]
        material = filterData[0][dataItem[2]]
        gridVal = filterData[0][dataItem[3]]
        hg = filterData[0][dataItem[5]]
        pu = filterData[0][dataItem[6]]
        sz = filterData[0][dataItem[7]]
        itemSum = int(filterData[0][dataItem[4]])
        sumInfo = []
        sumInfo.append({
            sumTitle[0]: po,
            sumTitle[1]: item,
            sumTitle[2]: material,
            sumTitle[3]: itemSum,
            sumTitle[4]: 0,
            sumTitle[5]: 0,
            sumTitle[6]: 0,
            sumTitle[7]: hg,
            sumTitle[8]: pu,
            sumTitle[9]: sz,
            sumTitle[10]: gridVal
        })

        # po和item不重复key
        poItemKey = []
        poItemKey.append(po + '-' + item)

        # po和item相同的列进行合计
        for idx, value in enumerate(filterData):
            if idx > 0:
                # po和item相等，将ItemSum进行相加
                if (value[dataItem[0]] + '-' + value[dataItem[1]]) in poItemKey:
                    for sumI, sumV in enumerate(sumInfo):
                        if sumV[sumTitle[0]] == value[dataItem[0]] and sumV[sumTitle[1]] == value[dataItem[1]]:
                            sumV[sumTitle[3]] += int(value[dataItem[4]])
                            break
                else:
                    po = value[dataItem[0]]
                    item = value[dataItem[1]]
                    itemSum = int(value[dataItem[4]])
                    material = value[dataItem[2]]
                    hg = value[dataItem[5]]
                    pu = value[dataItem[6]]
                    sz = value[dataItem[7]]
                    gridVal = value[dataItem[3]]
                    poItemKey.append(po + '-' + item)
                    sumInfo.append({
                        sumTitle[0]: po,
                        sumTitle[1]: item,
                        sumTitle[2]: material,
                        sumTitle[3]: itemSum,
                        sumTitle[4]: 0,
                        sumTitle[5]: 0,
                        sumTitle[6]: 0,
                        sumTitle[7]: hg,
                        sumTitle[8]: pu,
                        sumTitle[9]: sz,
                        sumTitle[10]: gridVal
                    })

        # 相同的po和Material进行大码的统计，并填充到对应的item中
        sumPo = sumInfo[0][sumTitle[0]]
        sumMaterial = sumInfo[0][sumTitle[2]]
        sumxxl = sum48 = sum44 = 0
        sumBigNumInfo = []
        sumBigNumInfo.append({
            sumTitle[0]: sumPo,
            sumTitle[2]: sumMaterial,
            sumTitle[4]: 0,
            sumTitle[5]: 0,
            sumTitle[6]: 0,
        })

        # po和Material不重复key
        poMaterialKey = []
        poMaterialKey.append(sumPo + '-' + sumMaterial)

        # 循环filter后的结果集，计算相同po和Material的合计
        for idx, value in enumerate(sumInfo):
            tempPk = str(value[sumTitle[0]] + '-' + value[sumTitle[2]])
            tempStyleType = self.type_by_po_info(tempPk)
            if tempPk in poMaterialKey:
                tempIndex = poMaterialKey.index(tempPk)
                if self.is_number(str(value[sumTitle[10]])[0:2]) and tempStyleType == 0:
                    sum44 = 0
                    sum48 = sumBigNumInfo[tempIndex][sumTitle[5]
                                                     ] + int(value[sumTitle[3]])
                elif self.is_number(str(value[sumTitle[10]])[0:2]) and tempStyleType == 1:
                    sum48 = 0
                    sum44 = sumBigNumInfo[tempIndex][sumTitle[6]
                                                     ] + int(value[sumTitle[3]])
                else:
                    sum44 = 0
                    sumxxl = sumBigNumInfo[tempIndex][sumTitle[4]
                                                      ] + int(value[sumTitle[3]])
                sumBigNumInfo[tempIndex][sumTitle[4]] = sumxxl
                sumBigNumInfo[tempIndex][sumTitle[5]] = sum48
                sumBigNumInfo[tempIndex][sumTitle[6]] = sum44
            else:
                sumPo = value[sumTitle[0]]
                sumMaterial = value[sumTitle[2]]
                poMaterialKey.append(sumPo + '-' + sumMaterial)
                sum48 = 0
                sum44 = 0
                sumxxl = 0
                if self.is_number(str(value[sumTitle[10]])[0:2]) and tempStyleType == 0:
                    sum48 = int(value[sumTitle[3]])
                elif self.is_number(str(value[sumTitle[10]])[0:2]) and tempStyleType == 1:
                    sum44 = int(value[sumTitle[3]])
                else:
                    sumxxl = int(value[sumTitle[3]])
                sumBigNumInfo.append({
                    sumTitle[0]: sumPo,
                    sumTitle[2]: sumMaterial,
                    sumTitle[4]: sumxxl,
                    sumTitle[5]: sum48,
                    sumTitle[6]: sum44
                })

        # 进行统计信息的赋值
        for si, sv in enumerate(sumInfo):
            for bi, bv in enumerate(sumBigNumInfo):
                if sv[sumTitle[0]] == bv[sumTitle[0]] and sv[sumTitle[2]] == bv[sumTitle[2]]:
                    sv[sumTitle[4]] = bv[sumTitle[4]]
                    sv[sumTitle[5]] = bv[sumTitle[5]]
                    sv[sumTitle[6]] = bv[sumTitle[6]]
            # 放入显示表格数据
            tempTableVal = []
            for i in range(10):
                tempTableVal.append(sv[sumTitle[i]])
            # 显示表格数据赋值
            self.table_value.append(tempTableVal)

    def update_db(self):
        dbCol = ['PO', 'Item', 'Material', 'ItemSum', 'Sumxxl',
                 'Sum48', 'Sum44', 'HG', 'PU', 'SZ', 'CreateDate']
        strCol = ",".join(str(i) for i in dbCol)
        todayTime = str(datetime.datetime.now()).split('.')[0]
        cn = pyodbc.connect(
            'DRIVER={SQL Server};SERVER=192.168.0.6;DATABASE=ESApp1;UID=sa;PWD=MS_guanli09')
        cn.autocommit = True
        cr = cn.cursor()
        # 删除数据
        truncatSql = 'TRUNCATE TABLE D_VasDetail'
        cr.execute(truncatSql)
        # 循环插入数据
        for ti, tv in enumerate(self.table_value):
            # insert数据
            sql = "INSERT INTO D_VasDetail (" + strCol + ")  VALUES ("
            insertsql = ""
            for insertI, insertV in enumerate(tv):
                insertsql += "'" + str(insertV) + "',"
            sql += insertsql + "'" + todayTime + "')"
            cr.execute(sql)
        cr.close()
        cn.close()

    # 查询已有PO相关数据
    def select_db_po_info(self):
        cn = pyodbc.connect(
            'DRIVER={SQL Server};SERVER=192.168.0.6;DATABASE=ESApp1;UID=sa;PWD=MS_guanli09')
        # 查询目前的【一部PO大表订单信息_明细】表的记录
        searchSql = "select 订单PO号, 款式缩写, 面料, 品名 from 一部PO大表订单信息_明细"
        cn.autocommit = True
        cr = cn.cursor()
        cr.execute(searchSql)
        self.vid = cr.fetchall()
        cr.close()
        cn.close()

    # 根据PO号， 款式缩写, 面料，判断是上衣还是裤子,裤子返回1，其他返回0，异常返回2
    def type_by_po_info(self, comVal):
        poVals = {}
        for po, style, ml, pm in self.vid:
            poVals[po + '-' + style + '-' + ml] = pm

        try:
            if str(poVals[comVal]).__contains__('裤子'):
                return 1
            else:
                return 0
        except:
            return 2

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

        return new_df

    def file_to_dataframe(self, io):
        file_name = self.excel_csv_change(io, 1)
        formatCsvData = self.csv_to_dataframe(file_name)
        change_file_name = self.excel_csv_change(file_name, 2)
        # shutil.move(change_file_name, self.local_vas_detail_file)
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
                if tempIdx == 1:
                    tempCsvVal[tempIdx] = str(tempCsvVal[tempIdx]).lstrip('0')
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
