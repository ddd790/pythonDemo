import os
import pandas as pd
import pdfplumber
import re
import datetime
import shutil
from decimal import Decimal
from aip import AipOcr


class VAS_GUI():
    # 报关单数据提取
    def get_files(self):
        print('数据操作进行中......' + str(datetime.datetime.now()).split('.')[0])
        APP_ID = '25101742'
        API_KEY = 'Z5qy26GRDUdDKlBRHGT21XZt'
        SECRET_KEY = 'p6BCz0xxGXSTbDR3MfWAfViBRbFilaAu'
        client = AipOcr(APP_ID, API_KEY, SECRET_KEY)
        # 追加的dataFrame的title
        self.add_data_title = ['境内收货人', '进口日期', '备案号', '境外发货人', '消费使用单位', '监管方式', '商品名称', '单价', '总价', '规格型号', '数量', '单位', '币制']
        self.local_trim_list_file = 'd:\\BGD'
        self.trim_list_file_finish = 'd:\\BGD结果'
        # 删除目录内文件
        if os.path.exists(self.trim_list_file_finish):
            shutil.rmtree(self.trim_list_file_finish)
        os.mkdir(self.trim_list_file_finish)

        # 最终dataframe
        self.table_data = pd.DataFrame(data=None, columns=self.add_data_title)
        # 删除列表
        self.delete_file = []
        self.arrangeVal = []
        # 循环文件，处理合并
        for lroot, ldirs, lfiles in os.walk(self.local_trim_list_file):
            for lfile in lfiles:
                try:
                    self.file_to_dataframe(os.path.join(lroot, lfile), str(lfile).split('.')[0])
                    self.delete_file.append(lfile)
                except:
                    pdf_file = self.get_file_content(os.path.join(lroot, lfile))
                    # 调用通用文字识别（高精度版）
                    options = {}
                    options["detect_direction"] = "true"
                    res_pdf = client.basicAccuratePdf(pdf_file, options)
                    # res_pdf_general = client.basicGeneralPdf(pdf_file, options)
                    if res_pdf['direction'] != 0:
                        print('文件不是正确方向，请旋转文件【' + lfile + '】,保存后再重新操作!')
                        continue
                    else:
                        self.ocr_to_dataframe(res_pdf)
                        self.delete_file.append(lfile)
                        # print(res_pdf)

        table_data = pd.DataFrame(self.arrangeVal, columns=self.add_data_title)
        table_data['数量'] = table_data['数量'].astype('float')
        table_data['单价'] = table_data['单价'].astype('float')
        table_data['总价'] = table_data['总价'].astype('float')
        # 导出excel,追加在old的后面
        excelUrl = self.trim_list_file_finish + '\\进口报关单.xlsx'
        writer = pd.ExcelWriter(excelUrl, engine='xlsxwriter')
        table_data.to_excel(writer, '进口报关单明细', index=False)
        writer.save()

        # 删除已经完成的文件
        print('-------------------删除文件列表：----------------------------')
        for del_val in self.delete_file:
            print(del_val)
            os.remove(self.local_trim_list_file + '\\' + del_val)
        # 更新数据库
        print('------------------------------------------------------------')
        print('已经完成操作！' + str(datetime.datetime.now()).split('.')[0])
        input('按回车退出 ')

    def ocr_to_dataframe(self, msg):
        ocr_msg = ''
        for i in msg.get('words_result'):
            ocr_msg = ocr_msg + ('{}\n'.format(i.get('words')))
        # print(ocr_msg)
        # '境内收货人', '进口日期', '备案号', '境外发货人', '消费使用单位', '监管方式', '商品名称', '单价', '总价', '规格型号', '数量', '单位', '币制'
        # 第二行内容
        tmp_jw = '\n境外发货人'
        if ocr_msg.__contains__('境外发费人(NO)'):
            tmp_jw = '\n境外发费人(NO)'
        elif ocr_msg.__contains__('境外发费人'):
            tmp_jw = '\n境外发费人'
        two_row_txt = self.get_value_two_word(ocr_msg, '备案号\n', tmp_jw).split('\n')
        consignee = two_row_txt[0]
        importation_data = two_row_txt[2]
        filing_number = two_row_txt[-1]
        # 境外发货人
        shipper = self.get_value_two_word(ocr_msg, '货物存放地点\n', '\n消费使用单位').strip().split('\n')[0]
        # 第三行内容
        three_row_txt = self.get_value_two_word(ocr_msg, '许可证号\n', '\n合同协议号').strip().split('\n')
        consumption_unit = three_row_txt[1]
        regulatory_methods = three_row_txt[2]
        table_txt = ''
        if ocr_msg.__contains__('特殊关系确认'):
            table_txt = self.get_value_two_word(ocr_msg, '征免\n', '\n特殊关系确认')
        else:
            table_txt = self.get_value_two_word(ocr_msg, '征免\n', None)
        # 按照币制进行分割
        currency = '美元'
        if table_txt.__contains__('欧元') :
            currency = '欧元'
        elif table_txt.__contains__('人民币') :
            currency = '人民币'
        table_detail_list = table_txt.split(currency)
        for val_idx in range(len(table_detail_list)):
            if len(table_detail_list[val_idx]) > 0:
                val_detail_list = table_detail_list[val_idx].split('\n')
                val_detail_list = [x for x in val_detail_list if len(x) > 1]
                if len(val_detail_list) > 0:
                    # print('---------------------val_detail_list----------------')
                    df_values = []
                    df_values.append(consignee)
                    df_values.append(importation_data)
                    df_values.append(filing_number)
                    df_values.append(shipper)
                    df_values.append(consumption_unit)
                    df_values.append(regulatory_methods)
                    # 商品名称
                    df_values.append(re.sub(r'[0-9]+', '', val_detail_list[0]))
                    # 单价
                    df_values.append(val_detail_list[2])
                    # 总价
                    if not self.is_number(val_detail_list[8]):
                        val_detail_list.insert(5, '')
                    df_values.append(val_detail_list[8])
                    # 规格型号
                    all_type = val_detail_list[7]
                    if not val_detail_list[-2].__contains__('('):
                        all_type += val_detail_list[-2]
                    df_values.append(all_type)
                    # 数量
                    df_values.append(re.sub('[^0-9.]', '', val_detail_list[-1]))
                    # 单位
                    df_values.append(re.sub(r'[0-9.]+', '', val_detail_list[-1]))
                    # 币制
                    df_values.append(currency)
                    self.arrangeVal.append(df_values)

    def file_to_dataframe(self, io, lfile):
        pdf = pdfplumber.open(io)
        count = 0  # 页数
        # '境内收货人', '进口日期', '备案号', '境外发货人', '消费使用单位', '监管方式', '商品名称', '单价', '总价', '规格型号', '数量', '单位', '币制'
        consignee = ''
        importation_data = ''
        filing_number = ''
        shipper = ''
        consumption_unit = ''
        regulatory_methods = ''
        for page in pdf.pages:
            count += 1
            if count == 1:
                # 抓取当前页的全部信息
                # print(page.extract_text()) 
                # 第一页的基本内容
                file_txt = str(page.extract_text()).split('海关批注及签章')[0]
                # 第二行内容
                two_row_txt = self.get_value_two_word(file_txt, '备案号', '境外发货人').strip().split(' ')
                consignee = two_row_txt[0]
                importation_data = two_row_txt[2]
                filing_number = two_row_txt[-1]
                # 境外发货人
                tmp_shipper = self.get_value_two_word(file_txt, '货物存放地点', '消费使用单位').strip().split(' ')
                for item in tmp_shipper:
                    if not self.is_chinese(item[0]):
                        shipper += item  + ' '
                    else:
                        break
                three_row_txt = self.get_value_two_word(file_txt, '启运港', '合同协议号').strip().split(' ')
                consumption_unit = three_row_txt[0]
                if three_row_txt[0].__contains__('\n'):
                    consumption_unit = three_row_txt[0].strip().split('\n')[1]
                regulatory_methods = three_row_txt[1]
            table_txt = str(file_txt).split('项号')[1]
            if table_txt.__contains__('特殊关系确认'):
                table_txt = self.get_value_two_word(table_txt, '征免', '特殊关系确认').strip()
            else:
                table_txt = str(page.extract_text()).split('征免\n')[1]
            table_detail_list = table_txt.split('\n')
            # 规格型号
            all_type = ''
            # 商品名称
            pro_name = ''
            # 单价
            pro_price = ''
            # 总价
            sum_price = ''
            for val_idx in range(len(table_detail_list)):
                val_detail_list = table_detail_list[val_idx].split(' ')
                if val_idx % 3 == 0:
                    df_values = []
                    pro_name = ''
                    # 商品名称
                    pro_name = re.sub(r'[0-9]+', '', val_detail_list[1])
                    # 单价
                    pro_price = val_detail_list[3]
                    if pro_name == '':
                        pro_name = val_detail_list[2]
                        pro_price = val_detail_list[4]
                if val_idx % 3 == 1:
                    # 型号
                    all_type = val_detail_list[0]
                    # 总价
                    sum_price = val_detail_list[-5]
                    if not self.is_number(sum_price):
                        sum_price = val_detail_list[-4]
                if val_idx % 3 == 2:
                    df_values.append(consignee)
                    df_values.append(importation_data)
                    df_values.append(filing_number)
                    df_values.append(shipper)
                    df_values.append(consumption_unit)
                    df_values.append(regulatory_methods)
                    # 商品名称
                    df_values.append(pro_name)
                    # 单价
                    df_values.append(pro_price)
                    # 总价
                    df_values.append(sum_price)
                    # 如果第一个是中文，就把数量前面的都追加到型号中
                    if self.is_chinese(val_detail_list[0][0]):
                        arr2 = val_detail_list.copy()
                        all_type += ''.join(arr2[:-2])
                    df_values.append(all_type)
                    # 数量
                    df_values.append(re.sub('[^0-9.]', '', val_detail_list[-2]))
                    # 单位
                    df_values.append(re.sub(r'[0-9.]+', '', val_detail_list[-2]))
                    # 币制
                    df_values.append(val_detail_list[-1])
                    self.arrangeVal.append(df_values)
        pdf.close()
        print('导出成功文件：' + lfile)

    # 获取一个字符串中两个字母中间的值(one为None时从第一位取, two为None时取到最后)
    def get_value_two_word(self, txt_str, one, two):
        if one == None:
            return txt_str[:txt_str.find(two)]
        if two == None:
            return txt_str[txt_str.find(one) + len(one):]
        return txt_str[txt_str.find(one) + len(one):txt_str.find(two)]

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

    def get_file_content(self, filePath):
        with open(filePath, "rb") as fp:
            return fp.read()

    def is_chinese(self, char):
        pattern = re.compile(r'[^\u4e00-\u9fa5]')
        if pattern.search(char):
            return False
        return True


def gui_start():
    VAS = VAS_GUI()
    VAS.get_files()


if __name__ == '__main__':
    gui_start()
