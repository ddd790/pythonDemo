import os
import pandas as pd
import pdfplumber
import re
import datetime
import shutil
from decimal import Decimal
from aip import AipOcr


class VAS_GUI():
    # 批量获取服务器数据，进行累加操作
    def get_files(self):
        print('数据操作进行中......' + str(datetime.datetime.now()).split('.')[0])
        APP_ID = '25101742'
        API_KEY = 'Z5qy26GRDUdDKlBRHGT21XZt'
        SECRET_KEY = 'p6BCz0xxGXSTbDR3MfWAfViBRbFilaAu'
        client = AipOcr(APP_ID, API_KEY, SECRET_KEY)
        # 数据库名
        self.dbName = 'ESApp1'
        # 追加的dataFrame的title
        self.add_data_title = ['报关单申报日期', '发票号', '手册号',
                               '报关单号', '加工厂', '报关品名', '数量', '单位', '收汇USD']
        # 数字类型的字段
        self.number_item = ['数量', 'version']
        # 根据勤哲的key匹配对应trimList中的key和value
        # self.local_trim_list_file = r'\\192.168.0.6\03-业务三部共享\EXPRESS 工艺\大货 工艺书\勤哲BOM最新PDF文件'
        self.local_trim_list_file = 'd:\\BOC'
        self.trim_list_file_finish = 'd:\\BOC结果'
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
                    self.file_to_dataframe(os.path.join(lroot, lfile), str(
                        lfile).split('.')[0])
                    self.delete_file.append(lfile)
                except:
                    pdf_file = self.get_file_content(
                        os.path.join(lroot, lfile))
                    # 调用通用文字识别（高精度版）
                    options = {}
                    options["detect_direction"] = "true"
                    res_pdf = client.basicAccuratePdf(pdf_file, options)
                    res_pdf_general = client.basicGeneralPdf(pdf_file, options)
                    if res_pdf['direction'] != 0:
                        print(lfile)
                        continue
                    else:
                        # print(res_pdf)
                        self.ocr_to_dataframe(res_pdf, res_pdf_general)

        table_data = pd.DataFrame(self.arrangeVal, columns=self.add_data_title)
        table_data['数量'] = table_data['数量'].astype('int')
        table_data['收汇USD'] = table_data['收汇USD'].astype('float')
        # 导出excel,追加在old的后面
        excelUrl = self.trim_list_file_finish + '\\出口成衣结算结果.xlsx'
        writer = pd.ExcelWriter(excelUrl, engine='xlsxwriter')
        table_data.to_excel(writer, '出口成衣结算', index=False)
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

    def ocr_to_dataframe(self, msg, general):
        ocr_msg = ''
        ocr_msg_g = ''
        for i in msg.get('words_result'):
            ocr_msg = ocr_msg + ('{}\n'.format(i.get('words')))
        for j in general.get('words_result'):
            ocr_msg_g = ocr_msg_g + ('{}\n'.format(j.get('words')))
        # print(ocr_msg)
        # 申报日, 合同协议号（发票号）, 预录入编号的后8位表示报关单号， 生产销售单位是加工厂,备案号是手册号
        send_date = self.get_value_two_word(ocr_msg, '备案号\n', '\n境外收货人')
        boc_no = self.get_value_two_word(
            ocr_msg, '预录入编号：', '海关编号：').replace('\n', '').strip()[-8:]
        two_row_txt = self.get_value_two_word(
            ocr_msg, '备案号\n', '\n境外收货人').split('\n')
        handbook_no = two_row_txt[-1]
        if handbook_no[0] == '2':
            send_date = handbook_no
        else:
            send_date = two_row_txt[-2]
        pro_com = self.get_value_two_word(
            ocr_msg_g, '许可证号', '\n合同协议号').strip().split('\n')[0]
        case_no = self.get_value_two_word(
            ocr_msg, '离境口岸', '包装种类').strip().split('\n')[1]
        if ocr_msg.__contains__('一般贸易'):
            handbook_no = '一般贸易'
            pro_com = ''
        else:
            case_no = ''
        table_txt = ''
        if ocr_msg.__contains__('特殊关系确认'):
            table_txt = self.get_value_two_word(ocr_msg, '征免\n', '特殊关系确认')
        else:
            table_txt = self.get_value_two_word(ocr_msg, '征免\n', None)
        table_detail_list = table_txt.split('美元')
        for val_idx in range(len(table_detail_list)):
            if len(table_detail_list[val_idx]) > 0:
                val_detail_list = table_detail_list[val_idx].split('\n')
                val_detail_list = [x for x in val_detail_list if len(x) > 1]
                if len(val_detail_list) > 0:
                    # print('---------------------val_detail_list----------------')
                    df_values = []
                    df_values.append(
                        send_date[0:4] + '/' + send_date[4:6] + '/' + send_date[6:8])
                    df_values.append(case_no)
                    df_values.append(handbook_no)
                    df_values.append(boc_no)
                    df_values.append(pro_com)
                    # print(val_detail_list)
                    # 报关品名
                    d_name = re.sub(r'[0-9]+', '', val_detail_list[0])
                    df_values.append(d_name.replace(':', ''))
                    # 数量
                    val_num = re.sub('\D', '', val_detail_list[1])
                    df_values.append(val_num)
                    # 单位
                    df_values.append(re.sub(r'[0-9]+', '', val_detail_list[1]))
                    # 收汇USD
                    df_values.append(
                        round(Decimal(val_num) * Decimal(val_detail_list[2]), 2))
                    self.arrangeVal.append(df_values)

    def file_to_dataframe(self, io, lfile):
        pdf = pdfplumber.open(io)
        count = 0  # 页数
        # 申报日, 合同协议号（发票号）, 预录入编号的后8位表示报关单号， 生产销售单位是加工厂,备案号是手册号
        send_date = ''
        case_no = ''
        boc_no = ''
        pro_com = ''
        handbook_no = ''
        for page in pdf.pages:
            count += 1
            if count == 1:
                # page.extract_text()  # 抓取当前页的全部信息
                print('文件名：' + lfile)
                # 第一页的基本内容
                file_txt = str(page.extract_text()).split(
                    '标记唛码及备注')[0]
                two_row_txt = self.get_value_two_word(
                    file_txt, '备案号', '境外收货人').strip().split(' ')
                handbook_no = two_row_txt[-1]
                if handbook_no[0] == '2':
                    send_date = handbook_no
                else:
                    send_date = two_row_txt[-2]
                pro_com = self.get_value_two_word(
                    file_txt, '许可证号', '合同协议号').strip().split(' ')[0]
                case_no = self.get_value_two_word(
                    file_txt, '离境口岸', '包装种类').strip().split(' ')[0].split('\n')[1]
                if file_txt.__contains__('一般贸易'):
                    handbook_no = '一般贸易'
                    pro_com = ''
                else:
                    case_no = ''
                boc_no = self.get_value_two_word(
                    file_txt, '预录入编号：', '海关编号：').strip()[-8:]
            table_txt = str(page.extract_text()).split(
                '项号')[1]
            if table_txt.__contains__('特殊关系确认'):
                table_txt = self.get_value_two_word(
                    table_txt, '征免', '特殊关系确认').strip()
            else:
                table_txt = str(page.extract_text()).split(
                    '征免\n')[1]
            table_detail_list = table_txt.split('\n')
            for val_idx in range(len(table_detail_list)):
                if val_idx % 3 == 0:
                    df_values = []
                    df_values.append(
                        send_date[0:4] + '/' + send_date[4:6] + '/' + send_date[6:8])
                    df_values.append(case_no)
                    df_values.append(handbook_no)
                    df_values.append(boc_no)
                    df_values.append(pro_com)
                    val_detail_list = table_detail_list[val_idx].split(' ')
                    # 报关品名
                    df_values.append(re.sub(r'[0-9]+', '', val_detail_list[1]))
                    # 数量
                    val_num = re.sub('\D', '', val_detail_list[2])
                    df_values.append(val_num)
                    # 单位
                    df_values.append(re.sub(r'[0-9]+', '', val_detail_list[2]))
                    # 收汇USD
                    df_values.append(
                        round(Decimal(val_num) * Decimal(val_detail_list[3]), 2))
                    self.arrangeVal.append(df_values)

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


def gui_start():
    VAS = VAS_GUI()
    VAS.get_files()


if __name__ == '__main__':
    gui_start()
