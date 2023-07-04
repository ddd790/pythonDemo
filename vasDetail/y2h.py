import os
import win32com.client
import gc
import pandas as pd
import datetime
import shutil
from docxtpl import DocxTemplate
import docx
from docxcompose.composer import Composer
from glob import glob
import time
from PyPDF2 import PdfFileReader, PdfFileWriter


class VAS_GUI():
    # 批量获取服务器数据，进行累加操作
    def get_files(self):
        print('数据操作进行中......' + str(datetime.datetime.now()).split('.')[0])
        self.local_trim_list_file = 'd:\\审核资料-成品模板'
        # self.local_trim_list_file = 'd:\\审核资料-面料TC模板'
        # self.local_trim_list_file = 'd:\\y2h'
        self.trim_list_file_finish = 'd:\\y2h成品合集'
        # self.trim_list_file_finish = 'd:\\y2h面料TC合集'
        # self.trim_list_file_finish = 'd:\\y2h合集'
        # 删除目录内文件
        if os.path.exists(self.trim_list_file_finish):
            shutil.rmtree(self.trim_list_file_finish)
        os.mkdir(self.trim_list_file_finish)

        # 循环文件,拆分成各自模板
        for lroot, ldirs, lfiles in os.walk(self.local_trim_list_file):
            for lfile in lfiles:
                if str(lfile).__contains__('.xls') or str(lfile).__contains__('.xlsx'):
                    self.file_to_dataframe(os.path.join(lroot, lfile))
        # word转PDF
        # 将目标文件夹所有文件归类，转换时只打开一个进程
        words = []
        for fn in os.listdir(self.trim_list_file_finish):
            if fn.endswith(('.doc', 'docx')):
                words.append(fn)

        # 新建 pdf 文件夹，所有生成的 PDF 文件都放在里面
        folder = self.trim_list_file_finish + '\\pdf\\'
        if not os.path.exists(folder):
            os.makedirs(folder)
        self.word2Pdf(self.trim_list_file_finish, words)

        # 合并所有pdf文件为一个pdf
        file_writer = PdfFileWriter()
        for root, dirs, files in os.walk(self.trim_list_file_finish + '\\pdf'):
            for file in files:
                if str(file).__contains__('.pdf'):
                    # 循环读取需要合并pdf文件
                    file_reader = PdfFileReader(os.path.join(root, file))
                    # 遍历每个pdf的每一页
                    for page in range(file_reader.getNumPages()):
                        # 写入实例化对象中
                        file_writer.addPage(file_reader.getPage(page))
        with open(self.trim_list_file_finish + '\\0all.pdf', 'wb') as out:
            file_writer.write(out)

        print('已经完成操作！' + str(datetime.datetime.now()).split('.')[0])
        input('按回车退出 ')

    def file_to_dataframe(self, io):
        # ---------------------- 劳动合同 ------------------
        # excel_header = ['工号', '姓名', '部门', '职务', '性别', '入职时间', '出生日期', '身份证号', '地址']
        # ---------------------- 审核资料-面料TC ------------------
        # excel_header = ['编号', 'TCNo', '供应商', 'SCNo', '发货日期', '收货工厂', '品名', '成份',
        #                 '采购合同号', '品号', '颜色', '米数m', '毛重', '净重', '认证重量', '卷数', '备注', '入库编号', '入库日期', '性质']
        # ---------------------- 审核资料-成品 ------------------
        excel_header = ['编号', '收货工厂', '成份', '采购合同号', '品号', '颜色', '备注', '裁剪领料日期', '裁剪投入量', '裁剪产出重量', '损耗', '物料库存', '缝纫领料日期', '缝纫投入量', '缝纫产出重量', '缝纫损耗', '整熨日期', '整熨投入量', '整熨产出重量', '整熨损耗', '检查日期', '检查投入量',
                        '检查产出重量', '检查损耗', '包装日期', '包装投入量', '包装产出重量', '包装损耗', '加工合同号', '成衣品名', '订单号', '款号', '成衣数量', '成衣净重', '辅料重量', '成品面料重量', '箱数', '库存', '成衣TCNo', '出运日期', '成品编号', '面料出库编号', '裁剪投入米数', '验货日期', 'TCNo']
        excelData = pd.read_excel(io, header=0, keep_default_na=False)
        df = pd.DataFrame(excelData.values, columns=excel_header)
        # 循环df内容，进行赋值
        # ---------------------- 劳动合同 ------------------
        # read_doc_file_path = self.local_trim_list_file + '\\劳动合同.docx'
        # for i in range(len(df)):
        #     # 退休返聘人员剔除
        #     # （女的 1973.3.23日之前的， 男的 1963.3.23日之前的，已经退休了）
        #     str_bir_time = str(df.iloc[i]['出生日期']).replace('.', '')
        #     now_age = self.get_age(str_bir_time)
        #     if (df.iloc[i]['性别'] == '女' and now_age < 50) or (df.iloc[i]['性别'] == '男' and now_age < 60):
        #         context = {
        #             "工号": df.iloc[i]['工号'],
        #             "姓名": df.iloc[i]['姓名'],
        #             "部门": df.iloc[i]['部门'],
        #             "职务": df.iloc[i]['职务'],
        #             "性别": df.iloc[i]['性别'],
        #             "入职时间": df.iloc[i]['入职时间'],
        #             "出生日期": df.iloc[i]['出生日期'],
        #             "身份证号": df.iloc[i]['身份证号'],
        #             "地址": df.iloc[i]['地址']
        #         }
        #         tpl = DocxTemplate(read_doc_file_path)
        #         tpl.render(context)
        #         tpl.save(self.trim_list_file_finish+r"\{}的劳动合同.docx".format(str(df.iloc[i]['工号']) + df.iloc[i]['姓名']))
        # ---------------------- 审核资料-面料TC ------------------
        # read_doc_file_path = self.local_trim_list_file + '\\审核资料-面料TC.docx'
        # for i in range(len(df)):
        #     context = {
        #         "编号": df.iloc[i]['编号'],
        #         "TCNo": df.iloc[i]['TCNo'],
        #         "供应商": df.iloc[i]['供应商'],
        #         "SCNo": df.iloc[i]['SCNo'],
        #         "发货日期": str(df.iloc[i]['发货日期'])[:10],
        #         "收货工厂": df.iloc[i]['收货工厂'],
        #         "品名": df.iloc[i]['品名'],
        #         "成份": df.iloc[i]['成份'],
        #         "采购合同号": df.iloc[i]['采购合同号'],
        #         "品号": df.iloc[i]['品号'],
        #         "颜色": df.iloc[i]['颜色'],
        #         "米数m": df.iloc[i]['米数m'],
        #         "毛重": df.iloc[i]['毛重'],
        #         "净重": df.iloc[i]['净重'],
        #         "认证重量": df.iloc[i]['认证重量'],
        #         "卷数": df.iloc[i]['卷数'],
        #         "备注": df.iloc[i]['备注'],
        #         "入库编号": df.iloc[i]['入库编号'],
        #         "入库日期": str(df.iloc[i]['入库日期'])[:10],
        #         "性质": df.iloc[i]['性质']
        #     }
        #     tpl = DocxTemplate(read_doc_file_path)
        #     tpl.render(context)
        #     tpl.save(self.trim_list_file_finish+r"\{}的审核资料-面料TC.docx".format(str(df.iloc[i]['编号']) + str(df.iloc[i]['收货工厂'])))
        # ---------------------- 审核资料-成品 ------------------
        read_doc_file_path = self.local_trim_list_file + '\\审核资料-成品.docx'
        for i in range(len(df)):
            context = {
                "编号": df.iloc[i]['编号'],
                "收货工厂": df.iloc[i]['收货工厂'],
                "成份": df.iloc[i]['成份'],
                "采购合同号": df.iloc[i]['采购合同号'],
                "品号": df.iloc[i]['品号'],
                "颜色": df.iloc[i]['颜色'],
                "备注": df.iloc[i]['备注'],
                "裁剪领料日期": str(df.iloc[i]['裁剪领料日期'])[:10],
                "裁剪投入量": df.iloc[i]['裁剪投入量'],
                "裁剪产出重量": df.iloc[i]['裁剪产出重量'],
                "损耗": str(round(float(df.iloc[i]['损耗']), 2) * 100) + '%',
                "物料库存": df.iloc[i]['物料库存'],
                "缝纫领料日期": str(df.iloc[i]['缝纫领料日期'])[:10],
                "缝纫投入量": df.iloc[i]['缝纫投入量'],
                "缝纫产出重量": df.iloc[i]['缝纫产出重量'],
                "缝纫损耗": df.iloc[i]['缝纫损耗'],
                "整熨日期": str(df.iloc[i]['整熨日期'])[:10],
                "整熨投入量": df.iloc[i]['整熨投入量'],
                "整熨产出重量": df.iloc[i]['整熨产出重量'],
                "整熨损耗": df.iloc[i]['整熨损耗'],
                "检查日期": str(df.iloc[i]['检查日期'])[:10],
                "检查投入量": df.iloc[i]['检查投入量'],
                "检查产出重量": df.iloc[i]['检查产出重量'],
                "检查损耗": df.iloc[i]['检查损耗'],
                "包装日期": str(df.iloc[i]['包装日期'])[:10],
                "包装投入量": df.iloc[i]['包装投入量'],
                "包装产出重量": df.iloc[i]['包装产出重量'],
                "包装损耗": df.iloc[i]['包装损耗'],
                "加工合同号": df.iloc[i]['加工合同号'],
                "成衣品名": df.iloc[i]['成衣品名'],
                "订单号": df.iloc[i]['订单号'],
                "款号": df.iloc[i]['款号'],
                "成衣数量": df.iloc[i]['成衣数量'],
                "成衣净重": df.iloc[i]['成衣净重'],
                "辅料重量": df.iloc[i]['辅料重量'],
                "成品面料重量": df.iloc[i]['成品面料重量'],
                "箱数": df.iloc[i]['箱数'],
                "库存": df.iloc[i]['库存'],
                "成衣TCNo": df.iloc[i]['成衣TCNo'],
                "出运日期": str(df.iloc[i]['出运日期'])[:10],
                "成品编号": df.iloc[i]['成品编号'],
                "面料出库编号": df.iloc[i]['面料出库编号'],
                "裁剪投入米数": round(float(df.iloc[i]['裁剪投入米数']), 2),
                "验货日期": str(df.iloc[i]['验货日期'])[:10],
                "TCNo": df.iloc[i]['TCNo']
            }
            tpl = DocxTemplate(read_doc_file_path)
            tpl.render(context)
            tpl.save(self.trim_list_file_finish+r"\{}的审核资料-成品.docx".format(str(df.iloc[i]['编号']) + str(df.iloc[i]['收货工厂'])))

    # 合并word文档
    def merge_file(self, files_list):
        number_of_sections = len(files_list)
        master = docx.Document()
        composer = Composer(master)

        for i in range(0, number_of_sections):
            doc_temp = docx.Document((files_list[i]))
            composer.append(doc_temp)
        composer.save(os.path.join(self.trim_list_file_finish, 'all.docx'))

    # Word2pdf
    def word2Pdf(self, filePath, words):
        # 如果没有文件则提示后直接退出
        if (len(words) < 1):
            print("\n【无 Word 文件】\n")
            return
        # 开始转换
        try:
            print("\n【开始 Word -> PDF 转换】")
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = 0
            word.DisplayAlerts = False
            doc = None
            for i in range(len(words)):
                try:
                    fileName = words[i]  # 文件名称
                    fromFile = os.path.join(filePath, fileName)  # 文件地址
                    toFileName = self.changeSufix2Pdf(fileName)  # 生成的文件名称
                    toFile = self.toFileJoin(filePath, toFileName)  # 生成的文件地址
                    # 某文件出错不影响其他文件打印
                    # time.sleep(1)
                    doc = word.Documents.Open(fromFile)
                    doc.SaveAs(toFile, 17)  # 生成的所有 PDF 都会在 PDF 文件夹中
                    print("转换到："+toFileName+"完成")
                except Exception as e:
                    print(e)
            # 关闭 Word 进程
            doc.Close()
            doc = None
            word.Quit()
            word = None
        except Exception as e:
            print(e)
        finally:
            gc.collect()

    # 修改后缀名
    def changeSufix2Pdf(self, file):
        return file[:file.rfind('.')]+".pdf"

    # 转换地址
    def toFileJoin(self, filePath, file):
        return os.path.join(filePath, 'pdf', file[:file.rfind('.')]+".pdf")

    def get_age(self, birthday):
        # 本函数根据输入的8位出生年月日数据返回截至当天的年龄
        today = str(datetime.datetime.now().strftime('%Y-%m-%d')).split("-")
        # 取出系统当天的年月日数据为列表[年,月,日]
        n_monthandday = today[1] + today[2]
        # 将月日连接在一起
        n_year = today[0]
        # 单独列出当年年份
        r_monthandday = birthday[4:]
        # 取出输入日期的月与日
        r_year = birthday[:4]
        # 取出输入日期的年份

        if (int(n_monthandday) >= int(r_monthandday)):
            # 如果月日比系统月日数据要小，刚直接用年份相减就是
            r_age = int(n_year)-int(r_year)
        else:
            r_age = int(n_year)-int(r_year)-1
        return r_age


def gui_start():
    VAS = VAS_GUI()
    VAS.get_files()


if __name__ == '__main__':
    gui_start()
