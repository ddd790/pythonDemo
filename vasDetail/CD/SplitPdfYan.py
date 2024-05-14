from ast import Try
import os
import datetime
import shutil
import fitz
from aip import AipOcr
import time


class VAS_GUI():
    # 批量获取服务器数据，进行累加操作
    def get_files(self):
        print('数据操作进行中......' + str(datetime.datetime.now()).split('.')[0])
        APP_ID = '25101742'
        API_KEY = 'Z5qy26GRDUdDKlBRHGT21XZt'
        SECRET_KEY = 'p6BCz0xxGXSTbDR3MfWAfViBRbFilaAu'
        client = AipOcr(APP_ID, API_KEY, SECRET_KEY)
        # 相关文件夹
        self.local_trim_list_file = 'd:\\SplitPdfYan'
        self.trim_list_file_temp = 'd:\\Pdf分割中间文件'
        self.trim_list_file_finish = 'd:\\Pdf分割结果'
        # 删除目录内文件
        if os.path.exists(self.trim_list_file_finish):
            shutil.rmtree(self.trim_list_file_finish)
        os.mkdir(self.trim_list_file_finish)
        # if os.path.exists(self.trim_list_file_temp):
        #     shutil.rmtree(self.trim_list_file_temp)
        os.mkdir(self.trim_list_file_temp)

        # 循环文件，拆分
        try:
            for lroot, ldirs, lfiles in os.walk(self.local_trim_list_file):
                for lfile in lfiles:
                    self.pyMuPDF_fitz(os.path.join(lroot, lfile), lfile)
        finally:
            for lroot, ldirs, lfiles in os.walk(self.trim_list_file_temp):
                for lfile in lfiles:
                    pdf_file = self.get_file_content(os.path.join(lroot, lfile))
                    # 调用通用文字识别
                    options = {}
                    options["detect_direction"] = "true"
                    # （高精度版）
                    res_pdf = client.basicAccuratePdf(pdf_file, options)
                    # 普通版
                    # res_pdf = client.basicGeneralPdf(pdf_file, options)
                    # 百度ocr
                    back_file_name = self.ocr_to_dataframe(res_pdf)
                    if back_file_name.__contains__('\n'):
                        back_file_name = back_file_name.replace('\n', '')
                    if back_file_name.__contains__('户名'):
                        back_file_name = back_file_name.split('户名')[1]
                    if back_file_name:
                        count = self.check_file_name(back_file_name)
                        if count == 0:
                            os.rename(os.path.join(lroot, lfile), self.trim_list_file_finish + '\\' + back_file_name + '.pdf')
                        else:
                            os.rename(os.path.join(lroot, lfile), self.trim_list_file_finish + '\\' + back_file_name + '-' + str(count) + '.pdf')
                    time.sleep(2)

            shutil.rmtree(self.trim_list_file_temp)
            shutil.rmtree(self.local_trim_list_file)
            os.mkdir(self.local_trim_list_file)
            print('已经完成操作！' + str(datetime.datetime.now()).split('.')[0])
            input('按回车退出 ')

    def check_file_name(self, file_name):
        file_name_list = []
        for lroot, ldirs, lfiles in os.walk(self.trim_list_file_finish):
            for lfile in lfiles:
                if lfile.__contains__(file_name):
                    file_name_list.append(file_name)
        return len(file_name_list)

    def pyMuPDF_fitz(self, pdfPath, lfile):
        # print(pdfPath)
        pdfDoc = fitz.open(pdfPath)
        for pg in range(pdfDoc.page_count):
            old_rect = pdfDoc[pg].rect
            base_width = old_rect.width
            base_height = old_rect.height
            # 一个pdf截取成三段
            for splite_pg in range(3):
                temp_y = 65
                temp_height = 305
                if splite_pg == 1:
                    temp_y = 306
                    temp_height = 535
                elif splite_pg == 2:
                    temp_y = 540
                    temp_height = 930
                file_pg = lfile.split('.')[0].strip() + '_' + str(pg) + '_' + str(splite_pg)
                old_clip = fitz.Rect(0, temp_y, base_width, temp_height)
                DOC3 = fitz.open()
                placerect = fitz.Rect(0, 0, base_width, 500)
                page = DOC3.new_page(width=base_width, height=500)
                page.show_pdf_page(placerect, pdfDoc, pg, clip=old_clip)
                DOC3.save(self.trim_list_file_temp + '/' + '我的新文档_%s.pdf' % file_pg, garbage=4, deflate=True)

            # 一个pdf截取成三段(旧版)
            # for splite_pg in range(3):
            #     file_pg = lfile.split('.')[0].strip() + '_' + str(pg) + '_' + str(splite_pg)
            #     old_clip = fitz.Rect(0, base_height * splite_pg / 3, base_width, base_height * (splite_pg + 1) / 3)
            #     DOC3 = fitz.open()
            #     placerect = fitz.Rect(0, 0, base_width, base_height/3)
            #     page = DOC3.new_page(width=base_width, height=base_height/3)
            #     page.show_pdf_page(placerect, pdfDoc, pg, clip=old_clip)
            #     DOC3.save(self.trim_list_file_temp + '/' + '我的新文档_%s.pdf' % file_pg, garbage=4, deflate=True)

    def ocr_to_dataframe(self, msg):
        ocr_msg = ''
        for i in msg.get('words_result'):
            ocr_msg = ocr_msg + ('{}\n'.format(i.get('words')))
        # print(ocr_msg)
        # 收款方户名作为文件名
        accept_name = self.get_value_two_word(ocr_msg, '收款方\n户名', '\n开户行').replace('/', '')
        return accept_name.strip()

    def get_value_two_word(self, txt_str, one, two):
        # 获取一个字符串中两个字母中间的值(one为None时从第一位取, two为None时取到最后)
        if one == None:
            return txt_str[:txt_str.find(two)]
        if two == None:
            return txt_str[txt_str.find(one) + len(one):]
        return txt_str[txt_str.find(one) + len(one):txt_str.find(two)]

    def get_file_content(self, filePath):
        with open(filePath, "rb") as fp:
            return fp.read()


def gui_start():
    VAS = VAS_GUI()
    VAS.get_files()


if __name__ == '__main__':
    gui_start()
