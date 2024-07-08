# -*- coding:utf-8 -*-
import os
import shutil
import pandas as pd
from tkinter import *
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment, Side, Border, Font
from openpyxl.worksheet.dimensions import RowDimension
from dateutil import parser
import pdfplumber
import numpy as np


class VAS_GUI():
    # 批量获取服务器数据，进行累加操作
    def get_files(self):
        print('数据操作进行中......')
        # 追加的dataFrame的title
        # self.add_data_title = ['用料名称', '品号', '颜色', '规格', '预估单耗']
        self.add_data_title = []
        self.add_data_title_1 = ['用料名称']
        self.add_data_title_2 = ['供应商', '有效幅宽/规格', '使用部位', '单耗', '单价', '克重', '成份' '损耗率（%）', '备注']
        # 根据勤哲的key匹配对应trimList中的key和value
        self.local_trim_list_file = 'd:\\IZACTrimlist'
        self.trim_list_file_finish = 'd:\\IZACTrimlist结果'
        self.en_cn_file = r'\\192.168.0.3\01-业务一部资料\软件\trimlistToBom\en_cn.xlsx'
        self.cn_sample_file = r'\\192.168.0.3\01-业务一部资料\A-Serena\2 - 辅料表\样例'
        # 删除目录内文件
        if os.path.exists(self.trim_list_file_finish):
            shutil.rmtree(self.trim_list_file_finish)
        os.mkdir(self.trim_list_file_finish)

        # 读取中英文翻译的配置文件
        # 英文行数（默认最大是301行）
        self.max_en = 301
        self.en_cn = {}
        self.get_en_cn_config(self.en_cn_file)

        # 循环本地临时文件，处理合并
        self.table_value = []
        for lroot, ldirs, lfiles in os.walk(self.local_trim_list_file):
            for lfile in lfiles:
                print(lfile)
                # UNDER COLLAR 的值为 SHELL，需要追加一行
                self.underCollarFlag = False
                self.file_to_dataframe(os.path.join(lroot, lfile), str(lfile).replace('.pdf', '').replace('.PDF', ''))

        # 循环读取文件，修改样式
        for root, dirs, files in os.walk(self.trim_list_file_finish):
            for file in files:
                self.change_file_style(os.path.join(root, file))

        # print(self.table_value)
        print('已经完成导出操作！请到D盘【IZACTrimlist结果】中查看文件吧~~~~')
        input('按回车退出 ')

    # 获取中英文对照配置文件
    def get_en_cn_config(self, io):
        # 打开工作表
        wb = load_workbook(io)
        ws = wb.active
        for i in range(1, self.max_en):
            data_en_value = ws.cell(row=i, column=1).value
            data_cn_value = ws.cell(row=i, column=2).value
            self.en_cn[data_en_value] = data_cn_value

    # 修改文件样式
    def change_file_style(self, io):
        # 打开工作表
        wb = load_workbook(io)
        ws = wb.active
        # 调整第列宽
        ws.column_dimensions['A'].width = 27.38
        ws.column_dimensions['B'].width = 25.25
        ws.column_dimensions['C'].width = 25.25
        ws.column_dimensions['D'].width = 25.25
        ws.column_dimensions['E'].width = 25.25

        # 冻结首行
        ws.freeze_panes = 'A3'

        # 定义表头颜色样式为橙色
        header_fill = PatternFill('solid', fgColor='FFFF00')
        header_fill_qing = PatternFill('solid', fgColor='00FFFF')
        header_fill_lan = PatternFill('solid', fgColor='EE82EE')

        # 定义对齐样式横向居中、纵向居中
        align = Alignment(horizontal='center', vertical='center', wrapText=True)
        # 定义对齐样式纵向居中, 自动换行
        align_center = Alignment(vertical='center', wrapText=True)

        # 定义边样式为细条
        side = Side('thin')
        # 定义边框样式，有底边和右边
        border = Border(top=side, bottom=side, left=side, right=side)

        # 用料名称所在行
        row_use_name = 0
        for i in range(1, self.max_en + 5):
            data_value = ws.cell(row=i, column=1).value
            if data_value == '用料名称':
                row_use_name = i
                break
        
        # 设置品号列的样式
        col_dict = {2: 'B', 3: 'C', 4: 'D', 5: 'E', 6: 'F', 7: 'G', 8: 'H', 9: 'I'}
        for i in range(2, 10):
            data_value = ws.cell(row=2, column=i).value
            if data_value == '品号':
                ws.column_dimensions[col_dict[i]].width = 38
            elif data_value == '供应商':
                ws.column_dimensions[col_dict[i]].width = 18
            elif data_value == '使用部位':
                ws.column_dimensions[col_dict[i]].width = 48

        # 设置基本数据的样式
        base_step = 0
        for row in ws.iter_rows(min_row=1, max_row=row_use_name):
            base_step = base_step + 1
            # ws.row_dimensions[base_step].height = 18
            RowDimension(ws, index=base_step, height=True)
            for cell in row:
                cell.alignment = align_center
                cell.font = Font(size=12)
                if str(cell.value).__contains__('接缝滑移'):
                    # 设置单元格填充颜色
                    cell.fill = header_fill_lan
                    cell.font = Font(bold=True)
                if str(cell.value).__contains__('面料描述'):
                    # 设置单元格填充颜色
                    cell.fill = header_fill_qing
                    # 设置单元格对齐方式
                    cell.font = Font(bold=True)

        # 设置用料名称单元格格式
        ws.row_dimensions[row_use_name].height = 30
        for cell in ws[row_use_name]:
            # 设置单元格填充颜色
            cell.fill = header_fill
            # 设置单元格对齐方式
            cell.alignment = align
            # 设置单元格边框
            cell.border = border
            cell.font = Font(bold=True)
        step = 0
        # 合并单元格行号
        merge_row_no = 0
        for row in ws.iter_rows(min_row=row_use_name + 1, max_row=ws.max_row):
            step = step + 1
            ws.row_dimensions[row_use_name + step].height = 26
            column_no = 0
            for cell in row:
                column_no = column_no + 1
                # 判断【衬布颜色请按照面料颜色决定】，对应的D列需要背景色为黄色
                if str(cell.value).__contains__('衬布颜色请按照面料颜色决定'):
                    # 设置单元格填充颜色
                    cell.fill = header_fill
                    merge_row_no = step + 2
                cell.alignment = align_center
                cell.border = border
        # 合并单元格
        ws.merge_cells(start_row=merge_row_no, start_column=1, end_row=merge_row_no, end_column=column_no)

        # 设置打印标题和打印列
        # ws.print_title_rows = '3:1'
        # ws.print_title_cols = "A:E"

        # 设置打印A4纵向
        ws.set_printer_settings(ws.PAPERSIZE_A4, ws.ORIENTATION_PORTRAIT)

        # 所有列设置为一页 逆向思维,先缩放到页面 然后适合高度改为FLASE
        ws.sheet_properties.pageSetUpPr.fitToPage = True  # 此行必须设置
        ws.page_setup.fitToHeight = False

        ws.print_options.gridLines = True  # 页面设置->工作表->网格线

        # 页边距
        ws.page_margins.left = 0.1  # 左
        ws.page_margins.right = 0.1  # 右
        ws.page_margins.top = 0.1  # 上
        ws.page_margins.bottom = 0.1  # 下
        ws.page_margins.header = 0.1  # 页眉
        ws.page_margins.footer = 0.1  # 页脚

        # 保存
        wb.save(io)

    def file_to_dataframe(self, io, lfile):
        pdf = pdfplumber.open(io)
        colorlist= []
        pdf_title = []
        colorlist = []
        # 打开电子发票的PDF文件  
        for page in pdf.pages:
            # 提取文本内容 
            text = page.extract_text()
            if not text.__contains__('TRIMS'):
                continue
            # 提取表格2中非None的颜色列表
            colorlist_tmp = [i for i in page.extract_tables()[1][0] if i != None and i != '']
            # 提取表格2中的数据
            colorlist = colorlist_tmp[1:]
            if len(colorlist) == 0:
                colorlist = [i for i in page.extract_tables()[1][1] if i != None and i != '']
            pdf_data = page.extract_tables()[1][1:]
            pdf_title = np.append(['用料名称', '供应商', '备注', '有效幅宽/规格'], colorlist)
            if len(pdf_data[0]) != len(pdf_title):
                pdf_title = np.append(['用料名称', '供应商', '有效幅宽/规格'], colorlist)
            # 去掉数据尾部的空值
            for i in range(len(pdf_data)):
                pdf_data[i] = pdf_data[i][:len(pdf_title)]
            data_df = pd.DataFrame(data=pdf_data, columns=pdf_title)
            data_df.replace('\n', '', regex=True, inplace=True)
            break
        # 根据颜色追加列
        self.add_data_title = self.add_data_title_1 + colorlist + self.add_data_title_2
        # 追加品号title
        add_title = []
        for item in colorlist:
            add_title.append('品号')
        table_data = pd.DataFrame(None, columns=self.add_data_title)
        for item in self.add_data_title:
            if item in pdf_title:
                table_data.loc[:, item] = data_df[item]
        # 追加款号行和title行
        title_list = []
        first_row_data = []
        second_row_data = self.add_data_title_1 + add_title + self.add_data_title_2
        for index in range(len(second_row_data)):
            tmp_val = ''
            if index == 0:
                tmp_val = '客户：IZAC'
            elif index == 1:
                tmp_val = '款号：' + lfile
            first_row_data.append(tmp_val)
        title_list.append(first_row_data)
        title_list.append(second_row_data)
        # 首行数据只保留颜色值
        third_row_data = []
        for item in self.add_data_title:
            tmp_item = item
            if item in self.add_data_title_1 or item in self.add_data_title_2:
                tmp_item = ''
            third_row_data.append(tmp_item)
        title_list.append(third_row_data)
        now_row = pd.DataFrame(title_list, columns=self.add_data_title)
        table_data = pd.concat([now_row, table_data]).reset_index(drop=True)
        # 追加空白行及固定行
        blank_row = []
        for i in self.add_data_title:
            blank_row.append('')
        for i in range(0, 8):
            table_data = pd.concat([table_data, pd.DataFrame([blank_row], columns=self.add_data_title)]).reset_index(drop=True)
        # 尾部追加固定内容
        add_data =  pd.DataFrame(None, columns=self.add_data_title)
        # 用料名称固定列
        m_name = ['衬布颜色请按照面料颜色决定，如果面料是深色用黑色衬，如果面料是浅色用白色衬布', '兜布', '有纺衬', '无纺衬', '拉丝衬', '无纺有胶衬', '无纺纸衬', '马鬃', '胸棉', '袖山棉', '拉丝直条', '拉丝斜条', '双面胶', '直条', '小白带', '加丝中打条', '洗涤', '商标', '合缝线', '锁眼线']
        c_code = ['', '', 'FW2157 - 黑色/白色', 'XH-5050 - 黑色/白色', 'XH-NP5050 - 黑色/白色', 'BW70 - 灰色/白色', '254 - 灰色/白色', 'FC0179 150 8010 - 本色', 'AH-80 - 黑色/白色', '688-80 - 黑色/白色', '9332-1 - 黑色/白色', '7158 - 灰色/白色', '双面胶 - 透明', '5850-1 - 黑色/白色', 'YL-C03 - 黑色/白色', 'ZD-3030 - 黑色/白色', '', '', '2974100', '2974080']
        vender = ['', '', '库夫纳', '金林', '金林', '科德宝', '科德宝', '库夫纳', '佳峰', '佳峰', '鑫海', '齐祥', '鑫海', '鑫海', '桥新', '齐祥', '', '', '高士', '高士']
        specs = ['', '', '150cm', '100cm', '100cm', '90cm', '100cm', '150cm', '100cm', '100cm', '1cm', '1.5cm', '1cm', '2cm', '0.3cm', '2cm+1.5cm', '', '', 'TEX24/27', 'TEX40']
        part = ['', '腰兜附上一层，内部兜袋，前肩条+前袖笼上端条+后袖笼条', '前片+贴边+领面/领座+大小袖山+后袖笼', '马面下+后下摆+后开祺+袖口+腰兜口+省尖+兜位+台场', '贴边领口+前片领口+下摆圆+后肩+马面上', '胸兜牌', '里兜牙+三角牌', '主胸鬃+挺肩鬃', '胸衬', '袖山棉条', '止口+肩缝', '后中缝+侧缝+外袖缝', '', '驳口条', '袖笼', '', '夹入穿者左侧内兜垫带，见工艺指示', '穿者左侧内贴边，见工艺指示', '面合缝+里合缝+打结', '扣眼']
        consumption = ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '1', '1', '', '']
        if lfile.__contains__('PANT'):
            m_name = ['衬布颜色请按照面料颜色决定，如果面料是深色用黑色衬，如果面料是浅色用白色衬布', '兜布', '无纺衬', '有纺衬', '板带衬', '拉丝斜条', '商标', '洗涤', '合缝线', '码边线', '锁眼线']
            c_code = ['', '', 'XH-5050 - 黑色/白色', 'HM050 - 黑色/白色', '4947 - 黑色/白色', '7158 - 灰色/白色', '', '', '2974100 - 顺面料色', '8754140 - 顺面料色', '2974080 - 顺面料色']
            vender = ['', '', '金林', '恒明', '', '齐祥', '', '', '高士', '高士', '高士']
            specs = ['', '', '100cm', '150cm', '1cm', '1.5cm', '', '', 'TEX24/27', 'TEX21', 'TEX40']
            part = ['', '前后兜袋', '门刀+门襟+后兜牙/后兜口+侧兜口贴+(腰里+腰面没有皮筋部分）', '腰面先粘一层无纺衬再粘有纺衬', '绊带', '后档', '位置见工艺指示', '位置见工艺指示', '面合缝+里合缝+打结', '码边', '扣眼']
            consumption = ['', '', '', '', '', '', '1', '1', '', '', '']
        add_data.loc[:, '用料名称'] = m_name
        add_data.loc[:, '供应商'] = vender
        add_data.loc[:, '有效幅宽/规格'] = specs
        add_data.loc[:, '使用部位'] = part
        add_data.loc[:, '单耗'] = consumption
        for item in self.add_data_title:
            if item not in self.add_data_title_1 and item not in self.add_data_title_2 and len(item.strip()) != 0:
                add_data.loc[:, item] = c_code
        table_data = pd.concat([table_data, add_data]).reset_index(drop=True)
        # 导出excel
        excelUrl = self.trim_list_file_finish + '\\' + lfile + '.xlsx'
        writer = pd.ExcelWriter(excelUrl, engine='xlsxwriter')
        table_data.to_excel(writer, lfile, index=False, header=None)
        writer.save()
def gui_start():
    VAS = VAS_GUI()
    VAS.get_files()


if __name__ == '__main__':
    gui_start()
