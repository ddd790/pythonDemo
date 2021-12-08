# -*- coding:utf-8 -*-
import os
import pandas as pd
import datetime
import pymssql
from tkinter import *


class VAS_GUI():
    # 批量获取服务器数据，进行累加操作
    def commit_batch(self):
        print('数据操作进行中......')
        # sql服务器名
        self.serverName = '192.168.0.6'
        # 登陆用户名和密码
        self.userName = 'sa'
        self.passWord = 'MS_guanli09'

        # 数据库的列
        self.dbCol = ['订单级别', '验货情况', '排产月份', '面料情况', '预计采购单耗', '是否有工艺', '样品情况', '订单PO号', '客户款式', '款式缩写', '面料', 'ITEM', 'ZROH', '面料供应商', '面料成份', '尺寸表样板类型', '款式', '品名', '面料描述', '颜色', '类别', '交期确认', '样卡状态', '工厂FACTORY', '目的地', '品牌', '品牌商标', '加工合同号', '辅料表面料描述', '是否半里子', '前身里代号', '前身里料品色号', '袖里代号', '袖里料品色号', '后袖笼拼接料代码', '后袖笼拼接料品色号', '第三种里料代码', '第三种里料品色号', '扣代号', '国外扣供应商品号色号', '内扣或两种以上扣代号', '内扣或两种以上扣型号', '内部辅料档', '胸衬', '领呢', '领座', '上衣口袋布成份', '肩垫', '袖笼条',
                      '防抻条', '第二种以上小面料', '特殊用料', '订单COMMENTS', '钎子', '拉链', '裤膝', '裤子兜布', '腰里代码', '腰里明细', '腰衬', '腰面夹牙腰面包条', '成品胸衬板型', '成品袖笼条板型', '上衣上线日', '上衣班组', '裤子上线日', '裤子班组', '整熨开始日', 'COMMENTS备注', '面料价格单位', '衣架型号', '是否有裁单', '一般贸易手册号', '胸斗牌衬尺寸', '面料样', '新原则生产周期', '辅料到料时间', '面料清关单号', '辅料清关单号', '供生产部调整工厂使用的截止日期', '订单式采购', 'BOM检查', 'TINA是否报工厂', '自动编号', '季节号', '加工费币种', 'PO号批注', '报客户工厂', '辅料情况', '扣子情况', 'VAS情况', '衣架情况', '面料状态', '变更说明', '英文品名', '集装箱到场日', '运输方式']

        # 查询目前数据库所有数据
        self.select_all_data()
        # print(self.old_all_data)
        csv = self.old_all_data

        for csvIdx in range(0, len(csv)):
            tempCsvVal = csv.iloc[csvIdx]
            # print(tempCsvVal)
            for tempIdx in range(0, len(self.dbCol)):
                tempCsvVal[tempIdx] = str(tempCsvVal[tempIdx]).strip()
                if len(tempCsvVal[tempIdx]) == 5 and self.is_number(tempCsvVal[tempIdx][0:5]) and tempCsvVal[tempIdx][0:1] == '4':
                    print(str(tempCsvVal['自动编号']) + '---' +
                          self.dbCol[tempIdx] + '---' + str(tempCsvVal[tempIdx]))

        print('已经完成数据操作！')

    def select_all_data(self):
        # 建立连接并获取cursor
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, "ESApp1")
        cursor = conn.cursor()
        strCol = ",".join(str(i) for i in self.dbCol)
        select_sql = 'select ' + strCol + ' from 一部PO大表管理_明细'
        cursor.execute(select_sql)
        row = cursor.fetchall()
        self.old_all_data = pd.DataFrame(data=list(row), columns=self.dbCol)
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
    VAS.commit_batch()


if __name__ == '__main__':
    gui_start()
