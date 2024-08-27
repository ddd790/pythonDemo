import os
import pandas as pd
import pymssql
import datetime
import shutil

class VAS_GUI():
    # 批量获取服务器数据，进行累加操作
    def get_files(self):
        print('数据操作进行中......' + str(datetime.datetime.now()).split('.')[0])
        # sql服务器名
        self.serverName = '192.168.0.11'
        # 登陆用户名和密码
        self.userName = 'sa'
        self.passWord = 'jiangbin@007'
        # 数据库名
        self.dbName = 'ESApp1'
        # 追加的dataFrame的title
        self.add_data_title = ['用料', '用料名称', '实测幅宽', '实测缩率', '报价单耗', 
                               '订料单耗', '客户', '款号','类别', '面料成份', '单耗负责人', '数量', '面料号', 'V_Key', 'FileName', 'Version']
        # 数字类型的字段
        self.number_item = ['Version', '报价单耗', '订料单耗']
        self.local_excel_detail_file = 'd:\\大货面辅料单耗用量表'

        # 查询已存在的记录
        self.select_old_value()
        # 根据V_Key，找到对应的version
        self.old_version = self.old_all_data.set_index("V_Key")["Version"].to_dict()
        # 最终dataframe
        self.table_data = pd.DataFrame(data=None, columns=self.add_data_title)
        self.table_value = []
        # 循环文件，处理合并
        for lroot, ldirs, lfiles in os.walk(self.local_excel_detail_file):
            for lfile in lfiles:
                if not str(lfile).__contains__('~'):
                    print('文件名：' + str(lfile).split('.')[0])
                    self.file_to_dataframe(os.path.join(lroot, lfile), str(lfile).split('.')[0])
        # 更新数据库
        self.update_db()
        # 删除目录内文件
        if os.path.exists(self.local_excel_detail_file):
            shutil.rmtree(self.local_excel_detail_file)
        os.mkdir(self.local_excel_detail_file)
        print('已经完成操作！' + str(datetime.datetime.now()).split('.')[0])
        input('按回车退出 ')

    def file_to_dataframe(self, io, lfile):
        # 读取excel
        df = pd.read_excel(io, sheet_name=0, nrows=100)
        # '客户', '款号','类别', '面料成份', '单耗负责人', '数量', '面料号'
        customer, style, category, composition, user, qty, fNum = df.iloc[3, 2], df.iloc[3, 4], df.iloc[3, 8], \
                                df.iloc[3, 13], df.iloc[3, 15], df.iloc[4, 4], df.iloc[4, 13]
        # v_key记录版本的key
        v_key = str(style) + '_' + str(category)
        # 取出用料中含有面料的行（10-21行中）
        rows_of_interest = df.iloc[9:21, 0]
        contains_fabric = rows_of_interest.str.contains('面料', case=False)
        row_indices_with_fabric = contains_fabric[contains_fabric].index.tolist()
        # 用料相关信息
        table_data = pd.DataFrame(df.iloc[row_indices_with_fabric, [0,1,4,6,10,12]])
        table_data.columns = self.add_data_title[0:6]
        table_data.fillna(0, inplace=True)
        # 订单相关信息
        table_data['客户'] = customer
        table_data['款号'] = style
        table_data['类别'] = category
        table_data['面料成份'] = composition
        table_data['单耗负责人'] = user
        table_data['数量'] = qty
        table_data['面料号'] = fNum
        table_data['V_Key'] = v_key
        table_data['FileName'] = str(lfile).split('.')[0]
        table_data['Version'] = 0 if self.old_version.get(v_key) == None else int(self.old_version.get(v_key)) + 1
        table_data['CreateDate'] = str(datetime.datetime.now()).split('.')[0]
        table_data.fillna('', inplace=True)
        self.table_value.append([tuple(row) for row in table_data.values])
        # 将文件转存到bak 目录下
        bak_file =  'd:\\单耗用量表_bak'
        if not os.path.exists(bak_file):
            os.mkdir(bak_file)
        shutil.copy(io, bak_file + '\\' + str(lfile) + '.xls')

    def select_old_value(self):
        # 建立连接并获取PO数据
        conn = pymssql.connect(self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        strCol = ",".join(str(i) for i in self.add_data_title)
        select_sql = 'select ' + strCol + ' from D_三部大货面辅料单耗用量表'
        cursor.execute(select_sql)
        row = cursor.fetchall()
        self.old_all_data = pd.DataFrame(data=list(row), columns=self.add_data_title)
        cursor.close()
        conn.close()

    def update_db(self):
        dbCol = self.add_data_title[:]
        dbCol.append('CreateDate')
        # 建立连接并获取cursor
        conn = pymssql.connect(self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        # 组装插入的值
        insertValue = []
        for tabVal in self.table_value:
            insertValue += tabVal
        insertSql = 'INSERT INTO D_三部大货面辅料单耗用量表 (' + (",".join(str(i) for i in dbCol)) + ') VALUES ('
        for colVal in dbCol:
            if colVal == 'CreateDate':
                insertSql += '%s'
            elif colVal in self.number_item:
                insertSql += '%d, '
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
