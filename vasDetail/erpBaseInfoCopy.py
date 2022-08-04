# coding=utf8
from email import charset
import pandas as pd
import datetime
import pymssql
import numpy as np
import math
from tkinter import *


class VAS_GUI():
    # 批量获取服务器数据，进行累加操作
    def get_datas(self):
        print('数据操作进行中......')
        # erp服务器名
        self.serverNameErp = '192.168.0.11'
        # erp登陆用户名和密码
        self.userNameErp = 'sa'
        self.passWordErp = 'jiangbin@007'
        # erp数据库名
        self.dbNameErp = 'MDFNEW'
        # sql服务器名
        self.serverName = '192.168.0.6'
        # 登陆用户名和密码
        self.userName = 'sa'
        self.passWord = 'MS_guanli09'
        # 数据库名
        self.dbName = 'ESApp1'
        # 工厂信息列
        # 追加的dataFrame的title 加工工厂 from及条件
        self.add_factory_title = ['ERPID', '创建日期', '工厂代码', '工厂名称', '英文名称', '工厂简称', '所在国家', '所在地区', '法人代表', '联系人', '地址', '英文地址', '电话',
                                  '传真', '邮编', '电子邮件', '主要产品', '主要客户', '人民币开户', '帐号', '税号', '工厂类别', '联系人一', '电话一', '传真一', '手机一',
                                  '邮箱一', '备注', '联系人二', '电话二', '传真二', '手机二', '邮箱二', '付款方式', '收款银行', '开户银行', '收款行地址', 'SwiftCode',
                                  '每月产量', '合作日期', '工商注册日期', '是否停用', '停用日期', '停用原因', 'P_KEY']
        self.select_sql_factory = '''SELECT f.ID as ERPID,f.CREATEDATE as 创建日期,f.Code as 工厂代码,f.Name as 工厂名称,f.ENName as 英文名称,f.FatherDept as 工厂简称, a.Name AS 所在国家, b.Name AS 所在地区,f.LawDelegate as 法人代表,f.LinkMan as 联系人,f.CNAddress as 地址,f.ENAddress as 英文地址,f.Telephone as 电话,f.Fax as 传真,f.PostCode as 邮编,f.Email as 电子邮件,f.Product as 主要产品,f.Customer as 主要客户,f.RMBAccount as 人民币开户,f.Account as 帐号,f.TaxCode as 税号,f.CustomType as 工厂类别,f.Relation1 as 联系人一,f.Phone1 as 电话一,f.Faxs1 as 传真一,f.HandPhone1 as 手机一,f.Email1 as 邮箱一,f.Remark as 备注,f.Relation2 as 联系人二,f.Phone2 as 电话二,f.Faxs2 as 传真二,f.HandPhone2 as 手机二,f.Email2 as 邮箱二,c.Name AS 付款方式,d.Name AS 收款银行,e.Name AS 开户银行,f.BankAddress AS 收款行地址,f.SwiftCode,f.OutPut AS 每月产量,f.CooperationDate AS 合作日期,f.RegisterDate AS 工商注册日期,
        CASE
        WHEN f.IsStop = 1 THEN '是' ELSE '否' 
        END AS 是否停用,f.StopDate AS 停用日期,f.StopRemark AS 停用原因, concat(f.Name, f.Account) AS P_KEY FROM Factory f 
        LEFT JOIN SelectInfo a ON a.ID = f.ProduceArea
        LEFT JOIN SelectInfo b ON b.ID = f.FactoryArea
        LEFT JOIN SelectInfo c ON c.ID = f.BalanceType
        LEFT JOIN SelectInfo d ON d.ID = f.Bank
        LEFT JOIN SelectInfo e ON e.ID = f.OpenBank'''

        # 追加的dataFrame的title 客户基础信息
        self.add_customer_title = ['ERPID', '客户简码', '客户代码', '发展日期', '客户名称', '英文名称', '客户简称', '所属地区', '地址', '英文地址', '电话', '成立日期',
                                   '分管部门', '传真', '电子邮件', '网址', '备注', '电话一', '手机一', '邮箱一', '传真一', '联系人二', '电话二', '传真二', '手机二', '邮箱二', '联系人一']
        self.select_sql_customer = '''SELECT a.ID as ERPID,a.AbCode as 客户简码, a.CustomNo as 客户代码,a.DevelopmentDate as 发展日期,a.CustomName as 客户名称,a.EnglishName as 英文名称,a.AbName as 客户简称,b.Name as 所属地区,a.Address as 地址,a.EnglishAddress as 英文地址, a.TelNo as 电话,a.FoundDate as 成立日期,a.ManageDepart as 分管部门,a.FaxNo as 传真,a.Email as 电子邮件,a.WebSite as 网址,a.Remark as 备注, a.Phone1 as 电话一, a.HandPhone1 as 手机一, a.Email1 as 邮箱一,a.Faxs1 as 传真一,  a.Relation2 as 联系人二, a.Phone2 as 电话二, a.Faxs2 as 传真二, a.HandPhone2 as 手机二, a.Email2 as 邮箱二, a.Relation1 as 联系人一
        FROM Custom a LEFT JOIN SelectInfo b ON b.ID = a.AreaID'''

        # 追加的dataFrame的title 供应商基础信息 （ProviderTypeFalg = 61 是面料供应商， ProviderTypeFalg = 62 是辅料供应商）
        self.add_provider_title = ['ERPID', '供应商代码', '供应商名称', '英文名称', '简称', '地址', '英文地址', '所属地区', '量产付款方式', '合同币种', '发展日期', '停用标志',
                                   '停用日期', '供应类型', '员工数量', '注册资金', '年销售额', '开户银行', '开户行地址', 'SwiftCode', '帐号', '电话', '传真',
                                   '邮编', '电子邮件', '法人代表', '供应商类别', '备注', '面辅料通用', '税号', '所属国别', '客户指定供应商', '供应商来源',
                                   '联系人一', '电话一', '传真一', '手机一', '邮箱一', '联系人二', '电话二', '传真二', '手机二', '邮箱二',
                                   '供应商是否为离岸帐户', '创建人', '最后修改人', '创建日期', '修改日期', 'P_KEY', '预付比例']
        self.select_sql_provider = '''SELECT a.ID as ERPID,a.ProviderNo as 供应商代码,a.ProviderName as 供应商名称,a.EnglishName as 英文名称,a.AbName as 简称,a.Address as 地址,a.EnglishAddress as 英文地址,b.Name as 所属地区,c.Name as 量产付款方式,d.Name as 合同币种,a.DevelopmentDate AS 发展日期, 
        CASE
        WHEN a.StopFlag = 1 THEN '是' ELSE '否' 
        END as 停用标志,a.StopDate as 停用日期,e.Name as 供应类型,a.EmployeeCount as 员工数量,a.TotalCapital as 注册资金,a.SaleOneYear as 年销售额, f.Name as 开户银行,f.BankAddress as 开户行地址,f.SwiftCode,f.Account as 帐号,a.TelephoneNo as 电话,a.FaxNo as 传真, a.PostCode as 邮编,a.Email as 电子邮件,a.ArtificialPerson as 法人代表,a.CustomType as 供应商类别,a.Remark as 备注,a.IsUniversal as 面辅料通用,a.TaxNo as 税号,a.ExtraField1 as 所属国别,a.IsAppoint as 客户指定供应商,a.ExtraField3 as 供应商来源, a.Relation1 as 联系人一, a.Phone1 as 电话一, a.Faxs1 as 传真一, a.HandPhone1 as 手机一, a.Email1 as 邮箱一,a.Relation2 as 联系人二,a.Phone2 as 电话二,a.Faxs2 as 传真二,a.HandPhone2 as 手机二,a.Email2 as 邮箱二,'' as 供应商是否为离岸帐户, ac.Name AS 创建人, am.Name AS 最后修改人, a.CreateDate AS 创建日期, a.LastModifiedDate AS 修改日期, concat(a.ProviderName, f.Account) AS P_KEY, a.AdvancedPaymentProportion AS 预付比例 FROM Provider a
        LEFT JOIN  SelectInfo b ON b.ID = a.Area
        LEFT JOIN  SelectInfo c ON c.ID = a.PaymentWay
        LEFT JOIN  SysMonetaryUnit d ON d.ID = a.MonetaryUnit
        LEFT JOIN  SelectInfo e ON e.ID = a.ProviderTypeFalg
        LEFT JOIN  (
        SELECT m.ProviderID, n.Name,m.BankAddress,m.SwiftCode,m.Account FROM PrvAccount m
        LEFT JOIN SelectInfo n ON n.ID = m.BankID WHERE m.IsDefault = 'True'
        ) as f ON f.ProviderID = a.ID
        LEFT JOIN AC_User ac ON a.CreateUserID = ac.ID
        LEFT JOIN AC_User am ON a.LastModifiedUserID = am.ID
        WHERE (ProviderTypeFalg = 61 OR ProviderTypeFalg = 62)'''

        # 追加的dataFrame的title 面辅料基础信息 （Type = 61 是面料， Type = 62 是辅料）
        self.add_material_title = ['ERPID', '物料编码', '物料简码', '物料品号', '景泰蓝品号', '英文名称', '物料类型', '物料大类', '物料小类', '采购币种', '单位', '开票品名',
                                   '报关品名描述', '密度', '克重', '成份', '成分中文', '采购备注', '颜色英', '颜色', '规格', '采购价', '采购转换比率', '是否停用',
                                   '供方', '常采供应商', '原始类别', '采购单位', '纱支', '建单日期', '最后修改日期', '建单人', '最后修改人']
        self.select_sql_material = '''SELECT a.ID as ERPID, 
	    CASE WHEN pt.Name IS NULL THEN '' ELSE pt.Name END +
        CASE WHEN pd.ProviderName IS NULL THEN '' ELSE pd.ProviderName END + 
        CASE WHEN a.Name IS NULL THEN '' ELSE a.Name END + 
        CASE WHEN mxt.SizeName IS NULL THEN '' ELSE REPLACE(mxt.SizeName,'?','') END + 
        CASE WHEN mxt.ColorName IS NULL THEN '' ELSE REPLACE(mxt.ColorName,'?','') END + 
        CASE WHEN mt.Name IS NULL THEN '' ELSE mt.Name END + 
        CASE WHEN mxt.Price IS NULL THEN '0' ELSE 
        CASE WHEN mxt.Price = 0 THEN '0' else
        CASE WHEN mxt.Price <= 0.0001 THEN '0.0001' ELSE 
        convert(nvarchar(50), CONVERT(decimal(20,10),mxt.Price))
        END
        END 
        END 
        AS 物料编码,
        a.CodePrefix as 物料简码,a.Name as 物料品号,a.SampleCode as 景泰蓝品号,a.EnName as 英文名称,
        CASE 
        WHEN a.Type = 61 THEN '面料' ELSE '辅料' 
        END as 物料类型, 
        mc.Name as 物料大类,sc.Name as 物料小类,mt.Name as 采购币种,mu1.Name as 单位,a.InvoiceName as 开票品名,a.MaterialInfo as 报关品名描述,a.Density as 密度,a.Breadth as 克重,a.Element as 成份,a.EnElement as 成分中文,a.PurchaseRemark as 采购备注,mxt.ColorNameEN as 颜色英,REPLACE(mxt.ColorName,'?','') as 颜色,REPLACE(mxt.SizeName,'?','') as 规格, 
        CASE WHEN mxt.Price = 0 OR mxt.Price IS NULL THEN 0 else
        CASE WHEN 
        mxt.Price <= 0.0001 THEN 0.0001 ELSE 
        mxt.Price
        end
        END as 采购价,a.ConvertRate as 采购转换比率, 
        CASE
        WHEN mxt.IsBreakDown = 1 THEN '是' ELSE '否' 
        END as 是否停用,pt.Name as 供方,pd.ProviderName as 常采供应商,si.Name as 原始类别,mu2.Name as 采购单位,a.Expand8 as 纱支,mxt.CreateDate as 建单日期,mxt.LastModifiedDate as 最后修改日期, mxt.CreateUserName as 建单人, am.Name as 最后修改人
        FROM MX_Material mxt
        LEFT JOIN MX_MaterialCategory a ON mxt.MaterialCategoryID = a.ID
        LEFT JOIN MaterialClass mc ON mc.ID = a.Class
        LEFT JOIN AC_User am ON mxt.LastModifiedUserID = am.ID
        LEFT JOIN MX_MaterialCategoryImage m ON a.ID = m.MaterialCategoryID
        LEFT JOIN view_MaterialSubClass sc ON a.SubClass = sc.ID
        LEFT JOIN view_MoneyType mt ON a.MonetaryUnit = mt.ID
        LEFT JOIN view_MaterialUnitNo mu1 ON a.UnitNo = mu1.ID
        LEFT JOIN view_MaterialUnitNo mu2 ON a.PurchaseUnit = mu2.ID
        LEFT JOIN view_ProviderType pt ON a.ProviderType = pt.ID
        LEFT JOIN view_Provider pd ON a.Provider = pd.ID
        LEFT JOIN selectinfo si ON a.PurchaseClass = si.ID
        WHERE (a.Type = 61 or a.Type = 62)'''

        # 自定义字段，需要对sql中的汉字进行转码的字段
        self.change_code_col = ['是否停用', '停用标志',
                                '物料类型', '报关品名描述', '颜色英', '颜色', '规格']

        # 操作勤哲的表格和sql对应的字典
        self.option_table_dic = {'D_Factory_Erp': self.select_sql_factory, 'D_Customer_Erp': self.select_sql_customer,
                                 'D_Provider_Erp': self.select_sql_provider, 'D_Material_Erp': self.select_sql_material}

        # 循环操作表格
        for key, value in self.option_table_dic.items():
            print('开始操作 ' + key + ' 表，请耐心等待！' +
                  str(datetime.datetime.now()).split('.')[0])
            self.select_column_value(key)
            option_data = self.select_erp_value(value)
            if not option_data.empty:
                self.update_db(key, option_data)

        # 删除重复key的数据
        print('开始删除物料重复数据！' + str(datetime.datetime.now()).split('.')[0])
        self.del_repeat_data()
        print('所有表操作完毕！' + str(datetime.datetime.now()).split('.')[0])

    def select_column_value(self, table_name):
        # 获取表的信息数据（列和类型）
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        select_sql = "select column_name,data_type from INFORMATION_SCHEMA.COLUMNS where Table_Name = '" + table_name + "'"
        cursor.execute(select_sql)
        row_list = list(cursor.fetchall())
        # 日期列
        self.datetime_val = []
        # title列
        self.add_data_title = []
        # title对应类型
        self.dic_add_col_data = {}
        for row in row_list:
            self.add_data_title.append(row[0])
            self.dic_add_col_data[row[0]] = row[1]
            if row[1] == 'datetime':
                self.datetime_val.append(row[0])
        cursor.close()
        conn.close()

    def select_erp_value(self, select_sql):
        # 根据表名和sql获取ERP信息的数据
        conn = pymssql.connect(
            self.serverNameErp, self.userNameErp, self.passWordErp, self.dbNameErp, charset='utf8')
        cursor = conn.cursor()
        cursor.execute(select_sql.encode("utf8"))
        row = cursor.fetchall()
        data_value = pd.DataFrame(
            data=list(row), columns=self.add_data_title[1:])
        cursor.close()
        conn.close()
        return data_value

    def update_db(self, tableName, data_value):
        dbCol = self.add_data_title[1:]
        self.strCol = ",".join(str(i) for i in self.add_data_title[1:])
        # 建立连接并获取cursor
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        cursor.execute('TRUNCATE TABLE ' + tableName)
        # 组装插入的值
        insertValue = []
        table_value = []
        # 主要转码的字段
        for change_code_row in self.change_code_col:
            if change_code_row in dbCol:
                data_value[change_code_row] = np.where(
                    data_value[change_code_row].notnull(), data_value[change_code_row], '')
                data_value[change_code_row] = data_value[change_code_row].map(
                    lambda x: x.encode('latin-1').decode('gbk'))
        if tableName == 'D_Material_Erp':
            data_value['物料编码'] = data_value['物料编码'].map(
                lambda x: self.del_zero(x))
        # 去除日期型的NAT数据
        for time_row in self.datetime_val:
            data_value[time_row] = pd.to_datetime(
                data_value[time_row]).dt.floor('d')
            data_value[time_row] = np.where(
                data_value[time_row].notnull(), data_value[time_row].dt.strftime('%Y-%m-%d %H:%M:%S'), None)
        table_value.append([tuple(None if isinstance(i, float) and math.isnan(
            i) else i for i in t) for t in data_value.values])
        for tabVal in table_value:
            insertValue += tabVal
        # print(insertValue)
        insertSql = 'INSERT INTO ' + tableName + \
            ' (' + self.strCol + ')' + ' VALUES ('
        for index in range(len(dbCol)):
            # 判断数据类型
            if self.dic_add_col_data[dbCol[index]] == 'int' or self.dic_add_col_data[dbCol[index]] == 'decimal':
                insertSql += '%d'
            else:
                insertSql += '%s'
            # 判断是否是最后一个
            if index != len(dbCol) - 1:
                insertSql += ', '
        insertSql += ')'
        cursor.executemany(insertSql, insertValue)
        conn.commit()
        conn.close()

    def del_zero(self, strNum):
        if strNum[-1] == '0' and (strNum[-2] != '币' and strNum[-2] != '元'):
            strNum = str(strNum).rstrip('0')
            strNum = str(strNum).rstrip('.')
        return strNum

    def del_repeat_data(self):
        # 建立连接并获取cursor
        conn = pymssql.connect(
            self.serverName, self.userName, self.passWord, self.dbName)
        cursor = conn.cursor()
        cursor.execute('''DELETE A FROM dbo.D_Material_Erp A INNER JOIN (SELECT MIN(MID) AS MID FROM dbo.D_Material_Erp WHERE 是否停用 = '否' GROUP BY 物料编码 HAVING COUNT(物料编码) > 1) B ON A.MID = B.MID''')
        conn.commit()
        conn.close()


def gui_start():
    VAS = VAS_GUI()
    VAS.get_datas()


if __name__ == '__main__':
    gui_start()
