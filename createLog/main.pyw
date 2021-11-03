from tkinter import *
from tkinter import ttk
import pyodbc
import tkinter.messagebox as tmessage
import pyperclip


class LOG_GUI():
    def __init__(self, init_window_name):
        self.init_window_name = init_window_name

    def set_init_window(self):
        # 设置标题
        self.init_window_name.title('EXCEL服务器工具！')
        # 设置窗口大小
        self.init_window_name.geometry('300x200')
        # tab页
        tab = ttk.Notebook(self.init_window_name, height=200, width=280)
        # 追加log的tab
        logFrame = Frame(tab)
        self.log_form_frame(logFrame)
        tab.add(logFrame, text="创建LOG表")
        # 操作表格的tab
        optionFrame = Frame(tab)
        self.option_table_frame(optionFrame)
        tab.add(optionFrame, text="操作表字段")
        # sheet2修改字段类型的tab
        sheet2ChangeFrame = Frame(tab)
        self.col_type_form_frame(sheet2ChangeFrame)
        tab.add(sheet2ChangeFrame, text="创建shee2子表")
        tab.pack()

    def option_table_frame(self, optionFrame):
        # 标签
        self.table_name_label = Label(optionFrame, text="表名：")
        self.table_name_label.grid(row=1, column=1)
        self.new_col_name_label = Label(optionFrame, text="新字段名：")
        self.new_col_name_label.grid(row=2, column=1)
        self.new_col_type_label = Label(optionFrame, text="新字段类型：")
        self.new_col_type_label.grid(row=3, column=1)
        self.option_type_label = Label(optionFrame, text="操作类型：")
        self.option_type_label.grid(row=4, column=1)
        self.old_col_name_label = Label(optionFrame, text="原字段名：")
        self.old_col_name_label.grid(row=5, column=1)

        # 录入框
        self.table_name_text = Text(optionFrame, width=20, height=1)
        self.table_name_text.grid(row=1, column=2)  # 表名数据录入框
        self.new_col_name_text = Text(optionFrame, width=20, height=1)
        self.new_col_name_text.grid(row=2, column=2)  # 新字段名数据录入框
        self.new_col_type_text = Text(optionFrame, width=20, height=1)
        self.new_col_type_text.grid(row=3, column=2)  # 新字段类型录入框
        self.variable = StringVar()
        self.variable.set("add")
        self.option_type_option = OptionMenu(
            optionFrame, self.variable, "add", "update", "delete")
        self.option_type_option.grid(row=4, column=2)  # 新字段类型录入框
        self.old_col_name_text = Text(optionFrame, width=20, height=1)
        self.old_col_name_text.grid(row=5, column=2)  # 原字段名录入框

        # 按钮
        self.commit_button = Button(optionFrame, text="操作表格",
                                    bg="lightblue", width=10, command=self.option_table)
        self.commit_button.grid(row=6, column=2)

    def log_form_frame(self, logFrame):
        # 标签
        self.main_table_label = Label(logFrame, text="主表名：")
        self.main_table_label.grid(row=1, column=1)
        self.sub_table_label = Label(logFrame, text="子表名：")
        self.sub_table_label.grid(row=2, column=1)
        self.poid_label = Label(logFrame, text="POID列：")
        self.poid_label.grid(row=3, column=1)

        # 录入框
        self.main_table_text = Text(logFrame, width=20, height=1)
        self.main_table_text.grid(row=1, column=2)  # 主表名数据录入框
        self.sub_table_text = Text(logFrame, width=20, height=1)
        self.sub_table_text.grid(row=2, column=2)  # 子表名数据录入框
        self.poid_text = Text(logFrame, width=20, height=1)
        self.poid_text.grid(row=3, column=2)  # poid录入框

        self.poid_label = Label(logFrame, text="※POID可以不填，默认是【自动编号】")
        self.poid_label.grid(row=4, column=2)

        # 按钮
        self.commit_button = Button(logFrame, text="生成log表",
                                    bg="lightblue", width=10, command=self.commit_form)
        self.commit_button.grid(row=5, column=2)

    def col_type_form_frame(self, typeFrame):
        # 标签
        self.main_table_sheet1_label = Label(typeFrame, text="sheet1明细表名：")
        self.main_table_sheet1_label.grid(row=1, column=1)

        # 录入框
        self.main_table_sheet1_text = Text(typeFrame, width=20, height=1)
        self.main_table_sheet1_text.grid(row=1, column=2)  # sheet1明细表名数据录入框

        # 按钮
        self.commit_button = Button(typeFrame, text="创建",
                                    bg="lightblue", width=10, command=self.col_type_form)
        self.commit_button.grid(row=3, column=2)

    def col_type_form(self):
        sheet1_table = self.main_table_sheet1_text.get(
            1.0, END).strip().replace("\n", "")
        if sheet1_table != '' and sheet1_table != '':
            # 确认信息，resurn 'True' or 'False'
            infoMess = tmessage.askyesno(
                title='提示', message='sheet1表：【' + sheet1_table + '】。\n是否确认要进行操作？')
            if infoMess:
                self.change_col_type(sheet1_table)
        else:
            tmessage.showerror('错误', '主表和子表的名字不能为空！')

    def commit_form(self):
        main_table = self.main_table_text.get(
            1.0, END).strip().replace("\n", "")
        sub_table = self.sub_table_text.get(
            1.0, END).strip().replace("\n", "")
        poid = self.poid_text.get(
            1.0, END).strip().replace("\n", "")
        if main_table == sub_table:
            tmessage.showerror('错误', '主表和子表的名字不能相同！')
        elif main_table != '' and sub_table != '':
            if poid == '':
                poid = '自动编号'
            # 确认信息，resurn 'True' or 'False'
            infoMess = tmessage.askyesno(
                title='提示', message='主表：【' + main_table + '】，子表：【' + sub_table + '】, POID：【' + poid + '】\n是否确认要进行操作？')
            if infoMess:
                self.create_log(main_table, sub_table, poid)
        else:
            tmessage.showerror('错误', '主表和子表的名字不能为空！')

    def option_table(self):
        table_name = self.table_name_text.get(
            1.0, END).strip().replace("\n", "")
        new_col_name = self.new_col_name_text.get(
            1.0, END).strip().replace("\n", "")
        new_col_type = self.new_col_type_text.get(
            1.0, END).strip().replace("\n", "")
        option_variable = self.variable.get()
        old_col_name = self.old_col_name_text.get(
            1.0, END).strip().replace("\n", "")
        if table_name != '' and new_col_name != '' and new_col_type != '':
            # 确认信息，resurn 'True' or 'False'
            infoMess = tmessage.askyesno(
                title='提示', message='操作表：【' + table_name + '】，字段名：【' + new_col_name + '】, 数据类型：【' + new_col_type + '】\n操作类型：【' + option_variable + '】, 原列名：【' + old_col_name + '】是否确认要进行操作？')
            if infoMess:
                self.option_col(table_name, new_col_name,
                                new_col_type, option_variable, old_col_name)
            tmessage.showinfo('操作成功', '表【' + table_name + '】已被修改！')
        else:
            tmessage.showerror('错误', '前三项为必填字段，不能为空！')

    def change_col_type(self, sheet1Table):
        try:
            cn = pyodbc.connect(
                'DRIVER={SQL Server};SERVER=192.168.0.11;DATABASE=ESApp1;UID=sa;PWD=jiangbin@007')
            cn.autocommit = True
            cr = cn.cursor()
            # 执行存储过程，创建sheet2子表
            cr.execute("set nocount on")
            cr.execute("exec changeColumnType '" + sheet1Table + "'")
            cr.close()
            cn.close()
            tmessage.showinfo('操作成功', '表【' + sheet1Table + '_sheet2】已被创建！')
        except:
            tmessage.showerror('错误', '人生苦短,数据库出错了,请稍后操作！')

    def option_col(self, table_name, new_col_name, new_col_type, option_variable, old_col_name):
        try:
            cn = pyodbc.connect(
                'DRIVER={SQL Server};SERVER=192.168.0.11;DATABASE=ESApp1;UID=sa;PWD=jiangbin@007')
            cn.autocommit = True
            cr = cn.cursor()
            # 执行存储过程，创建log相关表
            cr.execute("set nocount on")
            cr.execute("exec optionLogColumn '" + table_name + "','" + new_col_name +
                       "','" + new_col_type + "','" + option_variable + "','" + old_col_name + "'")
            cr.close()
            cn.close()
        except:
            tmessage.showerror('错误', '人生苦短,数据库出错了,请稍后操作！')

    def create_log(self, main_table, sub_table, poid):
        try:
            cn = pyodbc.connect(
                'DRIVER={SQL Server};SERVER=192.168.0.11;DATABASE=ESApp1;UID=sa;PWD=jiangbin@007')
            cn.autocommit = True
            cr = cn.cursor()
            # 检测是否存在数据库
            cr.execute("SELECT * FROM " + main_table + ";")
            cr.execute("SELECT * FROM " + sub_table + ";")
            # 执行存储过程，创建log相关表
            cr.execute("set nocount on")
            cr.execute("exec createLogTable '" +
                       main_table + "','" + sub_table + "'")
            # 创建触发器
            trigger_sql = """CREATE TRIGGER """ + sub_table + """_after_insert
                        on """ + sub_table + """
                        after insert
                        AS
                        BEGIN
                            DECLARE @ExcelServerRCID nvarchar(1000);
                            DECLARE @POID nvarchar(500);
                            SELECT @POID = """ + poid + """, @ExcelServerRCID = ExcelServerRCID from Inserted;
                            DECLARE @optionType nvarchar(6);
                            DECLARE @optionName nvarchar(500);
                            SELECT @optionName=操作人 from """ + main_table + """ where ExcelServerRCID = @ExcelServerRCID;
                            DECLARE @logNumber INT;
                            SELECT @logNumber = COUNT(*) FROM """ + sub_table + """_old_log where """ + poid + """ = @POID;
                            IF (@logNumber = 0)
                                BEGIN
                                    set @optionType = 'insert';
                                    insert into """ + sub_table + """_old_log SELECT * from inserted where ExcelServerRCID = @ExcelServerRCID;
                                    insert into """ + sub_table + """_Log SELECT *, @optionType AS 操作类型, GETDATE() AS 操作时间, @optionName AS 操作用户 from inserted where ExcelServerRCID = @ExcelServerRCID;
                                END
                            ELSE
                                BEGIN
                                    DECLARE @diffFlag INT;
                                    SELECT @diffFlag = COUNT(*) FROM (select * from """ + sub_table + """_old_log where ExcelServerRCID = @ExcelServerRCID AND """ + poid + """ = @POID EXCEPT select * from inserted where ExcelServerRCID = @ExcelServerRCID AND """ + poid + """ = @POID) as T;
                                    IF (@diffFlag = 1)
                                        BEGIN
                                            set @optionType = 'update';
                                            INSERT INTO """ + sub_table + """_Log SELECT *, @optionType AS 操作类型, GETDATE() AS 操作时间, @optionName AS 操作用户 from inserted WHERE ExcelServerRCID = @ExcelServerRCID AND """ + poid + """ = @POID;
                                        END
                                    DELETE FROM """ + sub_table + """_old_log WHERE ExcelServerRCID = @ExcelServerRCID AND """ + poid + """ = @POID;
                                    insert into """ + sub_table + """_old_log SELECT * from inserted;
                                END
                        END"""
            cr.execute(trigger_sql)
            # 创建【删除日志存储过程】
            del_proc_sql = """create proc """ + sub_table + """_Del_Log
                        as
                        BEGIN
                        DECLARE @logNumber INT;
                        SELECT @logNumber = COUNT(*) FROM """ + sub_table + """_old_log;
                        IF (@logNumber = 0)
                            BEGIN
                                INSERT INTO """ + sub_table + """_old_log SELECT * FROM """ + sub_table + """;
                            END
                        ELSE
                            BEGIN
                                SELECT row_number() OVER(ORDER BY (select 0)) as 'rowindexdel',* into #temp1 from """ + sub_table + """_old_log;
                                DECLARE @i_del int,@flag_del INT;
                                SELECT @flag_del = COUNT(rowindexdel)+1 FROM #temp1;
                                SET @i_del=1;
                                WHILE (@i_del<@flag_del)
                                    BEGIN
                                    DECLARE @ExcelServerRCID varchar(1000);
                                    DECLARE @POID nvarchar(500);
                                    SELECT @ExcelServerRCID = ExcelServerRCID, @POID = """ + poid + """ from #temp1 where rowindexdel=@i_del;
                                    DECLARE @delFlagNum int;
                                    SELECT @delFlagNum = COUNT(*) from """ + sub_table + """ WHERE ExcelServerRCID = @ExcelServerRCID AND """ + poid + """ = @POID;
                                    IF (@delFlagNum = 0)
                                        BEGIN
                                            INSERT INTO """ + sub_table + """_Log SELECT *, 'delete' AS 操作类型, GETDATE() AS 操作时间, 'TIMER' AS 操作用户 from """ + sub_table + """_old_log WHERE ExcelServerRCID = @ExcelServerRCID AND """ + poid + """ = @POID;
                                            DELETE FROM """ + sub_table + """_old_log WHERE ExcelServerRCID = @ExcelServerRCID AND """ + poid + """ = @POID;
                                        END
                                        SET @i_del=@i_del+1;
                                    END
                                    truncate table #temp1;
                                    DROP TABLE #temp1;
                            END
                        END"""
            cr.execute(del_proc_sql)
            cr.close()
            cn.close()
            pyperclip.copy("exec " + sub_table + "_Del_Log;")
            tmessage.showinfo('创建成功', '将【exec ' + sub_table +
                              '_Del_Log;】加入到【insertDelLog】存储过程中，并保存！\n（命令已经复制到剪贴板，直接粘贴到对应的位置即可！）')
        except:
            tmessage.showerror('错误', '人生苦短,数据库出错了,请稍后操作！')


def gui_start():
    init_window = Tk()  # 实例化出一个父窗口
    LOG = LOG_GUI(init_window)
    LOG.set_init_window()  # 设置根窗口默认属性

    init_window.mainloop()  # 父窗口进入事件循环，可以理解为保持窗口运行，否则界面不展示


if __name__ == '__main__':
    gui_start()
