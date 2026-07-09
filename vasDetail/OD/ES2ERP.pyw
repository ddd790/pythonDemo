import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import pandas as pd
import pyodbc
from sqlalchemy import create_engine, text
import pymysql

# -------------------------- 全局配置 --------------------------
# SQL Server固定信息
SQL_SERVER_CONFIG = {
    "host": "192.168.0.11",
    "port": 1433,
    "database": "ESApp1",
    "user": "sa",
    "password": "jiangbin@007"
}

# MySQL固定配置
MYSQL_SERVER_CONFIG = {
    "host": "127.0.0.1",
    "port": 3306,
    "database": "motives_erp_local",
    "user": "root",
    "password": "11111111"
}

# 新增：功能选项与MySQL表名的映射字典（核心修改）
FUNC_TABLE_MAP = {
    "供应商": "sys_supplier",
    "工厂": "sys_factory",
    "客户": "sys_customer",
    "货代": "sys_forwarder",
    "三方": "sys_third_company",
    "物料分类": "sys_material_type",
    "物料单位": "sys_biz_dict",
    "物料明细": "sys_material",
    "物料价格": "material_price"
}

# 新增：功能和勤哲表中的视图名的映射字典，
# 例：新建供应商对应view_supplier2ERP_replace视图,追加为view_material2ERP_append
FUNC_VIEW_MAP = {
    "供应商": "view_supplier2ERP",
    "工厂": "view_factory2ERP",
    "客户": "view_customer2ERP",
    "货代": "view_forwarder2ERP",
    "三方": "view_third2ERP",
    "物料分类": "view_material_type2ERP",
    "物料单位": "view_material_unit2ERP",
    "物料明细": "view_material2ERP",
    "物料价格": "view_material2ERP_base"
}

# -------------------------- 核心功能函数 --------------------------
def get_sqlserver_data(view_name, start_date=None):
    """连接SQL Server，查询指定视图数据并返回DataFrame"""
    conn_str = (
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={SQL_SERVER_CONFIG['host']},{SQL_SERVER_CONFIG['port']};"
        f"DATABASE={SQL_SERVER_CONFIG['database']};"
        f"UID={SQL_SERVER_CONFIG['user']};"
        f"PWD={SQL_SERVER_CONFIG['password']}"
    )
    
    try:
        conn = pyodbc.connect(conn_str, timeout=10)
        query = f"SELECT * FROM {view_name}"
        if start_date:
            query += f" WHERE create_at >= '{start_date}'"
        df = pd.read_sql(query, conn)
        conn.close()
        if df.empty:
            messagebox.warning("提示", f"视图{view_name}中无数据！")
            return None
        return df
    except Exception as e:
        messagebox.showerror("SQL Server连接/查询失败", f"错误信息：{str(e)}")
        return None

def import_to_mysql(df, mysql_host, mysql_db, mysql_user, mysql_pwd, table_name, if_exists="replace"):
    """
    将DataFrame导入MySQL指定表（核心修改：新增table_name参数）
    :param table_name: 目标MySQL表名（根据功能选择动态传入）
    """
    if df is None:
        return False
    
    try:
        engine_str = f"mysql+pymysql://{mysql_user}:{mysql_pwd}@{mysql_host}:{MYSQL_SERVER_CONFIG['port']}/{mysql_db}"
        engine = create_engine(engine_str, pool_pre_ping=True)
        
        with engine.begin() as conn:
            if if_exists == "replace":
                conn.exec_driver_sql("SET FOREIGN_KEY_CHECKS = 0;")
                if table_name == "sys_biz_dict":
                    conn.exec_driver_sql(f"DELETE FROM {table_name} WHERE dictionary = 'MaterialUnit';")
                else:
                    conn.exec_driver_sql(f"TRUNCATE TABLE {table_name};")
                conn.exec_driver_sql("SET FOREIGN_KEY_CHECKS = 1;")
            
            df.to_sql(
                name=table_name,
                con=conn,
                if_exists="append",
                index=False,
                chunksize=1000
            )
        messagebox.showinfo("成功", f"数据已{if_exists}模式导入MySQL表【{table_name}】！")
        return True
    except Exception as e:
        messagebox.showerror("MySQL导入失败", f"错误信息：{str(e)}")
        return False

def material_price_import(mysql_host, mysql_db, mysql_user, mysql_pwd):
    """物料价格导入：从MySQL取物料id/code，从SQL Server取价格，拼接后导入material_price表"""
    try:
        # 1. 从MySQL查询sys_material表的id和code
        engine_str = f"mysql+pymysql://{mysql_user}:{mysql_pwd}@{mysql_host}:{MYSQL_SERVER_CONFIG['port']}/{mysql_db}"
        engine = create_engine(engine_str, pool_pre_ping=True)

        with engine.connect() as conn:
            sys_material_df = pd.read_sql(text("SELECT id, code FROM sys_material"), conn)

        if sys_material_df.empty:
            messagebox.showwarning("提示", "MySQL的sys_material表中无数据！")
            return False

        # 2. 从SQL Server查询view_material2ERP_base的code、purchase_currency、price
        view_name = FUNC_VIEW_MAP["物料价格"]
        conn_str = (
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={SQL_SERVER_CONFIG['host']},{SQL_SERVER_CONFIG['port']};"
            f"DATABASE={SQL_SERVER_CONFIG['database']};"
            f"UID={SQL_SERVER_CONFIG['user']};"
            f"PWD={SQL_SERVER_CONFIG['password']}"
        )
        conn2 = pyodbc.connect(conn_str, timeout=10)
        base_df = pd.read_sql(f"SELECT code, purchase_currency, price FROM {view_name}", conn2)
        conn2.close()

        if base_df.empty:
            messagebox.showwarning("提示", f"视图{view_name}中无数据！")
            return False

        # 3. 按code拼接，构建新DataFrame
        merged = pd.merge(sys_material_df, base_df, on="code", how="inner")
        if merged.empty:
            messagebox.showwarning("提示", "没有匹配到任何物料价格数据！")
            return False

        result_df = pd.DataFrame({
            "create_at": "2026-1-1",
            "create_by": 1,
            "update_at": "2026-1-1",
            "update_by": 1,
            "version": 1,
            "currency": merged["purchase_currency"],
            "disabled": 0,
            "no": "WLQZ260101" + merged["id"].astype(str).str.zfill(5),
            "price": merged["price"],
            "price_type": "STANDARD",
            "remark": "",
            "status": "ENABLED",
            "customer_id": None,
            "material_id": merged["id"]
        })

        # 4. 导入MySQL material_price表
        with engine.connect() as conn:
            trans = conn.begin()
            conn.exec_driver_sql("SET FOREIGN_KEY_CHECKS = 0;")
            conn.exec_driver_sql("TRUNCATE TABLE material_price;")
            result_df.to_sql(
                name="material_price",
                con=conn,
                if_exists="append",
                index=False,
                chunksize=1000
            )
            conn.exec_driver_sql("SET FOREIGN_KEY_CHECKS = 1;")
            trans.commit()

        messagebox.showinfo("成功", f"物料价格已导入，共 {len(result_df)} 条记录！")
        return True

    except Exception as e:
        messagebox.showerror("物料价格导入失败", f"错误信息：{str(e)}")
        return False


# -------------------------- 按钮点击事件 --------------------------
def new_button_click():
    """新建按钮点击事件：导入数据（覆盖模式）"""
    # 1. 获取用户输入的MySQL信息
    mysql_host = entry_host.get().strip()
    mysql_db = entry_db.get().strip()
    mysql_user = entry_user.get().strip()
    mysql_pwd = entry_pwd.get().strip()
    
    # 2. 获取用户选择的功能选项及对应表名（核心修改）
    selected_func = func_var.get()
    target_table = FUNC_TABLE_MAP[selected_func]  # 从映射字典中获取表名
    target_view = FUNC_VIEW_MAP[selected_func]  # 从映射字典中获取视图名
    
    # 3. 校验输入
    if not (mysql_db and mysql_user and mysql_pwd):
        messagebox.warning("提示", "请填写完整的MySQL数据库名、用户名、密码！")
        return
    
    # 物料价格特殊处理
    if selected_func == "物料价格":
        material_price_import(mysql_host, mysql_db, mysql_user, mysql_pwd)
        return
    
    # 4. 取数并导入（传入动态表名）
    df = get_sqlserver_data(target_view + '_replace')
    import_to_mysql(df, mysql_host, mysql_db, mysql_user, mysql_pwd, target_table, if_exists="replace")

def append_button_click():
    """追加按钮点击事件：导入数据（追加模式）"""
    mysql_host = entry_host.get().strip()
    mysql_db = entry_db.get().strip()
    mysql_user = entry_user.get().strip()
    mysql_pwd = entry_pwd.get().strip()
    
    selected_func = func_var.get()
    target_table = FUNC_TABLE_MAP[selected_func]
    target_view = FUNC_VIEW_MAP[selected_func]
    
    if not (mysql_db and mysql_user and mysql_pwd):
        messagebox.warning("提示", "请填写完整的MySQL数据库名、用户名、密码！")
        return
    
    def on_confirm():
        date_str = entry_date.get().strip()
        if not date_str:
            messagebox.warning("提示", "请输入日期！")
            return
        try:
            datetime.strptime(date_str, "%Y-%m-%d")
        except ValueError:
            messagebox.warning("提示", "日期格式不正确，请输入YYYY-MM-DD格式！")
            return
        date_dialog.destroy()
        df = get_sqlserver_data(target_view + '_replace', start_date=date_str)
        import_to_mysql(df, mysql_host, mysql_db, mysql_user, mysql_pwd, target_table, if_exists="append")
    
    def on_cancel():
        date_dialog.destroy()
    
    date_dialog = tk.Toplevel()
    date_dialog.title("输入日期")
    date_dialog.geometry("300x150")
    date_dialog.resizable(False, False)
    date_dialog.transient(root)
    date_dialog.grab_set()
    
    label_date = ttk.Label(date_dialog, text="请输入日期：")
    label_date.pack(pady=15)
    
    today_str = datetime.now().strftime("%Y-%m-%d")
    entry_date = ttk.Entry(date_dialog, width=20)
    entry_date.insert(0, today_str)
    entry_date.pack(pady=5)
    
    frame_btn = ttk.Frame(date_dialog)
    frame_btn.pack(pady=15)
    
    btn_confirm = ttk.Button(frame_btn, text="确认", width=10, command=on_confirm)
    btn_confirm.pack(side="left", padx=10)
    
    btn_cancel = ttk.Button(frame_btn, text="取消", width=10, command=on_cancel)
    btn_cancel.pack(side="right", padx=10)

# -------------------------- GUI界面构建 --------------------------
if __name__ == "__main__":
    root = tk.Tk()
    root.title("数据导入工具")
    root.geometry("450x500")
    root.resizable(False, False)

    # 1. 数据库连接区域
    frame_db = ttk.LabelFrame(root, text="MySQL数据库连接", padding=(20, 10))
    frame_db.pack(fill="x", padx=20, pady=10)

    label_host = ttk.Label(frame_db, text="主机地址：")
    label_host.grid(row=0, column=0, padx=5, pady=5, sticky="w")
    entry_host = ttk.Entry(frame_db, width=30)
    entry_host.insert(0, MYSQL_SERVER_CONFIG['host'])
    entry_host.grid(row=0, column=1, padx=5, pady=5)

    label_db = ttk.Label(frame_db, text="数据库名：")
    label_db.grid(row=1, column=0, padx=5, pady=5, sticky="w")
    entry_db = ttk.Entry(frame_db, width=30)
    entry_db.insert(0, MYSQL_SERVER_CONFIG['database'])
    entry_db.grid(row=1, column=1, padx=5, pady=5)

    label_user = ttk.Label(frame_db, text="用户名：")
    label_user.grid(row=2, column=0, padx=5, pady=5, sticky="w")
    entry_user = ttk.Entry(frame_db, width=30)
    entry_user.insert(0, MYSQL_SERVER_CONFIG['user'])
    entry_user.grid(row=2, column=1, padx=5, pady=5)

    label_pwd = ttk.Label(frame_db, text="密码：")
    label_pwd.grid(row=3, column=0, padx=5, pady=5, sticky="w")
    entry_pwd = ttk.Entry(frame_db, width=30, show="*")
    entry_pwd.insert(0, MYSQL_SERVER_CONFIG['password'])
    entry_pwd.grid(row=3, column=1, padx=5, pady=5)

    # 2. 功能选择区域
    frame_func = ttk.LabelFrame(root, text="合作单位", padding=(20, 10))
    frame_func.pack(fill="x", padx=20, pady=10)

    func_var = tk.StringVar(value="供应商")  # 默认选中供应商
    rb_supplier = ttk.Radiobutton(frame_func, text="供应商", variable=func_var, value="供应商")
    rb_supplier.grid(row=0, column=0, padx=10, pady=1)
    rb_factory = ttk.Radiobutton(frame_func, text="工厂", variable=func_var, value="工厂")
    rb_factory.grid(row=0, column=1, padx=10, pady=1)
    rb_customer = ttk.Radiobutton(frame_func, text="客户", variable=func_var, value="客户")
    rb_customer.grid(row=0, column=2, padx=10, pady=1)
    rb_logistics = ttk.Radiobutton(frame_func, text="货代", variable=func_var, value="货代")
    rb_logistics.grid(row=0, column=3, padx=10, pady=1)
    rb_third = ttk.Radiobutton(frame_func, text="三方", variable=func_var, value="三方")
    rb_third.grid(row=0, column=4, padx=10, pady=1)

    # 3. 物料相关区域
    frame_material = ttk.LabelFrame(root, text="物料相关", padding=(20, 10))
    frame_material.pack(fill="x", padx=20, pady=10)

    rb_material_type = ttk.Radiobutton(frame_material, text="物料分类", variable=func_var, value="物料分类")
    rb_material_type.grid(row=0, column=0, padx=10, pady=1)
    rb_material_unit = ttk.Radiobutton(frame_material, text="物料单位", variable=func_var, value="物料单位")
    rb_material_unit.grid(row=0, column=1, padx=10, pady=1)
    rb_material = ttk.Radiobutton(frame_material, text="物料明细", variable=func_var, value="物料明细")
    rb_material.grid(row=0, column=2, padx=10, pady=1)
    rb_material_price = ttk.Radiobutton(frame_material, text="物料价格", variable=func_var, value="物料价格")
    rb_material_price.grid(row=0, column=3, padx=10, pady=1)

    # 4. 操作区域
    frame_operate = ttk.Frame(root, padding=(20, 10))
    frame_operate.pack(fill="x", padx=20, pady=20)

    btn_new = ttk.Button(frame_operate, text="导入", width=15, command=new_button_click)
    btn_new.grid(row=0, column=0, padx=20)

    btn_append = ttk.Button(frame_operate, text="追加", width=15, command=append_button_click)
    btn_append.grid(row=0, column=1, padx=20)

    root.mainloop()