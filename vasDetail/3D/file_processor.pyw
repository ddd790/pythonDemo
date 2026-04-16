import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import pandas as pd

class FileProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("文件处理工具")
        self.root.geometry("700x300")
        self.root.resizable(False, False)
        
        # 变量
        self.selected_file = tk.StringVar()
        self.function_var = tk.StringVar(value="客户文件读取")
        
        # 创建UI
        self.create_ui()
    
    def create_ui(self):
        # 1. 文件选择区域
        frame_file = ttk.LabelFrame(self.root, text="文件选择", padding=(20, 10))
        frame_file.pack(fill="x", padx=20, pady=10)
        
        ttk.Label(frame_file, text="选择文件：").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(frame_file, textvariable=self.selected_file, width=40, state="readonly").grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(frame_file, text="浏览", command=self.browse_file).grid(row=0, column=2, padx=5, pady=5)
        
        # 2. 功能选择区域
        frame_function = ttk.LabelFrame(self.root, text="功能选择", padding=(20, 10))
        frame_function.pack(fill="x", padx=20, pady=10)
        
        ttk.Radiobutton(frame_function, text="客户文件读取", variable=self.function_var, value="客户文件读取").grid(row=0, column=0, padx=20, pady=5)
        ttk.Radiobutton(frame_function, text="转化为JS", variable=self.function_var, value="转化为JS").grid(row=0, column=1, padx=20, pady=5)
        
        # 3. 执行按钮区域
        frame_button = ttk.Frame(self.root, padding=(20, 10))
        frame_button.pack(fill="x", padx=20, pady=10)
        
        ttk.Button(frame_button, text="执行", width=20, command=self.execute).grid(row=0, column=0, padx=20, pady=5)
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.selected_file.set(file_path)
    
    def execute(self):
        file_path = self.selected_file.get()
        if not file_path:
            self.show_message("错误", "请先选择文件")
            return
        
        function = self.function_var.get()
        
        try:
            if function == "客户文件读取":
                self.process_customer_file(file_path)
            elif function == "转化为JS":
                self.convert_to_js(file_path)
        except Exception as e:
            self.show_message("错误", f"执行失败: {str(e)}")
    
    def process_customer_file(self, file_path):
        """处理客户文件"""
        try:
            # 获取所有sheet页名
            sheet_names = pd.ExcelFile(file_path).sheet_names
            exp_sheets = [name for name in sheet_names if name.startswith('EXP')]
            
            if not exp_sheets:
                self.show_message("提示", "未找到EXP开头的sheet")
                return
            
            # 处理每个EXP开头的sheet
            output_dir = "D:\\packingFile\\results"
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            for sheet_name in exp_sheets:
                # 读取当前sheet
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                
                # 找到 Total 所在的行
                total_row = None
                for i, row in df.iterrows():
                    for col in df.columns:
                        if 'Total' in str(row.get(col, '')) or 'total' in str(row.get(col, '')):
                            total_row = i
                            break
                    if total_row is not None:
                        break
                
                if total_row is None:
                    continue
                
                # 搜索包含"每箱"的单元格
                per_box_col = None
                for i, row in df.iterrows():
                    for j, val in enumerate(row):
                        if '每箱' in str(val):
                            per_box_col = j
                            break
                    if per_box_col is not None:
                        break
                
                if per_box_col is None:
                    continue
                
                # 提取数据
                size_base_cols = {}
                results = []
                row_num = 1
                size_row = df.iloc[10]
                col_first_qty_row = {}
                empty_row_idx = 0
                
                for i in range(11, total_row):
                    box_order = df.iloc[i, 1]
                    box_qty = df.iloc[i, 4]
                    start_col = 7
                    end_col = per_box_col
                    
                    valid_cells = []
                    for j in range(start_col, end_col):
                        qty = df.iloc[i, j]
                        if i > 11 and j == i - 5:
                            size_base_cols[str(df.iloc[10, j]).replace('/', ' ')] = i - 11
                        
                        if pd.notna(qty) and str(qty).strip():
                            box_order = df.iloc[i, 1]
                            valid_cells.append((j, qty))
                            if str(df.iloc[10, j]).replace('/', ' ') not in col_first_qty_row:
                                col_first_qty_row[str(df.iloc[10, j]).replace('/', ' ')] = i - 11
                    
                    cell_count = len(valid_cells)
                    for idx, (col_idx, qty) in enumerate(valid_cells):
                        if cell_count >= 2 and idx > 0:
                            current_box_qty = 0
                        else:
                            current_box_qty = box_qty
                        
                        size_value = str(size_row.iloc[col_idx]) if col_idx < len(size_row) else ""
                        size_value = size_value.replace('/', ' ')
                        
                        results.append([row_num, size_value, box_order, current_box_qty, qty])
                        row_num += 1
                
                # 处理空白尺码
                for size, row_idx in size_base_cols.items():
                    if size in col_first_qty_row and row_idx < col_first_qty_row[size]:
                        results.append([empty_row_idx, size, 0, 0, 0])
                        empty_row_idx += 1
                # 排序并重新编号
                results.sort(key=lambda x: (x[0], x[1]))
                for i in range(len(results)):
                    results[i][0] = i + 1
                
                # 保存结果
                file_name = os.path.basename(file_path)
                output_file_name = f"result_{os.path.splitext(file_name)[0]}_{sheet_name}.xlsx"
                output_file_path = os.path.join(output_dir, output_file_name)
                
                output_df = pd.DataFrame(results, columns=['行号', '尺码', '箱序', '箱量', '件数'])
                output_df.to_excel(output_file_path, index=False)
            
            self.show_message("成功", "客户文件处理完成")
        except Exception as e:
            raise e
    
    def convert_to_js(self, file_path):
        """转化为JS"""
        try:
            # 只读取第一个sheet页
            df = pd.read_excel(file_path, sheet_name=0)
            
            # 检查是否存在箱序、箱量、件数列
            required_columns = ['箱序', '箱量', '件数']
            for col in required_columns:
                if col not in df.columns:
                    self.show_message("错误", f"文件中缺少{col}列")
                    return
            
            # 处理每个EXP开头的sheet
            output_dir = "D:\\packingFile\\results"
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # 提取数据
            js_data = []
            for _, row in df.iterrows():
                box_order = row['箱序']
                box_qty = row['箱量']
                qty = row['件数']
                # 只添加有数据的行
                if pd.notna(box_order) and pd.notna(box_qty) and pd.notna(qty):
                    js_data.append([box_order, box_qty, qty])
            
            if not js_data:
                self.show_message("提示", "未找到有效的数据行")
                return
            
            # 生成JS文件
            file_name = os.path.basename(file_path)
            js_file_name = f"result_{os.path.splitext(file_name)[0]}.js"
            js_file_path = os.path.join(output_dir, js_file_name)
            
            js_content = f"// 每组 3 个数字：[ F66的值, F67的值, F70的值 ]\n"
            js_content += f"const data = [\n"
            for item in js_data:
                js_content += f"  [{item[0]}, {item[1]}, {item[2]}],\n"
            js_content = js_content.rstrip(',\n') + '\n';
            js_content += f"];\n\n"
            js_content += f"let totalRows = data.length;\n\n"
            js_content += f"for (let i = 0; i < totalRows; i++) {{\n"
            js_content += f"  // F66 = 第1个数字\n"
            js_content += f"  let f66 = document.getElementById(`F66_${{i}}`);\n"
            js_content += f"  if (f66) f66.value = data[i][0];\n\n"
            js_content += f"  // F67 = 第2个数字\n"
            js_content += f"  let f67 = document.getElementById(`F67_${{i}}`);\n"
            js_content += f"  if (f67) f67.value = data[i][1];\n\n"
            js_content += f"  // F70 = 第3个数字\n"
            js_content += f"  let f70 = document.getElementById(`F70_${{i}}`);\n"
            js_content += f"  if (f70) f70.value = data[i][2];\n"
            js_content += f"}}\n"
            
            with open(js_file_path, 'w', encoding='utf-8') as f:
                f.write(js_content)
            
            self.show_message("成功", "转化为JS完成")
        except Exception as e:
            raise e
    
    def show_message(self, title, message):
        """显示消息框"""
        messagebox.showinfo(title, message)

if __name__ == "__main__":
    root = tk.Tk()
    app = FileProcessorApp(root)
    root.mainloop()