import os
import pandas as pd

def process_excel_files(input_dir, output_dir):
    """
    处理 Excel 文件并输出结果
    
    Args:
        input_dir (str): 输入 Excel 文件目录
        output_dir (str): 输出结果目录
    """
    # 确保输出目录存在
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # 遍历输入目录中的所有 Excel 文件
    for file_name in os.listdir(input_dir):
        if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
            file_path = os.path.join(input_dir, file_name)
            print(f"处理文件: {file_name}")
            
            try:
                # 获取所有sheet页名
                sheet_names = pd.ExcelFile(file_path).sheet_names
                
                # 过滤出EXP开头的sheet
                exp_sheets = [name for name in sheet_names if name.startswith('EXP')]
                
                if not exp_sheets:
                    print(f"警告: 文件 {file_name} 中未找到EXP开头的sheet")
                    continue
                
                # 处理每个EXP开头的sheet
                for sheet_name in exp_sheets:
                    print(f"处理sheet: {sheet_name}")
                    
                    # 读取当前sheet
                    df = pd.read_excel(file_path, sheet_name=sheet_name)
                    
                    # 找到 Total 所在的行
                    total_row = None
                    for i, row in df.iterrows():
                        # 遍历所有列，查找包含 Total 的单元格
                        for col in df.columns:
                            if 'Total' in str(row.get(col, '')) or 'total' in str(row.get(col, '')):
                                total_row = i
                                break
                        if total_row is not None:
                            break
                    
                    if total_row is None:
                        print(f"警告: sheet {sheet_name} 中未找到 Total 行")
                        continue
                    
                    # 搜索整个Excel文件，找到包含"每箱"的单元格
                    per_box_col = None
                    for i, row in df.iterrows():
                        for j, val in enumerate(row):
                            if '每箱' in str(val):
                                per_box_col = j
                                print(f"找到'每箱'在第{i+1}行，第{j+1}列")
                                break
                        if per_box_col is not None:
                            break
                    
                    if per_box_col is None:
                        print(f"警告: sheet {sheet_name} 中未找到 每箱 列")
                        continue
                    
                    # 确保至少有足够的列
                    if len(df.columns) < 8:  # 至少需要H列（索引7）
                        print(f"警告: sheet {sheet_name} 列数不足")
                        continue
                    
                    # 提取数据
                    size_base_rows = {}
                    size_base_cols = {}
                    results = []
                    row_num = 1
                    
                    # 获取第11行（索引10）的内容，用于尺码列
                    size_row = df.iloc[10]
                    
                    # 记录每个尺码第一次出现数量的行
                    col_first_qty_row = {}

                    empty_row_idx = 0
                    
                    # 从 B12 开始到 Total 行
                    for i in range(11, total_row):  # 11 是因为 pandas 索引从 0 开始，B12 对应索引 11
                        # 获取 B 列内容（箱序）
                        box_order = df.iloc[i, 1]  # B 列对应索引 1
                        
                        # 获取 E 列内容（箱量）
                        box_qty = df.iloc[i, 4]  # E 列对应索引 4
                        
                        # 获取 H 列到"每箱"所在列的有数据的单元格
                        start_col = 7  # H 列对应索引 7
                        end_col = per_box_col  # "每箱"所在的列
                        
                        # 收集有数据的单元格
                        valid_cells = []
                        for j in range(start_col, end_col):
                            qty = df.iloc[i, j]
                            # 尺码应该有数据的列存储尺码数据，将/替换成空格
                            if i > 11 and j == i - 5:
                                size_base_rows[str(df.iloc[10, j]).replace('/', ' ')] = i - 11
                            if i == 11:
                                size_base_cols[str(df.iloc[10, j]).replace('/', ' ')] = j - 7

                            if pd.notna(qty) and str(qty).strip():
                                box_order = df.iloc[i, 1]
                                valid_cells.append((j, qty))
                                # 记录每个尺码第一次出现数量的行
                                if str(df.iloc[10, j]).replace('/', ' ') not in col_first_qty_row:
                                    col_first_qty_row[str(df.iloc[10, j]).replace('/', ' ')] = i - 11
                        
                        # 处理箱量显示逻辑
                        cell_count = len(valid_cells)
                        for idx, (col_idx, qty) in enumerate(valid_cells):
                            # 如果有数据的单元格超过2个，只有第一个显示正常箱量，其他显示0
                            if cell_count >= 2 and idx > 0:
                                current_box_qty = 0
                            else:
                                current_box_qty = box_qty
                            
                            # 获取尺码（第12行对应列的内容，将/替换成空格）
                            size_value = str(size_row.iloc[col_idx]) if col_idx < len(size_row) else ""
                            size_value = size_value.replace('/', ' ')
                            
                            results.append([i - 11, size_value, box_order, current_box_qty, qty, col_idx])
                            row_num += 1
                    # size_base_rows中的尺码对应的行号如果小于col_first_qty_row中尺码对应的行号，将尺码插入到results中，
                    # 尺码为对应的尺码，箱序为0，箱量为0，件数为0，行号为0（解决空白尺码的问题）
                    for idx, (size, row_idx) in enumerate(size_base_rows.items()):
                        if size in col_first_qty_row and row_idx < col_first_qty_row[size]:
                            results.append([empty_row_idx, size, 0, 0, 0, size_base_cols[size]])
                            empty_row_idx += 1
                    print(results)
                    # 按照行号和箱序升序排序
                    results.sort(key=lambda x: (x[0], x[1]))
                    # 重排results中的行号，从1开始
                    # for i in range(len(results)):
                    #     results[i][0] = i + 1
                    # 生成输出文件名
                    output_file_name = f"result_{os.path.splitext(file_name)[0]}_{sheet_name}.xlsx"
                    output_file_path = os.path.join(output_dir, output_file_name)
                    
                    # 生成输出 DataFrame
                    output_df = pd.DataFrame(results, columns=['行号', '尺码', '箱序', '箱量', '件数', '列索引'])
                    
                    # 保存为 Excel 文件
                    output_df.to_excel(output_file_path, index=False)
                    print(f"输出文件: {output_file_name}")
                    
                    # 生成 JavaScript 文件
                    js_file_name = f"result_{os.path.splitext(file_name)[0]}_{sheet_name}.js"
                    js_file_path = os.path.join(output_dir, js_file_name)
                    
                    # 提取数据用于 JavaScript 文件（只提取箱序、箱量、件数）
                    js_data = []
                    for _, row in output_df.iterrows():
                        js_data.append([row['箱序'], row['箱量'], row['件数']])
                    
                    # 生成 JavaScript 内容
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
                    
                    # 保存 JavaScript 文件
                    with open(js_file_path, 'w', encoding='utf-8') as f:
                        f.write(js_content)
                    print(f"输出文件: {js_file_name}")
                
            except Exception as e:
                print(f"处理文件 {file_name} 时出错: {str(e)}")

if __name__ == "__main__":
    input_directory = "D:\\packingFile"
    output_directory = "D:\\packingFile\\results"
    process_excel_files(input_directory, output_directory)
