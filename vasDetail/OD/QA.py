import json
from collections import defaultdict

# 读取 JSON 文件
with open('D:\\pythonDemo\\vasDetail\\OD\\dist\\todolist\\all_tasks.json', 'r', encoding='utf-8') as f:
    tasks = json.load(f)

# 初始化一个嵌套的字典来存储按日期和部门分类的任务数量
date_department_task_count = defaultdict(lambda: defaultdict(int))

# 遍历每个任务
for task in tasks:
    name = task['name']
    created_date = task['created_date'].split('T')[0]
    # 提取部门信息
    if '1部' in name:
        department = '1部'
    elif '2部' in name:
        department = '2部'
    elif '3部' in name:
        department = '3部'
    elif '4部' in name:
        department = '4部'
    elif '5部' in name:
        department = '5部'
    elif '6部' in name:
        department = '6部'
    elif '财务部' in name:
        department = '财务部'
    elif '采购部' in name:
        department = '采购部'
    elif '船务部' in name:
        department = '船务部'
    elif '生产二部' in name:
        department = '生产二部'
    else:
        department = '其他'

    # 统计每个日期下各部门的任务数量
    date_department_task_count[created_date][department] += 1

# 转换数据为二维表格格式
table_data = []
for date, departments in date_department_task_count.items():
    for department, count in departments.items():
        table_data.append([date, department, count])

# 按日期和部门排序
table_data.sort(key=lambda x: (x[0], x[1]))

# 输出表头
print(f"{'日期':<12}{'部门':<10}{'问题数量':<10}")
print("-" * 32)

# 输出表格数据
for row in table_data:
    print(f"{row[0]:<12}{row[1]:<10}{row[2]:<10}")