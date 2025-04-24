import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, date
import json
import os
import threading
import time

class Task:
    def __init__(self, name, reminder_time=None):
        self.name = name
        self.completed = False
        self.reminder_time = reminder_time
        self.created_date = datetime.now()
        self.is_reminded = False

    def to_dict(self):
        return {
            'name': self.name,
            'completed': self.completed,
            'reminder_time': self.reminder_time.isoformat() if self.reminder_time else None,
            'created_date': self.created_date.isoformat(),
            'is_reminded': self.is_reminded
        }

    @classmethod
    def from_dict(cls, data):
        task = cls(data['name'])
        task.completed = data['completed']
        task.reminder_time = datetime.fromisoformat(data['reminder_time']) if data['reminder_time'] else None
        task.created_date = datetime.fromisoformat(data['created_date'])
        task.is_reminded = data['is_reminded']
        return task

class TodoApp:
    def __init__(self, root):
        self.root = root
        self.tasks = []
        self.filename = "all_tasks.json"
        self.sort_column = None
        self.sort_ascending = True
        self.setup_ui()
        self.load_data()
        self.setup_reminder_check()

    def setup_ui(self):
        self.root.title("待办事项管理")
        self.root.geometry("1000x600")

        # 顶部工具栏
        toolbar = ttk.Frame(self.root)
        toolbar.pack(fill=tk.X, padx=5, pady=5)

        self.add_btn = ttk.Button(toolbar, text="添加任务", command=self.show_edit_dialog)
        self.add_btn.pack(side=tk.LEFT, padx=2)

        self.edit_btn = ttk.Button(toolbar, text="编辑任务", command=self.edit_task)
        self.edit_btn.pack(side=tk.LEFT, padx=2)

        self.del_btn = ttk.Button(toolbar, text="删除任务", command=self.delete_task)
        self.del_btn.pack(side=tk.LEFT, padx=2)

        self.complete_btn = ttk.Button(toolbar, text="标记/取消完成", command=self.toggle_complete)
        self.complete_btn.pack(side=tk.LEFT, padx=2)

        # 任务列表
        self.tree = ttk.Treeview(self.root, columns=('name', 'status', 'created', 'reminder'), show='headings')
        self.tree.heading('name', text='任务名称', command=lambda: self.sort_by_column('name'))
        self.tree.heading('status', text='完成状态', command=lambda: self.sort_by_column('status'))
        self.tree.heading('created', text='创建时间', command=lambda: self.sort_by_column('created'))
        self.tree.heading('reminder', text='提醒时间', command=lambda: self.sort_by_column('reminder'))
        
        self.tree.column('name', width=300)
        self.tree.column('status', width=100)
        self.tree.column('created', width=200)
        self.tree.column('reminder', width=200)
        
        self.tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 配置标签样式
        self.tree.tag_configure('completed', background='#f0f0f0', foreground='green', font=('微软雅黑', 10, 'overstrike'))
        self.tree.tag_configure('pending', background='white')

        # 状态栏
        self.status = ttk.Label(self.root, text="就绪", relief=tk.SUNKEN)
        self.status.pack(side=tk.BOTTOM, fill=tk.X)

        # 绑定双击事件
        self.tree.bind("<Double-1>", self.edit_task)

    def sort_by_column(self, column):
        if self.sort_column == column:
            self.sort_ascending = not self.sort_ascending
        else:
            self.sort_column = column
            self.sort_ascending = True
        self.update_list()

    def show_edit_dialog(self, task=None):
        dialog = tk.Toplevel(self.root)
        dialog.title("编辑任务" if task else "添加新任务")
        
        # 表单字段
        fields = {
            'name': {'label': '任务名称:', 'value': task.name if task else ''},
            'reminder': {'label': '提醒时间:', 'value': task.reminder_time.strftime("%Y-%m-%d %H:%M") if task and task.reminder_time else ''}
        }

        entries = {}
        for i, (key, config) in enumerate(fields.items()):
            ttk.Label(dialog, text=config['label']).grid(row=i, column=0, padx=5, pady=5)
            entry = ttk.Entry(dialog, width=30)
            entry.insert(0, config['value'])
            entry.grid(row=i, column=1, padx=5, pady=5)
            entries[key] = entry

            # 为名称输入框添加自动聚焦
            if key == 'name':
                entry.focus_set()
                if task:  # 编辑模式时移动光标到末尾
                    entry.icursor(tk.END)
                # 绑定回车键到第一个输入框
                entry.bind('<Return>', lambda e: save_task())

        def save_task():
            name = entries['name'].get()
            if not name:
                messagebox.showerror("错误", "任务名称不能为空")
                return
            
            reminder_str = entries['reminder'].get()
            reminder_time = None
            
            if reminder_str.strip():
                try:
                    reminder_time = datetime.strptime(reminder_str, "%Y-%m-%d %H:%M")
                    if reminder_time < datetime.now():
                        messagebox.showerror("错误", "提醒时间不能早于当前时间")
                        return
                except ValueError:
                    messagebox.showerror("错误", "时间格式应为 YYYY-MM-DD HH:MM")
                    return

            if task:
                # 更新现有任务
                task.name = name
                task.reminder_time = reminder_time
            else:
                # 添加新任务
                self.tasks.append(Task(name, reminder_time))
            
            self.update_list()
            self.save_data()
            dialog.destroy()

        # 绑定回车键到保存操作
        dialog.bind('<Return>', lambda event: save_task())
        ttk.Button(dialog, text="保存 (Enter)", command=save_task).grid(row=len(fields), column=1, pady=10)

    def update_list(self):
        if self.sort_column:
            key_map = {
                'name': lambda t: t.name.lower(),
                'status': lambda t: t.completed,
                'created': lambda t: t.created_date,
                'reminder': lambda t: t.reminder_time if t.reminder_time else datetime.max
            }
            self.tasks.sort(
                key=key_map[self.sort_column],
                reverse=not self.sort_ascending
            )

        self.tree.delete(*self.tree.get_children())
        for task in self.tasks:
            status = "已完成" if task.completed else "未完"
            created = task.created_date.strftime("%Y-%m-%d %H:%M")
            reminder = task.reminder_time.strftime("%Y-%m-%d %H:%M") if task.reminder_time else "无"
            self.tree.insert('', 'end', values=(
                task.name,
                status,
                created,
                reminder
            ), tags=('completed' if task.completed else 'pending'))

    def edit_task(self, event=None):
        selected = self.tree.selection()
        if not selected:
            return
        index = self.tree.index(selected[0])
        task = self.tasks[index]
        self.show_edit_dialog(task)

    def delete_task(self):
        selected = self.tree.selection()
        if not selected:
            return
        index = self.tree.index(selected[0])
        del self.tasks[index]
        self.update_list()
        self.save_data()

    def toggle_complete(self):
        selected = self.tree.selection()
        if not selected:
            return
        index = self.tree.index(selected[0])
        self.tasks[index].completed = not self.tasks[index].completed
        self.update_list()
        self.save_data()

    def show_detail(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        index = self.tree.index(selected[0])
        task = self.tasks[index]
        
        detail = f"""任务详情：
名称：{task.name}
状态：{'已完成' if task.completed else '未完成'}
创建日期：{task.created_date}
提醒时间：{task.reminder_time.strftime('%Y-%m-%d %H:%M') if task.reminder_time else '无'}"""
        
        messagebox.showinfo("任务详情", detail)

    def setup_reminder_check(self):
        def check_reminders():
            while True:
                now = datetime.now()
                for task in self.tasks:
                    if (task.reminder_time and 
                        not task.is_reminded and 
                        now >= task.reminder_time):
                        self.root.after(0, lambda t=task: self.show_reminder(t))
                        task.is_reminded = True
                time.sleep(60)
        
        thread = threading.Thread(target=check_reminders, daemon=True)
        thread.start()

    def show_reminder(self, task):
        self.root.wm_attributes("-topmost", 1)
        messagebox.showwarning("提醒", f"任务 '{task.name}' 时间到了！")
        self.root.wm_attributes("-topmost", 0)

    def save_data(self):
        data = [task.to_dict() for task in self.tasks]
        with open(self.filename, 'w') as f:
            json.dump(data, f, indent=4)
        self.status.config(text="数据已自动保存")

    def load_data(self):
        if os.path.exists(self.filename):
            try:
                with open(self.filename, 'r') as f:
                    data = json.load(f)
                    self.tasks = [Task.from_dict(item) for item in data]
                self.update_list()
                self.status.config(text="数据已加载")
            except Exception as e:
                messagebox.showerror("错误", f"加载数据失败: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = TodoApp(root)
    def on_closing():
        app.save_data()
        root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()