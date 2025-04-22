import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from datetime import datetime, date, timedelta
import json
import os
import threading
import time

class Task:
    def __init__(self, name, category, reminder_time=None):
        self.name = name
        self.category = category
        self.completed = False
        self.reminder_time = reminder_time
        self.created_date = date.today()
        self.is_reminded = False

    def to_dict(self):
        return {
            'name': self.name,
            'category': self.category,
            'completed': self.completed,
            'reminder_time': self.reminder_time.isoformat() if self.reminder_time else None,
            'created_date': self.created_date.isoformat(),
            'is_reminded': self.is_reminded
        }

    @classmethod
    def from_dict(cls, data):
        task = cls(data['name'], data['category'])
        task.completed = data['completed']
        task.reminder_time = datetime.fromisoformat(data['reminder_time']) if data['reminder_time'] else None
        task.created_date = datetime.fromisoformat(data['created_date']).date()
        task.is_reminded = data['is_reminded']
        return task

class TodoApp:
    def __init__(self, root):
        self.root = root
        self.tasks = []
        self.filename = f"{date.today().isoformat()}.json"
        self.sort_column = None
        self.sort_ascending = True
        self.setup_ui()
        self.load_data()
        self.setup_reminder_check()

    def setup_ui(self):
        self.root.title("智能待办事项管理")
        self.root.geometry("800x500")

        # 顶部工具栏
        toolbar = ttk.Frame(self.root)
        toolbar.pack(fill=tk.X, padx=5, pady=5)

        self.add_btn = ttk.Button(toolbar, text="添加任务", command=self.show_add_dialog)
        self.add_btn.pack(side=tk.LEFT, padx=2)

        self.del_btn = ttk.Button(toolbar, text="删除任务", command=self.delete_task)
        self.del_btn.pack(side=tk.LEFT, padx=2)

        self.complete_btn = ttk.Button(toolbar, text="标记完成", command=self.toggle_complete)
        self.complete_btn.pack(side=tk.LEFT, padx=2)

        # 任务列表
        self.tree = ttk.Treeview(self.root, columns=('name', 'category', 'reminder'), show='headings')
        self.tree.heading('name', text='任务名称', command=lambda: self.sort_by_column('name'))
        self.tree.heading('category', text='分类', command=lambda: self.sort_by_column('category'))
        self.tree.heading('reminder', text='提醒时间', command=lambda: self.sort_by_column('reminder'))
        self.tree.column('name', width=300)
        self.tree.column('category', width=150)
        self.tree.column('reminder', width=200)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 状态栏
        self.status = ttk.Label(self.root, text="就绪", relief=tk.SUNKEN)
        self.status.pack(side=tk.BOTTOM, fill=tk.X)

        # 绑定双击事件
        self.tree.bind("<Double-1>", self.show_detail)

    def sort_by_column(self, column):
        """点击列头排序功能"""
        if self.sort_column == column:
            self.sort_ascending = not self.sort_ascending
        else:
            self.sort_column = column
            self.sort_ascending = True
        self.update_list()

    def show_add_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("添加新任务")
        
        ttk.Label(dialog, text="任务名称:").grid(row=0, column=0, padx=5, pady=5)
        name_entry = ttk.Entry(dialog, width=30)
        name_entry.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(dialog, text="分类:").grid(row=1, column=0, padx=5, pady=5)
        category_entry = ttk.Entry(dialog, width=30)
        category_entry.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(dialog, text="提醒时间:").grid(row=2, column=0, padx=5, pady=5)
        reminder_entry = ttk.Entry(dialog, width=30)
        reminder_entry.grid(row=2, column=1, padx=5, pady=5)
        reminder_entry.insert(0, datetime.now().strftime("%Y-%m-%d %H:%M"))

        def add_task():
            name = name_entry.get()
            category = category_entry.get()
            reminder_str = reminder_entry.get()
            
            try:
                reminder_time = datetime.strptime(reminder_str, "%Y-%m-%d %H:%M")
                if reminder_time < datetime.now():
                    messagebox.showerror("错误", "提醒时间不能早于当前时间")
                    return
            except ValueError:
                reminder_time = None

            self.tasks.append(Task(name, category, reminder_time))
            self.update_list()
            self.save_data()
            dialog.destroy()

        ttk.Button(dialog, text="添加", command=add_task).grid(row=3, column=1, pady=10)

    def update_list(self):
        """更新列表并应用过滤和排序"""
        # 过滤已完成的任务
        filtered_tasks = [t for t in self.tasks if not t.completed]
        
        # 执行排序
        if self.sort_column:
            key_map = {
                'name': lambda t: t.name.lower(),
                'category': lambda t: t.category.lower(),
                'reminder': lambda t: t.reminder_time if t.reminder_time else datetime.max
            }
            filtered_tasks.sort(
                key=key_map[self.sort_column],
                reverse=not self.sort_ascending
            )

        self.tree.delete(*self.tree.get_children())
        for task in filtered_tasks:
            reminder = task.reminder_time.strftime("%Y-%m-%d %H:%M") if task.reminder_time else "无"
            self.tree.insert('', 'end', values=(
                task.name,
                task.category,
                reminder
            ), tags=('completed' if task.completed else 'pending'))

    def delete_task(self):
        selected = self.tree.selection()
        if not selected:
            return
        index = self.tree.index(selected[0])
        # 通过显示列表反查原始数据索引
        filtered_tasks = [t for t in self.tasks if not t.completed]
        original_index = self.tasks.index(filtered_tasks[index])
        del self.tasks[original_index]
        self.update_list()
        self.save_data()

    def toggle_complete(self):
        selected = self.tree.selection()
        if not selected:
            return
        index = self.tree.index(selected[0])
        # 通过显示列表反查原始数据索引
        filtered_tasks = [t for t in self.tasks if not t.completed]
        original_index = self.tasks.index(filtered_tasks[index])
        self.tasks[original_index].completed = True
        self.update_list()
        self.save_data()

    def show_detail(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        index = self.tree.index(selected[0])
        # 通过显示列表反查原始数据索引
        filtered_tasks = [t for t in self.tasks if not t.completed]
        task = filtered_tasks[index]
        
        detail = f"""任务详情：
名称：{task.name}
分类：{task.category}
状态：未完成
创建时间：{task.created_date}
提醒时间：{task.reminder_time or '无'}"""
        
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
            with open(self.filename, 'r') as f:
                data = json.load(f)
                self.tasks = [Task.from_dict(item) for item in data]
            self.update_list()
            self.status.config(text="数据已加载")

if __name__ == "__main__":
    root = tk.Tk()
    app = TodoApp(root)
    
    def on_closing():
        app.save_data()
        root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()