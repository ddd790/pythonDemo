import tkinter as tk
from tkinter import filedialog, messagebox
import shutil
import os
from datetime import datetime


class FileCopyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("文件另存工具")
        self.root.geometry("500x400")
        self.root.resizable(False, False)
        
        self.selected_file = ""
        self.output_folder = ""
        self.software_type = tk.StringVar(value="word")
        
        self.create_widgets()
    
    def create_widgets(self):
        frame1 = tk.LabelFrame(self.root, text="选择模板", padx=10, pady=10)
        frame1.pack(fill="x", padx=10, pady=10)
        
        self.file_entry = tk.Entry(frame1, width=50)
        self.file_entry.pack(side="left", padx=(0, 10))
        
        btn_select_file = tk.Button(frame1, text="选择", command=self.select_file)
        btn_select_file.pack(side="left")
        
        frame2 = tk.LabelFrame(self.root, text="软件选择", padx=10, pady=10)
        frame2.pack(fill="x", padx=10, pady=5)
        
        rb_word = tk.Radiobutton(frame2, text="word", variable=self.software_type, value="word")
        rb_word.pack(side="left", padx=20)
        
        rb_excel = tk.Radiobutton(frame2, text="excel", variable=self.software_type, value="excel")
        rb_excel.pack(side="left", padx=20)
        
        rb_ppt = tk.Radiobutton(frame2, text="ppt", variable=self.software_type, value="ppt")
        rb_ppt.pack(side="left", padx=20)
        
        frame3 = tk.LabelFrame(self.root, text="输出位置", padx=10, pady=10)
        frame3.pack(fill="x", padx=10, pady=10)
        
        self.folder_entry = tk.Entry(frame3, width=50)
        self.folder_entry.pack(side="left", padx=(0, 10))
        
        btn_select_folder = tk.Button(frame3, text="选择", command=self.select_folder)
        btn_select_folder.pack(side="left")
        
        btn_save = tk.Button(self.root, text="保存", width=20, height=2, command=self.save_file)
        btn_save.pack(pady=20)
    
    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="选择模板文件",
            filetypes=[
                ("Word文件", "*.docx *.doc"),
                ("Excel文件", "*.xlsx *.xls"),
                ("PPT文件", "*.pptx *.ppt"),
                ("所有文件", "*.*")
            ]
        )
        if file_path:
            self.selected_file = file_path
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)
    
    def select_folder(self):
        folder_path = filedialog.askdirectory(title="选择输出文件夹")
        if folder_path:
            self.output_folder = folder_path
            self.folder_entry.delete(0, tk.END)
            self.folder_entry.insert(0, folder_path)
    
    def save_file(self):
        if not self.selected_file:
            messagebox.showwarning("警告", "请先选择模板文件！")
            return
        
        if not self.output_folder:
            messagebox.showwarning("警告", "请先选择输出位置！")
            return
        
        try:
            original_name = os.path.basename(self.selected_file)
            name, ext = os.path.splitext(original_name)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            new_name = f"{name}_{timestamp}{ext}"
            output_path = os.path.join(self.output_folder, new_name)
            
            shutil.copy2(self.selected_file, output_path)
            os.startfile(output_path)
            messagebox.showinfo("成功", f"文件已保存到：\n{output_path}")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败：{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = FileCopyApp(root)
    root.mainloop()
