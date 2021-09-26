# -- coding: utf-8 --
# @Time : 2021/9/23 16:13
# @Author : Shi Rui

import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import os
from scanner import Scanner
import traceback


class MainBox:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("属性导出工具")
        self.window.geometry("250x350")
        tk.Label(self.window, text="schema文件夹").pack()
        self.schema_entry = tk.Entry(self.window, width='30')
        self.schema_entry.pack()
        tk.Button(self.window, text="选取schema文件夹", command=self.select_schema_directory).pack()
        tk.Label(self.window).pack()
        tk.Label(self.window, text="材料属性数据库").pack()
        self.code_entry = tk.Entry(self.window, width='30')
        self.code_entry.pack()
        tk.Button(self.window, text="选取材料db文件", command=self.select_code_file).pack()
        tk.Label(self.window).pack()
        tk.Label(self.window, text="项目sqlite数据库").pack()
        self.db_entry = tk.Entry(self.window, width='30')
        self.db_entry.pack()
        tk.Button(self.window, text="选取sqlite文件", command=self.select_sqlite_file).pack()
        tk.Label(self.window).pack()
        tk.Button(self.window, text="导出", command=self.export).pack()
        self.window.mainloop()

    def select_sqlite_file(self):
        self.db_entry.delete(0, 'end')
        file_path = filedialog.askopenfilename()
        self.db_entry.insert(0, file_path)

    def select_code_file(self):
        self.code_entry.delete(0, 'end')
        file_path = filedialog.askopenfilename()
        self.code_entry.insert(0, file_path)

    def select_schema_directory(self):
        self.schema_entry.delete(0, 'end')
        file_path = filedialog.askdirectory()
        self.schema_entry.insert(0, file_path)

    def export(self):
        schema_path = self.schema_entry.get()
        db_path = self.db_entry.get()
        code_path = self.code_entry.get()
        schema_path = './data/schema'
        code_path = './data/ATCDI_tp_20190101.db'
        db_path = './data/00-00-all.db'
        if os.path.isdir(schema_path) and os.path.isfile(db_path) and os.path.isfile(code_path):
            try:
                Scanner(db_path, schema_path, code_path).start()
            except Exception:
                messagebox.showerror('错误', traceback.format_exc())

        else:
            messagebox.showerror("错误", '请检查文件路径输入是否正确')


if __name__ == '__main__':
    main_box = MainBox()
