import os
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import win32print
import win32ui


class PrintApp:
    def __init__(self, root):
        self.root = root
        self.root.title("文件批量打印应用程序")

        self.directory_label = tk.Label(root, text="选择目标文件夹:")
        self.directory_label.pack()

        self.select_directory_button = tk.Button(root, text="选择文件夹", command=self.select_directory)
        self.select_directory_button.pack()

        self.files_listbox = tk.Listbox(root, selectmode=tk.MULTIPLE)
        self.files_listbox.pack()

        self.printer_label = tk.Label(root, text="选择打印机:")
        self.printer_label.pack()

        self.printer_combobox = ttk.Combobox(root, values=self.get_printer_list())
        self.printer_combobox.pack()

        self.double_sided_var = tk.IntVar()
        self.double_sided_checkbox = tk.Checkbutton(root, text="双面打印", variable=self.double_sided_var)
        self.double_sided_checkbox.pack()

        self.copies_label = tk.Label(root, text="输入打印份数:")
        self.copies_label.pack()

        self.copies_entry = tk.Entry(root)
        self.copies_entry.pack()

        self.print_button = tk.Button(root, text="开始打印", command=self.start_printing)
        self.print_button.pack()

        self.file_settings = {}  # 用于存储每个文件的打印设置

    def select_directory(self):
        self.directory = filedialog.askdirectory()
        self.update_files_list()

    def update_files_list(self):
        self.files_listbox.delete(0, tk.END)
        files = os.listdir(self.directory)
        for file in files:
            self.files_listbox.insert(tk.END, file)

    def get_printer_list(self):
        printers = []
        printer_info = win32print.EnumPrinters(3)  # 获取打印机列表
        for p in printer_info:
            printers.append(p[2])
        return printers

    def start_printing(self):
        printer_name = self.printer_combobox.get()
        num_copies = int(self.copies_entry.get())
        double_sided = self.double_sided_var.get() == 1

        selected_files = [self.files_listbox.get(i) for i in self.files_listbox.curselection()]

        for file in selected_files:
            print_file = os.path.join(self.directory, file)
            file_settings = {
                "printer_name": printer_name,
                "num_copies": num_copies,
                "double_sided": double_sided
            }
            self.file_settings[file] = file_settings
            self.print_file(print_file, file_settings)

    def print_file(self, file, settings):
        printer_name = settings["printer_name"]
        num_copies = settings["num_copies"]
        double_sided = settings["double_sided"]

        # 实际的打印操作代码


if __name__ == "__main__":
    root = tk.Tk()
    app = PrintApp(root)
    root.mainloop()
