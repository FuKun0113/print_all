# import win32print
# printers = win32print.EnumPrinters(3)
# print(printers)
# printer = win32print.GetDefaultPrinter()
# print(printer)

# 使用custom tkinter创建程序GUI
# 使用pywin32库使用打印功能
# 安装依赖
# pip install customtkinter
# pip install pywin32

# 引入模块
import tkinter
import customtkinter
import win32print
import win32api
import os

# 功能部分
# ---------------------------------------------------------------------------------------------------------------------------------------
# 获取所有打印机
printers = []
for it in win32print.EnumPrinters(3):
    printers.append(it[2])
print(printers)
# 获取默认打印机
printer = win32print.GetDefaultPrinter()
print(printer)

# GUI部分
# ---------------------------------------------------------------------------------------------------------------------------------------
# 主题风格: "System" (standard), "Dark", "Light"
customtkinter.set_appearance_mode("System")
# 主题颜色: "blue" (standard), "green", "dark-blue"
customtkinter.set_default_color_theme("blue")


# 创建主程序
class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        # 创建窗口
        self.title("云雀批量打印工具")
        self.geometry(f"{800}x{700}")
        self.minsize(800, 700)

        # 自适应设置
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # 创建说明
        self.read_me = customtkinter.CTkLabel(master=self, text=r"本工具支持excel、pdf、word、文本、图片等文件的批量打印")
        self.read_me.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # 工具栏
        self.tool_bar = customtkinter.CTkFrame(master=self)
        self.tool_bar.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        # 选择打印机
        self.sel_printer = customtkinter.CTkOptionMenu(master=self.tool_bar, values=printers)
        self.sel_printer.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.sel_printer.set(printer)
        # 选择文件夹
        self.sel_path = customtkinter.CTkEntry(master=self.tool_bar, placeholder_text="请选择文件夹")
        self.sel_path.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        # 选择文件夹按钮
        self.sel_file = customtkinter.CTkButton(master=self.tool_bar, text="📂", width=50, command=self.file_path)
        self.sel_file.grid(row=0, column=2, padx=10, pady=10, sticky="nsew")
        # 读取文件列表按钮
        self.read_list = customtkinter.CTkButton(master=self.tool_bar, text="读取列表", command=self.load_list)
        self.read_list.grid(row=0, column=3, padx=10, pady=10, sticky="nsew")

        # 文件列表区
        self.list_area = customtkinter.CTkScrollableFrame(master=self)
        self.list_area.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")
        # 文件名
        self.file_name = customtkinter.CTkLabel(master=self.list_area, text="文件名")
        self.file_name.grid(row=0, column=0, padx=10)
        # 文件格式
        self.ex_name = customtkinter.CTkLabel(master=self.list_area, text="文件格式")
        self.ex_name.grid(row=0, column=1, padx=10)
        # 单面双面
        self.label2 = customtkinter.CTkLabel(master=self.list_area, text="单面双面")
        self.label2.grid(row=0, column=2, padx=10)
        # 打印份数
        self.label3 = customtkinter.CTkLabel(master=self.list_area, text="打印份数")
        self.label3.grid(row=0, column=3, padx=10)

        # 按钮区
        self.print_btn = customtkinter.CTkButton(master=self, text="打印")
        self.print_btn.grid(row=3, column=0, padx=10, pady=10, sticky="ns")

    # 函数区
    # ---------------------------------------------------------------------------------------------------------------------------------------
    # 选择文件夹函数
    def file_path(self):
        dir_path = tkinter.filedialog.askdirectory()
        self.sel_path.delete(0, "end")
        self.sel_path.insert(0, dir_path)

    # 读取文件列表函数
    def load_list(self):
        path = self.sel_path.get()
        files = os.listdir(path)  # 得到文件夹下的所有文件名称
        i = 1
        for file in files:  # 遍历文件夹
            self.file_title = customtkinter.CTkLabel(master=self.list_area,text=file)
            self.file_title.grid(row=i, column=0)
            i += 1


# 运行程序
if __name__ == "__main__":
    app = App()
    app.mainloop()
