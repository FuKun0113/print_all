import win32print

# 获取系统中的打印机及状态，输出一个list
def get_prints():
    # 创建一个list
    prints_list = []
    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
    for printer_info in printers:
        # 获取打印机名称
        _, _, printer_name, _= printer_info
        # 获取状态
        status = win32print.GetPrinter(win32print.OpenPrinter(printer_name), 2)['Status']
        # 将打印机名称及状态，以dict的形式存进list
        prints_list.append({'printer_name': printer_name, 'status': status})
    return prints_list


printer = win32print.GetDefaultPrinter()

print(win32print.GetPrinter(win32print.OpenPrinter(printer), 2))