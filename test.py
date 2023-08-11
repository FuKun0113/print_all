import win32print

def get_prints():
    prints_list = []
    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
    for printer_info in printers:
        _, _, printer_name, _= printer_info
        status = win32print.GetPrinter(win32print.OpenPrinter(printer_name), 2)['Status']
        prints_list.append({'printer_name': printer_name, 'status': status})
    return prints_list

