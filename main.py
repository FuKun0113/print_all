import win32print
printers = win32print.EnumPrinters(3)
print(printers)
printer = win32print.GetDefaultPrinter()
print(printer)