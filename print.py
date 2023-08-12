import win32print
from win32con import *

DC_CONSTANTS = [
    DC_BINNAMES, DC_BINS, DC_COLLATE, DC_COLORDEVICE, DC_COPIES, DC_DRIVER,
    DC_DUPLEX, DC_ENUMRESOLUTIONS, DC_EXTRA, DC_FIELDS,
    DC_FILEDEPENDENCIES, DC_MAXEXTENT, DC_MEDIAREADY, DC_MEDIATYPENAMES,
    DC_MEDIATYPES, DC_MINEXTENT, DC_ORIENTATION, DC_NUP, DC_PAPERNAMES,
    DC_PAPERS, DC_PAPERSIZE, DC_PERSONALITY, DC_PRINTERMEM, DC_PRINTRATE, DC_PRINTRATEPPM,
    DC_PRINTRATEUNIT, DC_SIZE, DC_STAPLE, DC_TRUETYPE, DC_VERSION,
]


def DC_INFO(constant):
    for a_global in globals().keys():
        if a_global.startswith("DC_") and globals().get(a_global) == constant:
            return a_global
    return "DC_UNKONWN"


for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS):
    print (printer)
    for constant in DC_CONSTANTS:
        try:
            x = win32print.DeviceCapabilities(printer[2], '', constant)
            print("\t", DC_INFO(constant), x)
        except:
            pass