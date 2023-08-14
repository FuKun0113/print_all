import os
import sys
import ctypes

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if is_admin():
    # Code of your main application
    os.system('python main_script.py')  # replace 'main_script.py' with the name of your main script
else:
    # Re-run the program with admin rights, might trigger UAC prompt
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)