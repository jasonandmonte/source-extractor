import os
import win32com.client
import shutil
import ctypes


shell = win32com.client.Dispatch("WScript.Shell")
folder = os.listdir(os.getcwd())
destination = os.getcwd()
n = 0

for i in folder:
    if ".lnk" in i:
        shortcut = shell.CreateShortCut(destination + "/" + folder[n])
        source = shortcut.Targetpath
        shutil.copy(source, destination)
        print("Copied file: ", shortcut.Targetpath)
    n += 1

ctypes.windll.user32.MessageBoxW(0, "All files copied to current directory", "Complete", 1)
