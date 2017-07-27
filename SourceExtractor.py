import tkinter as tk
from tkinter import filedialog, ttk
import os
import win32com.client
import shutil
import ctypes

def browse():
    root.withdraw()
    output_path = filedialog.askdirectory(title="Select Output Folder")
    entryText.set(output_path)
    browsed = output_path
    print(browsed)
    root.deiconify()
    return browsed


def run():
    shell = win32com.client.Dispatch("WScript.Shell")
    folder = os.listdir(os.getcwd())
    print(folder)
    print(entryText.get())
    destination = entryText.get()
    n = 0

    for i in folder:
        if ".lnk" in i:
            shortcut = shell.CreateShortCut(destination + "/" + folder[n])
            source = shortcut.Targetpath
            shutil.copy(source, destination)
            print("Copied file: ", shortcut.Targetpath)
        n += 1

    ctypes.windll.user32.MessageBoxW(0, "All files copied to current directory", "Complete", 1)


root = tk.Tk()
root.title("Source Extractor")
root.grid_columnconfigure(4, minsize=80)

entryText = tk.StringVar()

tk.Label(root, text="").grid(row=0, sticky='W', padx=4)
tk.Label(root, text="Select Output: ").grid(row=1, sticky='W', padx=4)
entry = tk.Entry(root, textvariable=entryText).grid(row=1, column=1, sticky='E', pady=4)

browseButton = tk.Button(root, text=" Browse ", command=browse).grid(row=1, column=3)
tk.Label(root, text="").grid(row=2, sticky='W', padx=4)
runButton = tk.Button(root, text=" Run ", command=run).grid(row=3, column=1)

tk.Label(root, text="").grid(row=4, sticky='W', padx=4)

root.mainloop()
