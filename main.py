import os
import sys
import subprocess

# def install(package):
#     subprocess.check_call([sys.executable, "-m", "pip", "install", package])
# install("docx")

# from docx import Document
from tkinter import *
from tkinter import ttk

def filter(files):
    result =[]
    for f in files:
        if ".docx" in f:
            result.append(f)

def selected(event):
    selection = combobox.get()
    document = Document(f'docs/{selection}')

root = Tk()
root.title("StudUtilities")
root.geometry("700x300")

dir_content = os.listdir('docs')
docs_list = filter(dir_content)

label = ttk.Label()
label.pack(anchor=NW, fill=X, padx=5, pady=5)



combobox = ttk.Combobox(root, values=docs_list)
combobox.pack(anchor=NW, fill=X, 	padx=6, pady=6)
combobox.bind("<<ComboboxSelected>>", selected)

root.mainloop()
