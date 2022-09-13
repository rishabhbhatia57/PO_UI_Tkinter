import tkinter as tk                    
from tkinter import ttk
from tkinter import *
from screeninfo import get_monitors
from tkcalendar import DateEntry
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
import tkinter.scrolledtext as st
import os
import threading
import base64
import tkinter.font as tkFont
from tab2 import Tab2
from tab1 import Tab1
from tab3 import Tab3
import json

# Getting Screen Dimensions
DeviceScreenHeight = ''
DeviceScreenWidth = ''
for m in get_monitors():
   DeviceScreenWidth =str(m.width)
   DeviceScreenHeight = str(m.height)

# Variables declaration
POFolderSelected = ''
OrderDateSelected = ''
ClientCodeSelected = ''
ClientCode = {
  "Pantaloons": "PL",
  "Shoppers Stop Limited": "SSL",
  "Lifestyle Limited": "LSL"
}

root = Tk()
root.geometry(DeviceScreenWidth+"x"+DeviceScreenHeight)
root.minsize(1700,1000)
root.maxsize(int(DeviceScreenWidth),int(DeviceScreenHeight))

root.title("Purchase Orders")

with open('C:/Users/HP/Desktop/PO Metadata/Configfiles-Folder/config.json', 'r') as jsonFile:
  config = json.load(jsonFile)
  themepath = config['appTheme']

root.tk.call("source", themepath)
root.tk.call("set_theme", "dark")


s = ttk.Style()
s.configure('TNotebook.Tab', font=('Calibri','15'), padding=[100, 10])
tabControl = ttk.Notebook(root)

tab1 = ttk.Frame(tabControl)

filenames = tk.StringVar() 
selectedDate = tk.StringVar() 

headingFont = ("Calibri",40,"bold")
fieldFont=  ("Calibri",20,"bold")
# TAB 1
tab1 = Tab1(root,tabControl)
# TAB 2
tab2 = Tab2(root,tabControl)
# TAB 3
tab3 = Tab3(root,tabControl)

tabControl.pack(expand = 1, fill ="both")

root.mainloop()  