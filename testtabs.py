import tkinter as tk                    
from tkinter import ttk, filedialog, scrolledtext, font
from tkinter.messagebox import showinfo, showwarning
from tkinter import *
from screeninfo import get_monitors
from tkcalendar import DateEntry
from tkinter.messagebox import showinfo, showwarning
import json, subprocess, os, threading, base64, sys
from datetime import datetime
from tkinter.ttk import *
from UI_scriptFunctions import select_folder,begin_order_processing,open_folder, open_folder_packaging, select_files
from config import ConfigFolderPath, headingFont,fieldFont,buttonFont,labelFont,pathFont,logFont,ClientsFolderPath
from UI_tabs import Tab1, Tab2
from UI_logscmd import TextHandler


# Getting Screen Dimensions
DeviceScreenHeight = ''
DeviceScreenWidth = ''
for m in get_monitors():
   DeviceScreenWidth =str(m.width)
   DeviceScreenHeight = str(m.height)


root = Tk()
root.geometry(DeviceScreenWidth+"x"+DeviceScreenHeight)
root.minsize(int(DeviceScreenWidth)-300,int(DeviceScreenHeight)-250)
root.maxsize(int(DeviceScreenWidth),int(DeviceScreenHeight)-50)

# root.iconbitmap('C:/Users/HP/Documents/GitHub/Triumph.ico')

root.title("Purchase Orders")

with open(ConfigFolderPath+'config.json', 'r') as jsonFile:
  config = json.load(jsonFile)
  themepath = config['appTheme']

root.tk.call("source", themepath)
root.tk.call("set_theme", "dark")


s = ttk.Style()
s.configure('TNotebook.Tab', font=('Calibri','15'), padding=[100, 10])

tabFrame = ttk.Frame(root)
logFrame = ttk.Frame(root)

tabControl = ttk.Notebook(tabFrame)

filenames = tk.StringVar() 
selectedDate = tk.StringVar() 
# TAB 1
tab1 = Tab1(root,tabControl)
# TAB 2
tab2 = Tab2(root,tabControl)
tabControl.pack(fill ="x")



tabFrame.pack(side='top',anchor=NW,fill ="x")
logFrame.pack(side='bottom',anchor=SW,fill ="x")
print('UI Loaded...')



root.mainloop()  