import tkinter as tk                    
from tkinter import ttk
from tkinter.ttk import *
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
import json
import sys
from pygtail import Pygtail
import subprocess
import webbrowser
from tkinter.messagebox import showwarning
import threading
from tkinter.font import Font

from UI_scriptFunctions import select_folder,begin_order_processing,open_folder, open_folder_packaging, select_files
from config import ConfigFolderPath, headingFont,fieldFont,buttonFont,labelFont,pathFont,logFont,CLIENTSFOLDERPATH, ICONIMAGE
from UI_tabs import Tab1, Tab2

# from UI_logscmd import PrintLogger


print('Initialzing Program...')


# Getting Screen Dimensions
DeviceScreenHeight = ''
DeviceScreenWidth = ''
for m in get_monitors():
   DeviceScreenWidth =str(m.width)
   DeviceScreenHeight = str(m.height)


root = Tk()
# root.attributes('-fullscreen', True)
root.geometry(DeviceScreenWidth+"x"+DeviceScreenHeight)
root.minsize(int(DeviceScreenWidth)-300,int(DeviceScreenHeight)-200)
root.maxsize(int(DeviceScreenWidth),int(DeviceScreenHeight))

root.state("zoomed")
# root.iconbitmap('C:/Users/HP/Documents/GitHub/Triumph.ico')

root.title("Purchase Orders")

with open(ConfigFolderPath, 'r') as jsonFile:
  config = json.load(jsonFile)
  themepath = config['appTheme']
  themeColor = config['themeColor']

root.tk.call("source", themepath)
root.tk.call("set_theme", themeColor)


s = ttk.Style()
s.configure('TNotebook.Tab', font=('Calibri','15'), padding=[100, 10])

# print(root.winfo_height()-20,root.winfo_height())
tabFrame = ttk.Frame(root)
logFrame = ttk.Frame(root)

root.columnconfigure(0, weight=1)
root.rowconfigure(1, weight=95) # 80%
root.rowconfigure(2, weight=5) # 10%

tabControl = ttk.Notebook(tabFrame)


filenames = tk.StringVar() 
selectedDate = tk.StringVar() 


# TAB 1
print('Initialzing PO Orders Screen...')
tab1 = Tab1(root,tabControl)
# TAB 2
print('Initialzing Packing-Slip Screen...')
tab2 = Tab2(root,tabControl)
print('Loading Screen...')
tabControl.pack(fill ="x")
# tabFrame.pack(side='top',anchor=NW,fill ="x") #,pady=(0,40)

root.iconbitmap(ICONIMAGE)
def callback(url):
   webbrowser.open_new_tab(url)


footer_frame = ttk.Frame(root)
inside_footer_frame = ttk.Frame(footer_frame)
inside_footer_frame.grid(row=0,column=0)

Developed_by = Label(inside_footer_frame, text="Developed by   -  ",font=Font(size=10,weight="bold"))
Developed_by.grid(row=0,column=1)

Developed_text = Label(inside_footer_frame, text="C-BIA Solutions & Services LLP",font=Font(size=10))
Developed_text.grid(row=0,column=2)

Website = Label(inside_footer_frame, text="          Website  -  ",font=Font(size=10,weight="bold"))
Website.grid(row=0,column=3)

version = Label(inside_footer_frame, text="          Version  -  ",font=Font(size=10,weight="bold"))
version.grid(row=0, column=5)

with open(ConfigFolderPath, 'r') as jsonFile:
   config = json.load(jsonFile)
   versionValue = Label(inside_footer_frame, text=config['version'],font=Font(size=10))
   versionValue.grid(row=0, column=6)

link = Label(inside_footer_frame, text="https://www.c-bia.com" ,font=Font(size=10, underline=1), cursor="hand2")
link.grid(row=0,column=4)
link.bind("<Button-1>", lambda e:
callback("https://www.c-bia.com/"))

# https://docs.python.org/3/library/tkinter.html#threading-model

print('UI Loaded...')

tabFrame.grid(row=1, sticky='news')
footer_frame.grid(row=2,pady=10)



root.mainloop()  