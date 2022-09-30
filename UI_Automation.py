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

from config import ConfigFolderPath, headingFont,fieldFont,buttonFont,labelFont,pathFont,logFont
from UI_tab1 import Tab1
from UI_tab2 import Tab2
from UI_logscmd import TextHandler
import subprocess
from tkinter.messagebox import showwarning


print('Initialzing Program...')


# Getting Screen Dimensions
DeviceScreenHeight = ''
DeviceScreenWidth = ''
for m in get_monitors():
   DeviceScreenWidth =str(m.width)
   DeviceScreenHeight = str(m.height)

root = Tk()
root.geometry(DeviceScreenWidth+"x"+DeviceScreenHeight)
root.minsize(int(DeviceScreenWidth)-300,int(DeviceScreenHeight)-250)
root.maxsize(int(DeviceScreenWidth),int(DeviceScreenHeight))

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
print('Initialzing PO Orders Screen...')
tab1 = Tab1(root,tabControl)



# TAB 2
print('Initialzing Packing-Slip Screen...')
tab2 = Tab2(root,tabControl)
print('Loading Screen...')
tabControl.pack(fill ="x")


consoleLabel = Label(logFrame,text='Console', font=labelFont)
consoleLabel.pack(side='top',anchor=NW, padx=20,pady=20)
logbox = st.ScrolledText(logFrame)
logbox.configure(state ='disabled')
# logbox.insert(tk.INSERT,'Logs:')
logbox.pack(expand = True,fill ="both",ipady=20,ipadx=10)

tabFrame.pack(side='top',anchor=NW,fill ="x")
logFrame.pack(side='bottom',anchor=SW,fill ="x")
print('UI Loaded...')



# logbox = st.ScrolledText(logframe)
# logbox.configure(state ='disabled')
# consoleLabel = Label(logframe,text='Logs', font=labelFont)
# consoleLabel.pack(side=tk.LEFT)
# logbox.pack(logframe,expand = True,fill ="both",ipady=10,ipadx=10)

root.mainloop()  