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
import json
import sys
from pygtail import Pygtail

from config import ConfigFolderPath, headingFont,fieldFont,buttonFont,labelFont,pathFont,logFont
from UI_tab1 import Tab1
from UI_tab2 import Tab2
from UI_logscmd import PrintLogger
import subprocess
from tkinter.messagebox import showwarning

from globalvar import logboxstate



# Getting Screen Dimensions
DeviceScreenHeight = ''
DeviceScreenWidth = ''
for m in get_monitors():
   DeviceScreenWidth =str(m.width)
   DeviceScreenHeight = str(m.height)

root = Tk()
root.geometry(DeviceScreenWidth+"x"+DeviceScreenHeight)
root.minsize(1700,1000)
root.maxsize(int(DeviceScreenWidth),int(DeviceScreenHeight))


root.title("Purchase Orders")

with open(ConfigFolderPath+'config.json', 'r') as jsonFile:
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


# TAB 1
tab1 = Tab1(root,tabControl)
# TAB 2
tab2 = Tab2(root,tabControl)
# TAB 3
# tab3 = Tab3(root,tabControl)

tabControl.pack(fill ="x")




logbox = st.ScrolledText(root)
logbox.pack(expand = True,fill ="both",ipady=10,ipadx=10)
global pl 
pl = PrintLogger(logbox)

# sys.stdout = pl
# logbox.insert(tk.INSERT,pl) # Inserting Text which is read only
while logboxstate:
  for line in Pygtail("C:/Users/HP/Documents/GitHub/PO_UI_Tkinter/log.txt"):
      logbox.insert(tk.INSERT,line) # Inserting Text which is read only
# logbox.insert(tk.INSERT,pl) # Inserting Text which is read only
logbox.configure(state ='disabled')



root.mainloop()  