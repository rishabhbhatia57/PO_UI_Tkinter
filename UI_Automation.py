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

from UI_scriptFunctions import select_folder,begin_order_processing,open_folder, open_folder_packaging, select_files
from config import ConfigFolderPath, headingFont,fieldFont,buttonFont,labelFont,pathFont,logFont,ClientsFolderPath
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
print('Initialzing PO Orders Screen...')
tab1 = Tab1(root,tabControl)
# TAB 2
print('Initialzing Packing-Slip Screen...')
tab2 = Tab2(root,tabControl)
print('Loading Screen...')
tabControl.pack(fill ="x")



tabFrame.pack(side='top',anchor=NW,fill ="x")


def callback(url):
   webbrowser.open_new_tab(url)

footer_frame = ttk.Frame(root)
inside_footer_frame = ttk.Frame(footer_frame)
inside_footer_frame.grid(row=0,column=0)

footerText = Label(inside_footer_frame, text="Developed by: C-BIA Solutions & Services LLP ",font=(logFont,10), cursor="hand2")
footerText.grid(row=0,column=1)

link = Label(inside_footer_frame, text="     Website:  https://c-bia.com/" ,font=(logFont,10), cursor="hand2")
link.grid(row=0,column=2)
link.bind("<Button-1>", lambda e:
callback("https://c-bia.com/"))


footer_frame.pack()


print('UI Loaded...')



# logbox = st.ScrolledText(logframe)
# logbox.configure(state ='disabled')
# consoleLabel = Label(logframe,text='Logs', font=labelFont)
# consoleLabel.pack(side=tk.LEFT)
# logbox.pack(logframe,expand = True,fill ="both",ipady=10,ipadx=10)

root.mainloop()  