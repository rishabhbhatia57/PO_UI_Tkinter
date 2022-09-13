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


class Tab3():
    def __init__(self,root,tabControl):

        headingFont = ("Calibri",40,"bold")
        fieldFont=  ("Calibri",20,"bold")
        tab3 = ttk.Frame(tabControl)
        tabControl.add(tab3, text ='Logs')
        logsframe = Frame(tab3)
        logsframe.grid(row=0,column=0)

        Tab3Heading = Label(logsframe,text='Logs',font=headingFont)
        Tab3Heading.grid(row=0,column=0,padx=20, pady=20,sticky=W)

        text_area = st.ScrolledText(logsframe, width = 140, height = 28, font = ("Times New Roman", 13))
        text_area.grid(column = 0, pady = 10, padx = 10)


        text_area.insert(tk.INSERT,"""
        juhwdabhk
        """) # Inserting Text which is read only
        text_area.configure(state ='disabled') # Making the text read only