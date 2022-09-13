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
from commanfunc import select_files,selectedFun

class Tab2():
    def __init__(self, root,tabControl):

        headingFont = ("Calibri",40,"bold")
        fieldFont=  ("Calibri",20,"bold")
        buttonFont=  ("Calibri",15,"bold")
        tab2 = ttk.Frame(tabControl)
        tabControl.add(tab2, text ='Packaging Slip')

        packagingframe = Frame(tab2)
        packagingframe.grid(row=0,column=0)
        Tab2Heading = Label(packagingframe,text='Generating Packaging Slip',font=headingFont)
        Tab2Heading.grid(row=0,column=0,padx=20, pady=20,sticky=W)

        RequirementSummaryPathField = ttk.Label(packagingframe,text='Requirement Summary Path',font=fieldFont)
        RequirementSummaryPathField.grid(row=1,column=0,padx=20, pady=20,sticky=W)

        RequirementSummaryButton = Button(packagingframe,text='Select File',command=lambda:select_files(RequirementSummaryPathText,ClientNameValue, OrderDateValue),font=buttonFont)
        RequirementSummaryButton.grid(row=1,column=1,padx=20, pady=20,sticky=W)

        RequirementSummaryPathText = Label(packagingframe,text="Path not selected",wraplength=300,font=("Calibri",10,"bold"))
        RequirementSummaryPathText.grid(row=1,column=3,padx=20, pady=20,sticky=W)

        ClientNameField = Label(packagingframe,text='Client Name',font=fieldFont)
        ClientNameField.grid(row=2,column=0,padx=20, pady=20,sticky=W)

        ClientNameValue = Label(packagingframe,text='',font=fieldFont)
        ClientNameValue.grid(row=2,column=1,padx=20, pady=20,sticky=W)

        OrderDateField = Label(packagingframe,text='Order Date',font=fieldFont)
        OrderDateField.grid(row=3,column=0,padx=20, pady=20,sticky=W)

        OrderDateValue = Label(packagingframe,text='',font=fieldFont)
        OrderDateValue.grid(row=3,column=1,padx=20, pady=20,sticky=W)

        ProcessButton = Button(packagingframe, command=threading.Thread(target=lambda:selectedFun(mode ='packaging', client=ClientNameValue.cget("text"), date=OrderDateValue.cget("text"), path=RequirementSummaryPathText.cget("text"))).start, text="Process", font=buttonFont)
        ProcessButton.grid(row=4,column=1,padx=20, pady=20,sticky=W)

        CancelButton = Button(packagingframe, text="Cancel", font=buttonFont)
        CancelButton.grid(row=4,column=2,padx=20, pady=20,sticky=W)

        PackagingSlipFolderPath = Label(packagingframe,text='Packaging slip Folder Path: ',font=fieldFont)
        PackagingSlipFolderPath.grid(row=5,column=0,padx=20, pady=20,sticky=W)

        PackagingSlipFolderValue = Label(packagingframe,text='Path',font=fieldFont)
        PackagingSlipFolderValue.grid(row=5,column=1,padx=20, pady=20,sticky=W)


        pass