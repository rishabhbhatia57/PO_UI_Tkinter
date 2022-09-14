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
import subprocess
from datetime import datetime
from logscmd import PrintLogger
import sys


from commanfunc import select_folder,selectedFun,openfolder

path1 = 'C:/Users/HP/Desktop/PO Metadata/Configfiles-Folder/'



class Tab1():
    def __init__(self, root,tabControl):

        with open(path1+'client.json', 'r') as jsonFile:
            config = json.load(jsonFile)
            print(config)

            ClientCode = config

        headingFont = ("Calibri",40,"bold")
        fieldFont=  ("Calibri",20,"bold")
        buttonFont=  ("Calibri",15,"bold")

        tab1 = ttk.Frame(tabControl)
        tabControl.add(tab1, text ='PO Orders')
        # Orderframe = Frame(tab1,highlightbackground="blue", highlightthickness=2)
        Orderframe = Frame(tab1)
        Orderframe.grid(row=0,column=0)
        
        Tab1Heading = Label(Orderframe,text='PO Order Processing',font=headingFont)
        Tab1Heading.grid(row=0,column=0,padx=20, pady=20,sticky=W)

        ClientNameField = Label(Orderframe,text='Client Name',font=fieldFont)
        ClientNameField.grid(row=1,column=0,padx=20, pady=20,sticky=W)

        options = list(ClientCode.keys())

        clicked = StringVar()
        clicked.set(options[0]) # Default Value selected
        ClientNameEntry = OptionMenu(Orderframe,clicked,*options)
        ClientNameEntry.grid(row=1,column=1,padx=20, pady=20,sticky=W)
        ClientNameEntry.config(font=tkFont.Font(family='Calibri', size=15))
        menuoptions = Orderframe.nametowidget(ClientNameEntry.menuname)
        menuoptions.config(font=tkFont.Font(family='Calibri', size=15))

        
        

        POFolderPathField = ttk.Label(Orderframe,text='POFolderPath')
        POFolderPathField.grid(row=2,column=0,padx=20, pady=20,sticky=W)


        POFolderPathEntry = Button(Orderframe,text='Select Folder',command=lambda:select_folder(POFolderPathText), font=buttonFont)
        POFolderPathEntry.grid(row=2,column=1,padx=20, pady=20,sticky=W)

        POFolderPathText = Label(Orderframe,text="No Folder selected",wraplength=300,font=("Calibri",10,"bold"))
        POFolderPathText.grid(row=2,column=2,padx=20, pady=20,sticky=W)


        OrderDateField = Label(Orderframe,text='Order Date',font=fieldFont)
        OrderDateField.grid(row=3,column=0,padx=20, pady=20,sticky=W)

        sel= StringVar()

        OrderDateEntry = DateEntry(Orderframe,selectmode='day',textvariable=sel,font=fieldFont)
        OrderDateEntry.grid(row=3,column=1,padx=20, pady=20,sticky=W)

        ProcessButton = Button(Orderframe, command=threading.Thread(target=lambda:selectedFun(mode ='consolidation', client=clicked.get(), date=OrderDateEntry.get_date(), path=POFolderPathText['text'])).start, text="Process", font=buttonFont)
        ProcessButton.grid(row=4,column=1,padx=20, pady=20,sticky=W)

        CancelButton = Button(Orderframe, text="Cancel", font=buttonFont)
        CancelButton.grid(row=4,column=2,padx=20, pady=20,sticky=W)

        RequirementSummaryPathField = Label(Orderframe,text='Requirement Summary Path: ',font=fieldFont)
        RequirementSummaryPathField.grid(row=5,column=0,padx=20, pady=20,sticky=W)

        with open(path1+'config.json', 'r') as jsonFile:
            config = json.load(jsonFile)

            
            
            RequirementSummaryPathButton = Button(Orderframe,text='Get Path',command=lambda:openfolder(params=[config['targetFolder'], ClientCode[clicked.get()], OrderDateEntry.get_date(), '60-Requirement-Summary'],frame=Orderframe))
            RequirementSummaryPathButton.grid(row=5,column=1,padx=20, pady=20,sticky=W)

            RequirementSummaryPath = Label(Orderframe,text='No Path selected')
            RequirementSummaryPath.grid(row=5,column=2,padx=20, pady=20,sticky=W)

            def my_date_client(*argus):

                changedDate = sel.get()
                
                changedDate = datetime.strptime(changedDate, '%m/%d/%y')
                print(changedDate)
                year = changedDate.strftime('%Y')
                date = str(changedDate.strftime('%Y-%m-%d'))
                chnagedClientcode = clicked.get()
                
                path = config['targetFolder']+'/'+ClientCode[clicked.get()]+'-'+year+'/'+date+'/'+'60-Requirement-Summary'
                RequirementSummaryPath.config(text=path)
        
            clicked.trace('w',my_date_client)
            sel.trace('w',my_date_client)


       


