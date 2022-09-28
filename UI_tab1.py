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
import sys

from UI_scriptFunctions import select_folder,beginOrderProcessing,openfolder
from UI_logscmd import PrintLogger
from config import ConfigFolderPath, headingFont,fieldFont,buttonFont,labelFont,pathFont,logFont


class Tab1():
    def __init__(self, root,tabControl):

        with open(ConfigFolderPath+'client.json', 'r') as jsonFile:
            config = json.load(jsonFile)
            ClientCode = config

        tab1 = ttk.Frame(tabControl)
        tabControl.add(tab1, text ='PO Orders')
        # Orderframe = Frame(tab1,highlightbackground="blue", highlightthickness=2)
        Orderframe = Frame(tab1)
        Orderframe.grid(row=0,column=0)
        
        Tab1Heading = Label(Orderframe,text='PO Order Processing',font=headingFont)
        Tab1Heading.grid(row=0,column=0,padx=20, pady=20,sticky=W,columnspan=2)

        ClientNameField1 = Label(Orderframe,text='Client Name',font=labelFont)
        ClientNameField1.grid(row=1,column=0,padx=20, pady=20,sticky=W)

        options = list(ClientCode.keys())

        clicked = StringVar()
        clicked.set('--select--') # Default Value selected
        ClientNameField1 = OptionMenu(Orderframe,clicked,*options)
        ClientNameField1.grid(row=1,column=1,padx=20, pady=20,sticky=W,columnspan=2)
        ClientNameField1.config(font=tkFont.Font(family='Calibri', size=15))
        menuoptions = Orderframe.nametowidget(ClientNameField1.menuname)
        menuoptions.config(font=tkFont.Font(family='Calibri', size=15))

    
        POFolderPathField1 = ttk.Label(Orderframe,text='POFolderPath',font=labelFont)
        POFolderPathField1.grid(row=2,column=0,padx=20, pady=20,sticky=W)


        POFolderPathButton1 = Button(Orderframe,text='Select Folder',command=lambda:select_folder(POFolderPathValue1),font=buttonFont)
        POFolderPathButton1.grid(row=2,column=1,padx=20, pady=20,sticky=W)

        POFolderPathValue1 = Label(Orderframe,text="No Folder selected", font=pathFont)
        POFolderPathValue1.grid(row=2,column=2,padx=20, pady=20,sticky=W)


        OrderDateField1 = Label(Orderframe,text='Order Date', font=labelFont)
        OrderDateField1.grid(row=3,column=0,padx=20, pady=20,sticky=W)

        sel= StringVar()

        OrderDateButton1 = DateEntry(Orderframe,selectmode='day',date_pattern='dd-mm-Y',textvariable=sel,font=buttonFont)
        OrderDateButton1.grid(row=3,column=1,padx=20, pady=20,sticky=W)

        ProcessButton1 = Button(Orderframe, command=threading.Thread(target=lambda:beginOrderProcessing(mode ='consolidation', client=clicked.get(), date=OrderDateButton1.get_date(), path=POFolderPathValue1['text'])).start, text="Process",font=buttonFont)
        ProcessButton1.grid(row=4,column=1,padx=20, pady=20,sticky=W)

        CancelButton1 = Button(Orderframe, text="Cancel",font=buttonFont)
        CancelButton1.grid(row=4,column=2,padx=20, pady=20,sticky=W)

        RequirementSummaryPathField1 = Label(Orderframe,text='Requirement Summary Path ', font=labelFont)
        RequirementSummaryPathField1.grid(row=5,column=0,padx=20, pady=20,sticky=W)

        with open(ConfigFolderPath+'config.json', 'r') as jsonFile:
            config = json.load(jsonFile)

            RequirementSummaryPathButton = Button(Orderframe,text='Copy Path',command=lambda:openfolder(params=[config['targetFolder'], ClientCode[clicked.get()], OrderDateButton1.get_date(), '60-Requirement-Summary'],frame=Orderframe),font=buttonFont)
            RequirementSummaryPathButton.grid(row=5,column=1,padx=20, pady=20,sticky=W)

            RequirementSummaryPath = Label(Orderframe,text='No Path selected',font=pathFont)
            RequirementSummaryPath.grid(row=5,column=2,padx=20, pady=20,sticky=W)

            def my_date_client(*argus):

                changedDate = sel.get()
                Year = changedDate[6:10]
                Month = changedDate[3:5]
                Date = changedDate[0:2]
                # print("1 ",changedDate, type(changedDate))
                # print(Year, Month, Date)
                # changedDate = changedDate + ' 00:00:00'
                # print("2", changedDate)
                # changedDate = datetime.strptime(changedDate, '%d-%m-%Y').date()
                # print("3", changedDate, type(changedDate))
                # year = changedDate.strftime('%Y')
                # date = str(changedDate.strftime('%Y-%m-%d'))
                year= Year
                date= Year+ "-"+Month+ "-"+ Date
                chnagedClientcode = clicked.get()
                
                path = config['targetFolder']+'/'+ClientCode[clicked.get()]+'-'+year+'/'+date+'/'+'60-Requirement-Summary'
                RequirementSummaryPath.config(text=path)
        
            clicked.trace('w',my_date_client)
            sel.trace('w',my_date_client)


       


