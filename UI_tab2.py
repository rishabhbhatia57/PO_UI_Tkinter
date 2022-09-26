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
from datetime import datetime

from config import ConfigFolderPath, headingFont,fieldFont,buttonFont,labelFont,pathFont,logFont
from UI_scriptFunctions import select_files,selectedFun, openfolderpackaging



class Tab2():
    def __init__(self, root,tabControl):
        client = StringVar()
        date = StringVar()

        tab2 = ttk.Frame(tabControl)
        tabControl.add(tab2, text ='Packaging Slip')

        packagingframe = Frame(tab2)
        packagingframe.grid(row=0,column=0)
        Tab2Heading = Label(packagingframe,text='Generating Packaging Slip',font=headingFont)
        Tab2Heading.grid(row=0,column=0,padx=20, pady=20,sticky=W,columnspan=2)  

        RequirementSummaryPathField2 = ttk.Label(packagingframe,text='Requirement Summary Path',font=labelFont)
        RequirementSummaryPathField2.grid(row=1,column=0,padx=20, pady=20,sticky=W)

        RequirementSummaryButton2 = Button(packagingframe,text='Select File',command=lambda:select_files(RequirementSummaryPathValue2, client,date), font=buttonFont)
        RequirementSummaryButton2.grid(row=1,column=1,padx=20, pady=20,sticky=W)

        RequirementSummaryPathValue2 = Label(packagingframe,text="No Path Selected", font=pathFont)
        RequirementSummaryPathValue2.grid(row=1,column=2,padx=20, pady=20,sticky=W)

        ClientNameField2 = Label(packagingframe,text='Client Name',font=labelFont)
        ClientNameField2.grid(row=2,column=0,padx=20, pady=20,sticky=W)
        
        
        ClientNameValue2 = Label(packagingframe, textvariable=client,font=labelFont)
        ClientNameValue2.grid(row=2,column=1,padx=20, pady=20,sticky=W,columnspan=2)

        OrderDateField2 = Label(packagingframe,text='Order Date',font=labelFont)
        OrderDateField2.grid(row=3,column=0,padx=20, pady=20,sticky=W)

       
        OrderDateValue2 = Label(packagingframe, textvariable=date,font=labelFont)
        OrderDateValue2.grid(row=3,column=1,padx=20, pady=20,sticky=W)

        ProcessButton2 = Button(packagingframe, command=threading.Thread(target=lambda:selectedFun(mode ='packaging', client=ClientNameValue2.cget("text"), date=OrderDateValue2.cget("text"), path=RequirementSummaryPathValue2.cget("text"))).start, text="Process",font=buttonFont)
        ProcessButton2.grid(row=4,column=1,padx=20, pady=20,sticky=W)

        CancelButton2 = Button(packagingframe, text="Cancel", font=buttonFont)
        CancelButton2.grid(row=4,column=2,padx=20, pady=20,sticky=W)

        PackagingSlipFolderPath2 = Label(packagingframe,text='Packaging slip Folder Path ', font=labelFont)
        PackagingSlipFolderPath2.grid(row=5,column=0,padx=20, pady=20,sticky=W)

        with open(ConfigFolderPath+'config.json', 'r') as jsonFile:

            config = json.load(jsonFile)

            PackagingSummaryPathButton2 = Button(packagingframe,text='Copy Path',command=lambda:openfolderpackaging(params=[config['targetFolder'], client.get(), date.get(), '70-Packaging-Slip'],frame=packagingframe),font=buttonFont)
            PackagingSummaryPathButton2.grid(row=5,column=1,padx=20, pady=20,sticky=W)

            PackagingSlipFolderValue2 = Label(packagingframe,text='No Path Selected', font=pathFont)
            PackagingSlipFolderValue2.grid(row=5,column=2,padx=20, pady=20,sticky=W)


            def my_date_client(*argus):

                changedDate = date.get()
                print(date.get())
                changedDate = datetime.strptime(changedDate, '%Y-%m-%d')

                year = changedDate.strftime('%Y')
                date1 = str(changedDate.strftime('%Y-%m-%d'))
                changedclient = client.get()
                with open(ConfigFolderPath+'client.json', 'r') as jsonFile:
                    clientcode = json.load(jsonFile)
                    clientcode = clientcode[changedclient]

                path = config['targetFolder']+'/'+clientcode+'-'+year+'/'+date1+'/'+'70-Packaging-Slip'
                PackagingSlipFolderValue2.config(text=path)

            client.trace('w',my_date_client)
            date.trace('w',my_date_client)



        