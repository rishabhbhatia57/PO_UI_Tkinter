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
from openpyxl import load_workbook,Workbook
import openpyxl.utils.cell




def select_folder(showPath):
    filetypes = (
        ('pdf files', '*.pdf'),
        ('All files', '*.*')
    )
    global POFolderSelected 
    POFolderSelected = fd.askdirectory()
    showPath.config(text=POFolderSelected)
    
    showinfo(
        title='Selected Folder',
        message=POFolderSelected
    )
    return POFolderSelected

def select_files(showPath,showReqPath,showOrderdate):
    filetypes = (
        ('pdf files', '*.pdf'),
        ('All files', '*.*')
    )
    global ReqFileSelected 
    ReqFileSelected = fd.askopenfilename()
    showPath.config(text=ReqFileSelected)

    ReqSumWorkbook = load_workbook(ReqFileSelected,data_only=True) 
    ReqSumSheet = ReqSumWorkbook.active


    showReqPath.config(text=ReqSumSheet.cell(2,6).value)
    showOrderdate.config(text=ReqSumSheet.cell(2,4).value)

    
    showinfo(
        title='Selected Files',
        message=ReqFileSelected
    )
    return ReqFileSelected

def openfolder(Targetfolder,clientcode,date):
    Targetpath = Targetfolder+'/'+clientcode+'-'+date.strftime('%Y')+'/'+str(date)

    # RequirementSummarypath = Targetfolder+'/'+clientcode+'-'+date.strftime('%Y')+'/'+str(date)+'/60-Requirement-Summary'
    path = os.path.realpath(Targetpath)
    os.startfile(path)

def selectedFun(mode, client, date,path):
    ClientCode = {
  "Pantaloons": "PL",
  "Shoppers Stop Limited": "SSL",
  "Lifestyle Limited": "LSL"
}

    ClientCodeSelected = client
    OrderDateSelected = date
    requestedpath = path
    if ClientCodeSelected == '' or OrderDateSelected == '' or requestedpath == '':
        showinfo(
        title='Invalid Selection',
        message="Invalid Client Name, path or Order Date Selected"
    )
    else:
        with open('C:/Users/HP/Desktop/PO Metadata/Configfiles-Folder/config.json', 'r') as jsonFile:
            config = json.load(jsonFile)    
            pythonenvpath = config['pythonPath']
            pythonScriptPath = config['appsScriptPath']

        encodedClientCodeSelected = ClientCode[ClientCodeSelected]
        encodedOrderDateSelected = str(OrderDateSelected).replace(' ', "#")
        enodedPOFolderSelected = requestedpath.replace(' ', "#") 

          
        # pb = ttk.Progressbar(
        #     Orderframe,
        #     orient='horizontal',
        #     mode='indeterminate',
        #     length=500
        # )
        # pb.grid(row=8,column=2,padx=0, pady=20,sticky=tk.W)  
        # pb.start

        # Running code on CMD
        os.system(pythonenvpath +" "+ pythonScriptPath+" "+mode+" "+encodedClientCodeSelected+" "+encodedOrderDateSelected+" "+enodedPOFolderSelected)
        print('Here')

        