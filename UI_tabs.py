import tkinter as tk                    
from tkinter import ttk, filedialog, scrolledtext, font
from tkinter import *
from screeninfo import get_monitors
from tkcalendar import DateEntry
from tkinter.messagebox import showinfo
import json, subprocess, os, threading, base64, sys
from datetime import datetime
import tkinter.scrolledtext as st
from UI_logscmd import PrintLogger


from UI_scriptFunctions import select_folder,begin_order_processing,open_folder, open_folder_packaging, select_files
from config import ConfigFolderPath, headingFont,fieldFont,buttonFont,labelFont,pathFont,logFont,ClientsFolderPath


class Tab1():
    def __init__(self, root,tabControl):

        with open(ClientsFolderPath, 'r') as jsonFile:
            config = json.load(jsonFile)
            po_client_code = config

        po_tab = ttk.Frame(tabControl)
        tabControl.add(po_tab, text ='PO Orders')

        main_frame = ttk.Frame(po_tab)
        log_frame = ttk.Frame(po_tab)


        po_frame = Frame(main_frame,highlightbackground="blue", highlightthickness=2)
        po_frame.grid(row=0,column=0)
        
        po_heading = Label(main_frame,text='PO Order Processing',font=headingFont)
        po_heading.grid(row=0,column=0,padx=20, pady=20,sticky=W,columnspan=2)

        po_client_Name = Label(main_frame,text='Client Name',font=labelFont)
        po_client_Name.grid(row=1,column=0,padx=20, pady=20,sticky=W)

        client_options = list(po_client_code.keys()) # loads from client.json

        po_selected_client = StringVar()
        po_selected_client.set('--select--') # Default Value selected
        po_client_option = OptionMenu(main_frame,po_selected_client,*client_options)
        po_client_option.grid(row=1,column=1,padx=20, pady=20,sticky=W,columnspan=2)
        po_client_option.config(font=font.Font(family='Calibri', size=15))
        menuo_ptions = main_frame.nametowidget(po_client_option.menuname)
        menuo_ptions.config(font=font.Font(family='Calibri', size=15))

    
        po_folder_path = ttk.Label(main_frame,text='POFolderPath',font=labelFont)
        po_folder_path.grid(row=2,column=0,padx=20, pady=20,sticky=W)


        po_folder_path_btn = Button(main_frame,text='Select Folder',command=lambda:select_folder(po_folder_path_value),font=buttonFont)
        po_folder_path_btn.grid(row=2,column=1,padx=20, pady=20,sticky=W)

        po_folder_path_value = Label(main_frame,text="No Folder selected", font=pathFont, wraplength=800)
        po_folder_path_value.grid(row=2,column=2,padx=20, pady=20,sticky=W)


        po_order_date = Label(main_frame,text='Order Date', font=labelFont)
        po_order_date.grid(row=3,column=0,padx=20, pady=20,sticky=W)

        po_selected_date= StringVar()
        po_order_date_btn = DateEntry(main_frame,selectmode='day',date_pattern='dd-mm-Y',textvariable=po_selected_date,font=buttonFont)
        po_order_date_btn.grid(row=3,column=1,padx=20, pady=20,sticky=W)

        po_process_btn = Button(main_frame, command=threading.Thread(target=lambda:begin_order_processing(mode ='consolidation', client=po_selected_client.get(), date=po_order_date_btn.get_date(), path=po_folder_path_value['text'])).start, text="Process",font=buttonFont)
        po_process_btn.grid(row=4,column=1,padx=20, pady=20,sticky=W)

        po_cancel_btn = Button(main_frame, text="Cancel",font=buttonFont)
        po_cancel_btn.grid(row=4,column=2,padx=20, pady=20,sticky=W)

        po_requirements_summary = Label(main_frame,text='Requirements Summary Path ', font=labelFont)
        po_requirements_summary.grid(row=5,column=0,padx=20, pady=20,sticky=W)

        with open(ConfigFolderPath+'config.json', 'r') as jsonFile:
            config = json.load(jsonFile)

            po_requirements_summary_btn = Button(main_frame,text='Copy Path',command=lambda:open_folder(params=[config['targetFolder'], po_client_code[po_selected_client.get()], po_order_date_btn.get_date(), '60-Requirement-Summary'],frame=main_frame),font=buttonFont)
            po_requirements_summary_btn.grid(row=5,column=1,padx=20, pady=20,sticky=W)

            po_requirements_summary_path = Label(main_frame,text='No Path selected',font=pathFont, wraplength=800)
            po_requirements_summary_path.grid(row=5,column=2,padx=20, pady=20,sticky=W)

            def po_date_client(*argus):

                po_changed_date = po_selected_date.get()
                year = po_changed_date[6:10]
                date = po_changed_date[6:10] + "-"+ po_changed_date[3:5] + "-" + po_changed_date[0:2]
                po_changed_client_code = po_client_code[po_selected_client.get()]
                
                temp_req_sum_path = config['targetFolder']+'/'+po_changed_client_code+'-'+year+'/'+date+'/'+'60-Requirement-Summary'
                po_requirements_summary_path.config(text=temp_req_sum_path,wraplength=800)
        
            po_selected_client.trace('w',po_date_client)
            po_selected_date.trace('w',po_date_client)


        main_frame.pack(side='top',anchor=NW,fill ="x")
        
        consoleLabel = Label(log_frame,text='Console logs', font=labelFont)
        consoleLabel.pack(side='top',anchor=NW, padx=20,pady=10)

        # logbox = st.ScrolledText(log_frame)

        logbox = st.ScrolledText(log_frame)
        logbox.pack(expand = True,fill ="both",ipady=20,ipadx=10)

        pl = PrintLogger(log_frame)
        # logbox.insert(tk.INSERT,'Logs:')
        # sys.stdout = pl
        # logbox.pack(expand = True,fill ="both",ipady=20,ipadx=10)
        log_frame.pack(fill ="x")


class Tab2():
    def __init__(self, root,tabControl):
        pkg_client_var = StringVar()
        pkg_date_var = StringVar()

        pkg_tab = ttk.Frame(tabControl)
        tabControl.add(pkg_tab, text ='Packing Slip')

        pkg_frame = Frame(pkg_tab)
        pkg_frame.grid(row=0,column=0)
        pkg_heading = Label(pkg_frame,text='Generating Packing Slip',font=headingFont)
        pkg_heading.grid(row=0,column=0,padx=20, pady=20,sticky=W,columnspan=2)  

        pkg_requirements_summary = ttk.Label(pkg_frame,text='Requirements Summary Path',font=labelFont)
        pkg_requirements_summary.grid(row=1,column=0,padx=20, pady=20,sticky=W)

        pkg_requirements_summary_btn = Button(pkg_frame,text='Select File',command=lambda:select_files(pkg_requirements_summary_path, pkg_client_var,pkg_date_var), font=buttonFont)
        pkg_requirements_summary_btn.grid(row=1,column=1,padx=20, pady=20,sticky=W)

        pkg_requirements_summary_path = Label(pkg_frame,text="No Path Selected", font=pathFont, wraplength=800)
        pkg_requirements_summary_path.grid(row=1,column=2,padx=20, pady=20,sticky=W)

        pkg_client_Name = Label(pkg_frame,text='Client Name',font=labelFont)
        pkg_client_Name.grid(row=2,column=0,padx=20, pady=20,sticky=W)
        
        pkg_client_Name_value = Label(pkg_frame, textvariable=pkg_client_var,font=labelFont)
        pkg_client_Name_value.grid(row=2,column=1,padx=20, pady=20,sticky=W,columnspan=2)

        pkg_order_date = Label(pkg_frame,text='Order Date',font=labelFont)
        pkg_order_date.grid(row=3,column=0,padx=20, pady=20,sticky=W)

       
        pkg_order_date_value = Label(pkg_frame, textvariable=pkg_date_var,font=labelFont)
        pkg_order_date_value.grid(row=3,column=1,padx=20, pady=20,sticky=W)

        pkg_process_btn = Button(pkg_frame, command=threading.Thread(target=lambda:begin_order_processing(mode ='packing', client=pkg_client_Name_value.cget("text"), date=pkg_order_date_value.cget("text"), path=pkg_requirements_summary_path.cget("text"))).start, text="Process",font=buttonFont)
        pkg_process_btn.grid(row=4,column=1,padx=20, pady=20,sticky=W)

        pkg_cancel_btn = Button(pkg_frame, text="Cancel", font=buttonFont)
        pkg_cancel_btn.grid(row=4,column=2,padx=20, pady=20,sticky=W)

        pkg_packing_slip = Label(pkg_frame,text='Packing slip Folder Path ', font=labelFont)
        pkg_packing_slip.grid(row=5,column=0,padx=20, pady=20,sticky=W)

        with open(ConfigFolderPath+'config.json', 'r') as jsonFile:

            config = json.load(jsonFile)

            pkg_packing_slip_btn = Button(pkg_frame,text='Copy Path',command=lambda:open_folder_packaging(params=[config['targetFolder'], pkg_client_var.get(), pkg_date_var.get(), '70-Packing-Slip'],frame=pkg_frame),font=buttonFont)
            pkg_packing_slip_btn.grid(row=5,column=1,padx=20, pady=20,sticky=W)

            pkg_packing_slip_folder_path = Label(pkg_frame,text='No Path Selected', font=pathFont, wraplength=800)
            pkg_packing_slip_folder_path.grid(row=5,column=2,padx=20, pady=20,sticky=W)


            def pkg_date_client(*argus):

                pkg_changed_date = pkg_date_var.get()
                pkg_year = pkg_changed_date[6:10]
                pkg_date = pkg_changed_date[6:10]+ "-" + pkg_changed_date[3:5] + "-"+ pkg_changed_date[0:2]
                pkg_changed_client = pkg_client_var.get()

                
                with open(ClientsFolderPath, 'r') as jsonFile:
                    pkg_client_code = json.load(jsonFile)
                    pkg_client_code = pkg_client_code[pkg_changed_client]

                temp_pkg_slip_path = config['targetFolder']+'/'+pkg_client_code+'-'+pkg_year+'/'+pkg_date+'/'+'70-Packing-Slip'
                pkg_packing_slip_folder_path.config(text=temp_pkg_slip_path)

            pkg_client_var.trace('w',pkg_date_client)
            pkg_date_var.trace('w',pkg_date_client)





