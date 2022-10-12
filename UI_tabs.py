import tkinter as tk                    
from tkinter import ttk, filedialog, scrolledtext, font
from tkinter import *
from screeninfo import get_monitors
from tkcalendar import DateEntry
from tkinter.messagebox import showinfo
import json, subprocess, os, threading, base64, sys
from datetime import datetime
import tkinter.scrolledtext as st
from UI_logscmd import ConsoleUi
from PIL import ImageTk, Image



from UI_scriptFunctions import select_folder,begin_order_processing,open_folder, open_folder_packaging, select_files
from config import ConfigFolderPath, headingFont,fieldFont,buttonFont,labelFont,pathFont,logFont,ClientsFolderPath




class Tab1():
    def __init__(self, root,tabControl):

        with open(ClientsFolderPath, 'r') as jsonFile:
            config = json.load(jsonFile)
            po_client_code = config

        po_tab = ttk.Frame(tabControl)
        tabControl.add(po_tab, text ='PO Orders')

        po_main_frame = ttk.Frame(po_tab)
        po_log_frame = ttk.Frame(po_tab)


        # po_frame = Frame(po_main_frame,highlightbackground="blue", highlightthickness=2)
        # po_frame.grid(row=0,column=0)
        
        po_heading = Label(po_main_frame,text='PO Order Processing',font=headingFont)
        po_heading.grid(row=0,column=0,padx=10, pady=10,sticky=W,columnspan=2)


        self.po_img = ImageTk.PhotoImage(Image.open("./config/Triumph_International_Logo.png"))
        po_image_frame = Frame(po_main_frame) #,  highlightbackground="blue", highlightthickness=2, height=10,width=10
        po_image_frame.grid(row=0,column=2,columnspan=3)
        po_image_frame.place(anchor='ne', relx=0.99, rely=0.03)
        po_img_label = Label(po_image_frame, image = self.po_img)
        po_img_label.pack()
        

        po_client_Name = Label(po_main_frame,text='Client Name',font=labelFont)
        po_client_Name.grid(row=1,column=0,padx=10, pady=10,sticky=W)

        client_options = list(po_client_code.keys()) # loads from client.json

        po_selected_client = StringVar()
        po_selected_client.set('--select--') # Default Value selected
        po_client_option = OptionMenu(po_main_frame,po_selected_client,*client_options)
        po_client_option.grid(row=1,column=1,padx=10, pady=10,sticky=W,columnspan=4)
        po_client_option.config(font=font.Font(family='Calibri', size=15))
        menuo_ptions = po_main_frame.nametowidget(po_client_option.menuname)
        menuo_ptions.config(font=font.Font(family='Calibri', size=15))

        po_order_date = Label(po_main_frame,text='Order Date', font=labelFont)
        po_order_date.grid(row=1,column=3,padx=10, pady=10,sticky=W)

        po_selected_date= StringVar()
        po_order_date_btn = DateEntry(po_main_frame,selectmode='day',date_pattern='dd-mm-Y',textvariable=po_selected_date,font=buttonFont)
        po_order_date_btn.grid(row=1,column=4,padx=10, pady=10,sticky=W)

    
        po_folder_path = ttk.Label(po_main_frame,text='POFolderPath',font=labelFont)
        po_folder_path.grid(row=2,column=0,padx=10, pady=10,sticky=W)


        po_folder_path_btn = Button(po_main_frame,text='Select Folder',command=lambda:select_folder(po_folder_path_value),font=buttonFont)
        po_folder_path_btn.grid(row=2,column=1,padx=10, pady=10,sticky=W)

        po_folder_path_value = Label(po_main_frame,text="No Folder selected", font=pathFont, wraplength=800)
        po_folder_path_value.grid(row=2,column=2,padx=10, pady=10,sticky=W, columnspan=8)


        po_thread = threading.Thread(target=lambda:begin_order_processing(mode ='consolidation', client=po_selected_client.get(), 
        date=po_order_date_btn.get_date(), path=po_folder_path_value['text'], consoleLabel=po_consoleLabel, thread_name=po_thread),
        name='po_thread').start
        po_process_btn = Button(po_main_frame, command=po_thread, text="Process",font=buttonFont)
        # po_process_btn = Button(po_main_frame, command=lambda:begin_order_processing(mode ='consolidation', client=po_selected_client.get(), date=po_order_date_btn.get_date(), path=po_folder_path_value['text'], consoleLabel=po_consoleLabel), text="Process",font=buttonFont)
        po_process_btn.grid(row=4,column=1,padx=10, pady=10,sticky=W)

        po_cancel_btn = Button(po_main_frame, text="Cancel",font=buttonFont)
        po_cancel_btn.grid(row=4,column=2,padx=10, pady=10,sticky=W)

        po_requirements_summary = Label(po_main_frame,text='Requirements Summary Path ', font=labelFont)
        po_requirements_summary.grid(row=5,column=0,padx=10, pady=10,sticky=W)

        with open(ConfigFolderPath+'config.json', 'r') as jsonFile:
            config = json.load(jsonFile)

            po_requirements_summary_btn = Button(po_main_frame,text='Copy Path',command=lambda:open_folder(params=[config['targetFolder'], po_client_code[po_selected_client.get()], po_order_date_btn.get_date(), '60-Requirement-Summary'],frame=po_main_frame),font=buttonFont)
            po_requirements_summary_btn.grid(row=5,column=1,padx=10, pady=10,sticky=W)

            po_requirements_summary_path = Label(po_main_frame,text='No Path selected',font=pathFont, wraplength=800)
            po_requirements_summary_path.grid(row=5,column=2,padx=10, pady=10,sticky=W, columnspan=6)

            def po_date_client(*argus):

                po_changed_date = po_selected_date.get()
                year = po_changed_date[6:10]
                date = po_changed_date[6:10] + "-"+ po_changed_date[3:5] + "-" + po_changed_date[0:2]
                po_changed_client_code = po_client_code[po_selected_client.get()]
                
                temp_req_sum_path = config['targetFolder']+'/'+po_changed_client_code+'-'+year+'/'+date+'/'+'60-Requirement-Summary'
                po_requirements_summary_path.config(text=temp_req_sum_path,wraplength=800)
        
            po_selected_client.trace('w',po_date_client)
            po_selected_date.trace('w',po_date_client)


        po_main_frame.pack(side='top',anchor=NW,fill ="x")
        # pb = ttk.Progressbar(
        #     po_log_frame,
        #     orient='horizontal',
        #     mode='determinate',
        #     length=600
        # )
        # pb['value'] = 33
        # pb.pack(padx=10,pady=5)
        po_consoleLabel = Label(po_log_frame,text='Console logs', font=labelFont)
        po_consoleLabel.pack(side='top',anchor=NW, padx=10,pady=5)
        self.console = ConsoleUi(po_log_frame)

        po_log_frame.pack(fill ="x")


class Tab2():
    def __init__(self, root,tabControl):
        pkg_client_var = StringVar()
        pkg_date_var = StringVar()

        pkg_tab = ttk.Frame(tabControl)
        tabControl.add(pkg_tab, text ='Packing Slip')

        pkg_main_frame = ttk.Frame(pkg_tab)
        pkg_log_frame = ttk.Frame(pkg_tab)

        pkg_frame = Frame(pkg_main_frame)
        pkg_frame.grid(row=0,column=0)

        pkg_heading = Label(pkg_main_frame,text='Generating Packing Slip',font=headingFont)
        pkg_heading.grid(row=0,column=0,padx=10, pady=10,sticky=W,columnspan=2) 

        self.pkg_img = ImageTk.PhotoImage(Image.open("./config/Triumph_International_Logo.png"))
        pkg_image_frame = Frame(pkg_main_frame) #,  highlightbackground="blue", highlightthickness=2, height=10,width=10
        pkg_image_frame.grid(row=0,column=2,columnspan=3)
        pkg_image_frame.place(anchor='ne', relx=0.99, rely=0.03)
        pkg_img_label = Label(pkg_image_frame, image = self.pkg_img)
        pkg_img_label.pack() 

        pkg_requirements_summary = ttk.Label(pkg_main_frame,text='Requirements Summary Path',font=labelFont)
        pkg_requirements_summary.grid(row=1,column=0,padx=10, pady=10,sticky=W)

        pkg_requirements_summary_btn = Button(pkg_main_frame,text='Select File',command=lambda:select_files(pkg_requirements_summary_path, pkg_client_var,pkg_date_var), font=buttonFont)
        pkg_requirements_summary_btn.grid(row=1,column=1,padx=10, pady=10,sticky=W)

        pkg_requirements_summary_path = Label(pkg_main_frame,text="No Path Selected", font=pathFont, wraplength=800)
        pkg_requirements_summary_path.grid(row=1,column=2,padx=10, pady=10,sticky=W,columnspan=8)

        pkg_client_Name = Label(pkg_main_frame,text='Client Name',font=labelFont)
        pkg_client_Name.grid(row=2,column=0,padx=10, pady=10,sticky=W)
        
        pkg_client_Name_value = Label(pkg_main_frame, textvariable=pkg_client_var,font=labelFont)
        pkg_client_Name_value.grid(row=2,column=1,padx=10, pady=10,sticky=W,columnspan=2)

        pkg_order_date = Label(pkg_main_frame,text='Order Date',font=labelFont)
        pkg_order_date.grid(row=2,column=3,padx=10, pady=10,sticky=W)

       
        pkg_order_date_value = Label(pkg_main_frame, textvariable=pkg_date_var,font=labelFont)
        pkg_order_date_value.grid(row=2,column=4,padx=10, pady=10,sticky=W,columnspan=3)

        pkg_thread = threading.Thread(target=lambda:begin_order_processing(mode ='packing', client=pkg_client_Name_value.cget("text"), date=pkg_order_date_value.cget("text"), path=pkg_requirements_summary_path.cget("text"), consoleLabel=pkg_consoleLabel, thread_name=pkg_thread), name='pkg_thread').start
        pkg_process_btn = Button(pkg_main_frame, command=pkg_thread, text="Process",font=buttonFont)
        pkg_process_btn.grid(row=4,column=1,padx=10, pady=10,sticky=W)

        pkg_cancel_btn = Button(pkg_main_frame, text="Cancel", font=buttonFont)
        pkg_cancel_btn.grid(row=4,column=2,padx=10, pady=10,sticky=W)

        pkg_packing_slip = Label(pkg_main_frame,text='Packing slip Folder Path ', font=labelFont)
        pkg_packing_slip.grid(row=5,column=0,padx=10, pady=10,sticky=W)

        with open(ConfigFolderPath+'config.json', 'r') as jsonFile:

            config = json.load(jsonFile)

            pkg_packing_slip_btn = Button(pkg_main_frame,text='Copy Path',command=lambda:open_folder_packaging(params=[config['targetFolder'], pkg_client_var.get(), pkg_date_var.get(), '70-Packing-Slip'],frame=pkg_main_frame),font=buttonFont)
            pkg_packing_slip_btn.grid(row=5,column=1,padx=10, pady=10,sticky=W)

            pkg_packing_slip_folder_path = Label(pkg_main_frame,text='No Path Selected', font=pathFont, wraplength=800)
            pkg_packing_slip_folder_path.grid(row=5,column=2,padx=10, pady=10,sticky=W,columnspan=8)


            def pkg_date_client(*argus):

                pkg_changed_date = pkg_date_var.get()
                
                pkg_year = pkg_changed_date[0:4]
                pkg_date = pkg_changed_date[0:4]+ "-" + pkg_changed_date[5:7] + "-"+ pkg_changed_date[8:11]
                print(pkg_changed_date,pkg_date)
                pkg_changed_client = pkg_client_var.get()

                
                with open(ClientsFolderPath, 'r') as jsonFile:
                    pkg_client_code = json.load(jsonFile)
                    pkg_client_code = pkg_client_code[pkg_changed_client]

                temp_pkg_slip_path = config['targetFolder']+'/'+pkg_client_code+'-'+pkg_year+'/'+pkg_date+'/'+'70-Packing-Slip'
                pkg_packing_slip_folder_path.config(text=temp_pkg_slip_path)

            pkg_client_var.trace('w',pkg_date_client)
            pkg_date_var.trace('w',pkg_date_client)

        pkg_main_frame.pack(side='top',anchor=NW,fill ="x")
        pkg_consoleLabel = Label(pkg_log_frame,text='Console logs', font=labelFont)
        pkg_consoleLabel.pack(side='top',anchor=NW, padx=10,pady=5)
        self.console = ConsoleUi(pkg_log_frame)

        pkg_log_frame.pack(fill ="x")





