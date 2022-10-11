import tkinter as tk
import sys
import logging
import tkinter.scrolledtext as st
from tkinter import ttk
import webbrowser
from tkinter.ttk import *
from tkinter import *
from tkinter.font import Font

import queue
import BKE_log
from config import logFont

logger = BKE_log.setup_custom_logger('root')


class QueueHandler(logging.Handler):
    """Class to send logging records to a queue

    It can be used from different threads
    """

    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        self.log_queue.put(record)


class ConsoleUi:
    """Poll messages from a logging queue and display them in a scrolled text widget"""
    
    def __init__(self, frame):
        
        self.frame = frame
        # def callback(url):
        #     webbrowser.open_new_tab(url)
        # footer_frame = ttk.Frame(frame)
        # inside_footer_frame = ttk.Frame(footer_frame)
        # inside_footer_frame.grid(row=0,column=0)

        # Developed_by = Label(inside_footer_frame, text="Developed by   -  ",font=Font(size=10,weight="bold"), cursor="hand2")
        # Developed_by.grid(row=0,column=1)

        # Developed_text = Label(inside_footer_frame, text="C-BIA Solutions & Services LLP",font=Font(size=10), cursor="hand2")
        # Developed_text.grid(row=0,column=2)

        # Website = Label(inside_footer_frame, text="          Website  -  ",font=Font(size=10,weight="bold"), cursor="hand2")
        # Website.grid(row=0,column=3)

        # link = Label(inside_footer_frame, text="https://c-bia.com/" ,font=Font(size=10, underline=1), cursor="hand2") #foreground='#EA0920'
        # link.grid(row=0,column=4)
        # link.bind("<Button-1>", lambda e:
        # callback("https://c-bia.com/"))
        # footer_frame.pack()
        # https://docs.python.org/3/library/tkinter.html#threading-model
        
        # Create a ScrolledText wdiget
        self.scrolled_text = st.ScrolledText(frame, state='disabled')#height=20
        self.scrolled_text.pack(expand = True,fill ="x",ipady=20,ipadx=10)
        self.scrolled_text.configure(font=logFont)
        self.scrolled_text.tag_config('INFO', foreground='White')
        self.scrolled_text.tag_config('DEBUG', foreground='gray')
        self.scrolled_text.tag_config('WARNING', foreground='orange')
        self.scrolled_text.tag_config('ERROR', foreground='red')
        self.scrolled_text.tag_config('CRITICAL', foreground='red', underline=1)
        # Create a logging handler using a queue
        self.log_queue = queue.Queue()
        self.queue_handler = QueueHandler(self.log_queue)
        formatter = logging.Formatter('%(asctime)s %(levelname)s %(message)s')
        self.queue_handler.setFormatter(formatter)
        logger.addHandler(self.queue_handler)
        # Start polling messages from the queue
        self.frame.after(100, self.poll_log_queue)

    def display(self, record):
        msg = self.queue_handler.format(record)
        self.scrolled_text.configure(state='normal')
        self.scrolled_text.insert(tk.END, msg + '\n', record.levelname)
        self.scrolled_text.configure(state='disabled')
        # Autoscroll to the bottom
        self.scrolled_text.yview(tk.END)

    def poll_log_queue(self):
        # Check every 100ms if there is a new message in the queue to display
        while True:
            try:
                record = self.log_queue.get(block=False)
            except queue.Empty:
                break
            else:
                self.display(record)
        self.frame.after(100, self.poll_log_queue)
