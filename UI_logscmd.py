import tkinter as tk
import sys
import logging

# class PrintLogger(): # create file like object
#     def __init__(self, textbox): # pass reference to text widget
#         self.textbox = textbox # keep ref

#     def write(self, text):
#         # text = sys.stdout.read()
#         # print(text)
#         self.textbox.insert(tk.END, text) # write text to textbox
#             # could also scroll to end of textbox here to make sure always visible

#     def flush(self): # needed for file like object
#         pass


class TextHandler(logging.Handler):
    
    """This class allows you to log to a Tkinter Text or ScrolledText widget"""
    def __init__(self, text):
        # run the regular Handler __init__
        logging.Handler.__init__(self)
        # Store a reference to the Text it will log to
        self.text = text

    def emit(self, record):
        msg = self.format(record)
        def append():
            self.text.configure(state='normal')
            self.text.insert(Tkinter.END, msg + '\n')
            self.text.configure(state='disabled')
            # Autoscroll to the bottom
            self.text.yview(Tkinter.END)
        # This is necessary because we can't modify the Text from other threads
        self.text.after(0, append)