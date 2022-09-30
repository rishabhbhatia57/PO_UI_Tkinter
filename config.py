import tkinter.font as tkFont
import tkinter as tk                    
from tkinter import ttk
from tkinter import *
import json


# Path used to get config path, client path and app theme
ConfigFolderPath = "./config/"

with open(ConfigFolderPath+'config.json', 'r') as jsonFile:
  config = json.load(jsonFile)
  MasterFolderPath = config['masterFolder']
  ClientsFolderPath = config['clients']

# Fonts used throughout application
headingFont = ("Calibri",30)
fieldFont =  ("Calibri",15)
buttonFont = ("Calibri",15)
labelFont = ("Calibri",15)
pathFont = ("Calibri",12)
logFont = ("Calibri",15)
# menuFont = tkFont.Font(family='Calibri', size=15)
# menuOptionsFont = tkFont.Font(family='Calibri', size=15)