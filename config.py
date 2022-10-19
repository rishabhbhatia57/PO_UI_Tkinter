import tkinter.font as tkFont
import tkinter as tk                    
from tkinter import ttk
from tkinter import *
import json


# Path used to get config path, client path and app theme
ConfigFolderPath = "./config/config.json"

with open(ConfigFolderPath, 'r') as jsonFile:
  config = json.load(jsonFile)
  itemMasterPath = config['itemMasterPath']
  igstMasterPath = config['igstMasterPath']
  sgstMasterPath = config['sgstMasterPath']
  locationMasterPath = config['locationMasterPath']
  location2MasterPath = config['location2MasterPath']
  closingStockMasterPath = config['closingStockMasterPath']
  # MasterFolderPath = config['masterFolder']
  ClientsFolderPath = config['clients']

# Fonts used throughout application
headingFont = ("Calibri",30)
fieldFont =  ("Calibri",15)
buttonFont = ("Calibri",15)
labelFont = ("Calibri",15)
pathFont = ("Calibri",12)
logFont = ("Consolas",10)
# menuFont = tkFont.Font(family='Calibri', size=15)
# menuOptionsFont = tkFont.Font(family='Calibri', size=15)