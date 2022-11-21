import tkinter.font as tkFont
import tkinter as tk                    
from tkinter import ttk
from tkinter import *
import json
import BKE_log

logger = BKE_log.setup_custom_logger('root')

# Path used to get config path, client path and app theme
ConfigFolderPath = "./config/config.json"
PO_CLIENTS = []
PKG_CLIENTS = []

with open(ConfigFolderPath, 'r') as jsonFile:
  config = json.load(jsonFile)
  ITEMMASTERPATH = config['itemMasterPath']
  IGSTMASTERPATH = config['igstMasterPath']
  SGSTMASTERPATH = config['sgstMasterPath']
  LOCATIONMASTERPATH = config['locationMasterPath']
  LOCATION2MASTERPATH = config['location2MasterPath']
  CLOSINGSTOCKMASTERPATH = config['closingStockMasterPath']
  # MasterFolderPath = config['masterFolder']
  PACKINGSLIPTEMPLATEPATH = config['packingSlipTemplatePath']
  REQSUMTEMPLATEPATH = config['reqSumTemplatePath']
  FORMULASHEETPATH = config['formulaSheetPath']
  DESTINATIONPATH = config['targetFolder']
  TEMPLATESPATH = config['templateFolder']
  CLIENTSFOLDERPATH = config['clients']
  ICONIMAGE = config['iconImage']
  LOGOIMAGE = config['logoImage']

  if "orderNo_to_5_digit_format" in config:
    ORDERN0TO5DIGIT = config['orderNo_to_5_digit_format']
  else:
    ORDERN0TO5DIGIT = 'N'


def get_client_code(clientname):
    clientCode = ''
    with open(CLIENTSFOLDERPATH, 'r') as jsonFile:
        config = json.load(jsonFile)
        for i in range(len(config['Clients'])):
            if config['Clients'][i]['ClientName'] == clientname:
                clientCode = config['Clients'][i]['ClientCode']
                break
    if clientCode == '':
        logger.error('Could not fetch the client code. Make sure the client selected exists.')
    return clientCode

def get_client_name(clientcode):
    clientName = ''
    with open(CLIENTSFOLDERPATH, 'r') as jsonFile:
        config = json.load(jsonFile)
        for i in range(len(config['Clients'])):
            if config['Clients'][i]['ClientCode'] == clientcode:
                clientName = config['Clients'][i]['ClientName']
                break
    if clientName == '':
        logger.error('Could not fetch the client Name. Make sure the client selected exists and has proper Client code.')
    return clientName




def seprate_clients():
  try:
    with open(CLIENTSFOLDERPATH, 'r') as jsonFile:
      clients_json = json.load(jsonFile)
    for i in range(len(clients_json['Clients'])):
      if clients_json['Clients'][i]['PDFImport'] == 'Y':
        PO_CLIENTS.append(clients_json['Clients'][i]['ClientName'])
        PKG_CLIENTS.append(clients_json['Clients'][i]['ClientName'])
      else:
        PKG_CLIENTS.append(clients_json['Clients'][i]['ClientName'])
    # print('po_clients: '+str(PO_CLIENTS), 'pkg_clients: '+str(PKG_CLIENTS)) 
  
  except Exception as e:
    print('Seprating Clients: '+str(e))


# seprating clients on the basis of the PDFImport value
seprate_clients()

# Fonts used throughout application
headingFont = ("Calibri",30)
fieldFont =  ("Calibri",15)
buttonFont = ("Calibri",15)
labelFont = ("Calibri",15)
pathFont = ("Calibri",12)
logFont = ("Consolas",10)
# menuFont = tkFont.Font(family='Calibri', size=15)
# menuOptionsFont = tkFont.Font(family='Calibri', size=15)