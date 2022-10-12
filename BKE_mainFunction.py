from flask import Flask, jsonify, request
import json
from flask_cors import CORS
import os
from datetime import datetime
import sys
import json
import base64
from tkinter.messagebox import showinfo
from config import ConfigFolderPath, MasterFolderPath, ClientsFolderPath
import BKE_log
from BKE_functions import scriptStarted, downloadFiles, scriptEnded, checkFolderStructure, mergeExcelsToOne,mergeToPivotRQ, generatingPackingSlip, pdfToTable,getFilesToProcess, check_master_files


with open(ConfigFolderPath+'config.json', 'r') as jsonFile:
    config = json.load(jsonFile)
    formulasheetpath = config['formulaFolder']
    masterspath = config['masterFolder']
    # configpath = config['formulaFolder']
    templatespath = config['templateFolder']
    destinationpath = config['targetFolder']


logger = BKE_log.setup_custom_logger('root')


def startProcessing(mode,clientname,orderdate,processing_source):
    try:
        isExistConfigFolderPath = os.path.exists(ConfigFolderPath)
        isExistMasterFolderPath = os.path.exists(MasterFolderPath)

        if isExistConfigFolderPath == False and isExistMasterFolderPath== False:
            showinfo(
            title='Invalid Selection',
            message='Can not find path to Config and Master Folder'
        )
        else:
            print("Mode: ",mode)
            print("Client Code: ",clientname)
            print("Order date: ", orderdate)
            print("Source Path: ", processing_source)
            # Phase I
            if mode == 'consolidation':
                clientcode = clientname
                logger.info("Client Name: "+clientname+" Client Code: "+clientcode+" Order Date: "+orderdate+" PO Folder Path: '"+processing_source+"'" )
                # print(mode)
                # logger.info(mode)
                # 1. Notify that the script is Started
                scriptStarted()
                # 2. Checking the folder structure 
                # print(destinationpath,clientcode,orderdate,'consolidation')

                checkmasterfiles = check_master_files(RootFolder=destinationpath,OrderDate=orderdate,ClientCode=clientcode,formulaWorksheet=formulasheetpath, TemplateFiles=templatespath)
                if checkmasterfiles['valid'] == True:
                    logger.info('All master files are validated!')
                    checkFolderStructure(RootFolder=destinationpath,ClientCode=clientcode,OrderDate=orderdate,mode = 'consolidation')
                    # 3. To download PDF Files from Google Drive and Store it in week/DownloadFiles Folder
                    downloadFiles(RootFolder=destinationpath,POSource=processing_source,OrderDate=orderdate,ClientCode=clientcode) # Done
                    # 4. Converted PDF files to Excel Files, perform Cleaning, and store to week/uploadFiles Folder
                    getFilesToProcess(RootFolder=destinationpath,POSource=processing_source,OrderDate=orderdate,ClientCode=clientcode)
                    # 5. Merge all the coverted excel file to a single excel file and store in week/MergeExcelsFiles folder
                    mergeExcelsToOne(RootFolder=destinationpath,POSource=processing_source,OrderDate=orderdate,ClientCode=clientcode)
                    # 6. PivotTable - Template Creation
                    mergeToPivotRQ(RootFolder=destinationpath,POSource=processing_source,OrderDate=orderdate,ClientCode=clientcode,formulaWorksheet=formulasheetpath, TemplateFiles=templatespath)

                    scriptEnded()
                else:
                    scriptEnded()

            if mode == 'packing':
                clientcode = clientname
                # print(clientname)
                # with open(ClientsFolderPath, 'r') as jsonFile:
                #     config = json.load(jsonFile)
                #     clientNameDict = config
                #     key_list = list(clientNameDict.keys())
                #     val_list = list(clientNameDict.values())
                #     position = val_list.index(clientname)
                #     clientcode = key_list[position]
                # print(clientname, clientcode)
                # logger.info("Client Name: "+clientname+"\nClient Code: "+clientname+"\nOrder Date: "+orderdate+"\nPO Folder Path: '"+processing_source+"'" )
                # print("Client Name: "+clientname+" Client Code: "+clientname+" Order Date: "+orderdate+" PO Folder Path: '"+processing_source+"'")
            # Phase II
                scriptStarted()
                checkmasterfiles = check_master_files(RootFolder=destinationpath,OrderDate=orderdate,ClientCode=clientcode,formulaWorksheet=formulasheetpath, TemplateFiles=templatespath)
                if checkmasterfiles['valid'] == True:
                    checkFolderStructure(RootFolder=destinationpath,ClientCode=clientcode,OrderDate=orderdate,mode = 'packing')
                    generatingPackingSlip(RootFolder=destinationpath,ReqSource=processing_source,OrderDate=orderdate,ClientCode=clientname,formulaWorksheet=formulasheetpath,TemplateFiles=templatespath)
                # 7. Notify that the script is Ended
                    scriptEnded()
                else:
                    scriptEnded()

    except Exception as e:
        logger.error("Exception: "+ str(e))
        print("Exception: "+ str(e))
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)
