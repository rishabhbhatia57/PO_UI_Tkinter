from flask import Flask, jsonify, request
import json
from flask_cors import CORS
import os
from datetime import datetime
import sys
import json
import base64
from tkinter.messagebox import showinfo
from config import ConfigFolderPath, CLIENTSFOLDERPATH, ITEMMASTERPATH, IGSTMASTERPATH, SGSTMASTERPATH, LOCATIONMASTERPATH, LOCATION2MASTERPATH, CLOSINGSTOCKMASTERPATH, PACKINGSLIPTEMPLATEPATH, REQSUMTEMPLATEPATH, FORMULASHEETPATH, DESTINATIONPATH, TEMPLATESPATH
import BKE_log
from BKE_functions import scriptStarted, downloadFiles, scriptEnded, checkFolderStructure, mergeExcelsToOne,mergeToPivotRQ, generatingPackingSlip, po_check_master_files, pkg_check_master_files
from pdf_to_excel import getFilesToProcess


logger = BKE_log.setup_custom_logger('root')


def startProcessing(mode,clientname,orderdate,processing_source):
    try:
        isExistConfigFolderPath = os.path.exists(ConfigFolderPath)
        # isExistMasterFolderPath = os.path.exists(MasterFolderPath)
        isExistitemMasterPath = os.path.exists(IGSTMASTERPATH)
        isExistigstMasterPath = os.path.exists(IGSTMASTERPATH)
        isExistsgstMasterPath = os.path.exists(SGSTMASTERPATH)
        isExistlocationMasterPath = os.path.exists(LOCATIONMASTERPATH)
        isExistlocation2MasterPath = os.path.exists(LOCATION2MASTERPATH)
        isExistclosingStockMasterPath = os.path.exists(CLOSINGSTOCKMASTERPATH)

        if isExistConfigFolderPath == False and isExistitemMasterPath== False and isExistigstMasterPath== False and isExistsgstMasterPath== False and isExistlocationMasterPath== False and isExistlocation2MasterPath== False and isExistclosingStockMasterPath== False  :
            showinfo(
            title='Invalid Selection',
            message='Can not find path to Config and Master Files'
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
                # print(DESTINATIONPATH,clientcode,orderdate,'consolidation')

                checkmasterfiles = po_check_master_files(formulaWorksheet=FORMULASHEETPATH)
                if checkmasterfiles['valid'] == True:
                    # converting str to datetime
                    OrderDate = datetime.strptime(orderdate, '%Y-%m-%d')
                    # extracting year from the order date
                    year = OrderDate.strftime("%Y")
                    # formatting order date {2022-00-00) format
                    OrderDate = OrderDate.strftime('%Y-%m-%d')

                    base_path = DESTINATIONPATH + '/' + clientcode + '-' + year + '/' + str(OrderDate) 
                    checkFolderStructure(RootFolder=DESTINATIONPATH,ClientCode=clientcode,OrderDate=orderdate,mode = 'consolidation', base_path=base_path)
                    # 3. To download PDF Files from Google Drive and Store it in week/DownloadFiles Folder
                    downloadFiles(RootFolder=DESTINATIONPATH,POSource=processing_source,OrderDate=orderdate,ClientCode=clientcode, base_path=base_path) # Done
                    # 4. Converted PDF files to Excel Files, perform Cleaning, and store to week/uploadFiles Folder
                    getFilesToProcess(RootFolder=DESTINATIONPATH,POSource=processing_source,OrderDate=orderdate,ClientCode=clientcode, base_path=base_path)
                    # 5. Merge all the coverted excel file to a single excel file and store in week/MergeExcelsFiles folder
                    mergeExcelsToOne(RootFolder=DESTINATIONPATH,POSource=processing_source,OrderDate=orderdate,ClientCode=clientcode, base_path=base_path)
                    # 6. PivotTable - Template Creation
                    mergeToPivotRQ(RootFolder=DESTINATIONPATH,POSource=processing_source,OrderDate=orderdate,ClientCode=clientcode,formulaWorksheet=FORMULASHEETPATH, TemplateFiles=TEMPLATESPATH, base_path=base_path, reqSumTemplatePath=REQSUMTEMPLATEPATH)

                    scriptEnded()
                else:
                    scriptEnded()

            if mode == 'packing':
                clientcode = clientname
                # print(clientname)
                # with open(CLIENTSFOLDERPATH, 'r') as jsonFile:
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
                # converting str to datetime
                OrderDate = datetime.strptime(orderdate, '%Y-%m-%d')
                # extracting year from the order date
                year = OrderDate.strftime("%Y")
                # formatting order date {2022-00-00) format
                OrderDate = OrderDate.strftime('%Y-%m-%d')

                base_path = DESTINATIONPATH + '/' + clientcode + '-' + year + '/' + str(OrderDate)
                checkmasterfiles = pkg_check_master_files(formulaWorksheet=FORMULASHEETPATH)
                if checkmasterfiles['valid'] == True:
                    checkFolderStructure(RootFolder=DESTINATIONPATH,ClientCode=clientcode,OrderDate=orderdate,mode = 'packing', base_path=base_path)
                    generatingPackingSlip(RootFolder=DESTINATIONPATH,ReqSource=processing_source,OrderDate=orderdate,ClientCode=clientname,formulaWorksheet=FORMULASHEETPATH,TemplateFiles=TEMPLATESPATH, base_path=base_path, packingSlipTemplatePath=PACKINGSLIPTEMPLATEPATH)
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
