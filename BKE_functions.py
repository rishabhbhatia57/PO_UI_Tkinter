import os
import shutil
from datetime import datetime
import pandas as pd
import xlsxwriter
import glob
import numpy as np
from openpyxl import load_workbook,Workbook
import openpyxl.utils.cell
import time
from config import ConfigFolderPath
import tabula
import csv
import json
import sys

import BKE_log
from config import MasterFolderPath

logger = BKE_log.setup_custom_logger('root')

def downloadFiles(RootFolder,POSource,OrderDate,ClientCode):
    #converting str to datetime
    # print(OrderDate)
    OrderDate = datetime.strptime(OrderDate, '%Y-%m-%d')
    # extracting year from the order date
    year = OrderDate.strftime("%Y")
    # formatting order date {2022-00-00) format
    OrderDate = OrderDate.strftime('%Y-%m-%d')
    source_folder = POSource
    destination_folder = RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/10-Download-Files/"
    try:
        for file_name in os.listdir(POSource):
            # construct full file path
            
            source = source_folder +"/"+ file_name
            destination = destination_folder
            # copy only files
            if os.path.isfile(source):

                shutil.copy(source, destination)
                logger.info("File '"+file_name+"' copied")# from source '"+source_folder+"' to destination '"+destination_folder+"'")
                print("File '"+file_name+"' copied")# from source '"+source_folder+"' to destination '"+destination_folder+"'")
    except Exception as e:
        logger.error("Error while copying files: "+str(e))
        print("Error while copying files: "+str(e))


def scriptStarted():
    logger.info('Starting script')
    print('Starting script')
    print('Starting is running...')
    print('Do not close this window while processing...')
    return "Script Started."


def scriptEnded():
    logger.info('Script Ended')
    print('Script Ended')
    return "Script Ended"
    

def checkFolderStructure(RootFolder,ClientCode,OrderDate):
    try:
        #converting str to datetime
        OrderDate = datetime.strptime(OrderDate, '%Y-%m-%d')

        # extracting year from the order date
        year = OrderDate.strftime("%Y")

        # formatting order date {2022-00-00) format
        OrderDate = OrderDate.strftime('%Y-%m-%d')
        
        # Checking if the folder exists or not, if doesnt exists, then script will create one.
        # logger.info('Checking if the folder exists or not, if doesnt exists, then script will create one.')
        # print('Checking if the folder exists or not, if doesnt exists, then script will create one.')
        logger.info("Creating the new directory...")
        print("Creating the new directory...")
        DatedPath = RootFolder +"/"+ClientCode+"-"+year+"/"+str(OrderDate)
        isExist = os.path.exists(DatedPath)
        if not isExist:
            # logger.info("Creating a new folder '"+ClientCode+"-"+year+"/"+str(OrderDate)+"' at location "+ RootFolder)
            # print("Creating a new folder '"+ClientCode+"-"+year+"/"+str(OrderDate)+"' at location "+ RootFolder)
            os.makedirs(DatedPath)
            os.makedirs(DatedPath+"/10-Download-Files")
            os.makedirs(DatedPath+"/20-Intermediate-Files")
            os.makedirs(DatedPath+"/30-Extract-CSV")
            os.makedirs(DatedPath+"/40-Extract-Excel")
            os.makedirs(DatedPath+"/50-Consolidate-Orders")
            os.makedirs(DatedPath+"/60-Requirement-Summary")
            os.makedirs(DatedPath+"/70-Packaging-Slip")
            os.makedirs(DatedPath+"/80-Logs")
            logger.info("Folder '"+ClientCode+"-"+year+"/"+str(OrderDate)+"' is created.")
            print("Folder '"+ClientCode+"-"+year+"/"+str(OrderDate)+"' is created.")
        else:
            logger.info("Folder '"+ClientCode+"-"+year+"/"+str(OrderDate)+"' exists.")
            print("Folder '"+ClientCode+"-"+year+"/"+str(OrderDate)+"' exists.")

    except Exception as e:
        logger.error("Error while checking folder structure:  "+str(e))
        print("Error while checking folder structure:  "+str(e))


def mergeExcelsToOne(RootFolder,POSource,OrderDate,ClientCode): 
    #converting str to datetime
    OrderDate = datetime.strptime(OrderDate, '%Y-%m-%d')
    # extracting year from the order date
    year = OrderDate.strftime("%Y")
    # formatting order date {2022-00-00) format
    OrderDate = OrderDate.strftime('%Y-%m-%d')
    inputpath = RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/40-Extract-Excel/"
    outputpath = RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/50-Consolidate-Orders/"

    try:
        # logger.info('Checking 40-Extract-Excel directory exists or not.')
        file_list =  glob.glob(inputpath+ "/*.xlsx")
        if len(file_list) == 0:
            logger.info('No excel files found to merge.')
            print('No excel files found to merge.')
            return
        else:
            excl_list = []
            print("Merging files...")
            for f in os.listdir(inputpath):
                # logger.info("Accessing '"+f+"' right now: ")
                # print("Accessing '"+f+"' right now: ")
                df = pd.read_excel(inputpath+"/"+f)
                df.insert(0, "file_name", f)
                excl_list.append(df)
            excl_merged = pd.concat(excl_list, ignore_index=True,)
            excl_merged.to_excel(outputpath+"/"+'Consolidate-Orders.xlsx', index=False)
            logger.info("Merged "+str(len(file_list))+ " excel files into a single excel file 'Consolidate-Orders.xlsx'")
            print("Merged "+str(len(file_list))+ " excel files into a single excel file 'Consolidate-Orders.xlsx'")
            return 'All excels are merged into a single excel file'
    except Exception as e:
        logger.info("Error while merging files: "+str(e))
        print("Error while merging files: "+str(e))


def mergeToPivotRQ(RootFolder,POSource,OrderDate,ClientCode,Formulasheet):


    with open(ConfigFolderPath+'client.json', 'r') as jsonFile:
        config = json.load(jsonFile)
        ClientName = config

        key_list = list(ClientName.keys())
        val_list = list(ClientName.values())
        position = val_list.index(ClientCode)

    try:
        #converting str to datetime
        OrderDate = datetime.strptime(OrderDate, '%Y-%m-%d')
        # extracting year from the order date
        year = OrderDate.strftime("%Y")
        # formatting order date {2022-00-00) format
        OrderDate = OrderDate.strftime('%Y-%m-%d')

        formulaWorksheet = load_workbook(Formulasheet+'/FormulaSheet.xlsx',data_only=True) 
        # Data_only = True is used to get evaluated formula value instead of formula
        formulaSheet = formulaWorksheet['FormulaSheet']
        if not os.path.exists(RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/50-Consolidate-Orders/Consolidate-Orders.xlsx"):
            logger.info("Could not find the consolidated order folder to generate requirement summary file")
            print("Could not find the consolidated order folder to generate requirement summary file")
            return
        else:
            df = pd.read_excel(RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/50-Consolidate-Orders/Consolidate-Orders.xlsx")
            df_pivot = pd.pivot_table(df, index="ArticleEAN", values='Qty', 
            columns=['Vendor Name','PO Number','Receiving Location'], aggfunc='sum')
            df_pivot['Grand Total'] = 0
            df_pivot['Closing Stock'] = 0
            df_pivot['Diff CS - GT'] = 0
            df_pivot['Rate'] = 0
            df_pivot.to_excel(RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/60-Requirement-Summary/Requirement-Summary.xlsx")


            # Adding Processing Date, Order Number and Closing Stock, Diffrence Between Grand Total and 
            # Closing Stock Field into pivot sheet for tempalte
            pivotWorksheet = load_workbook(RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/60-Requirement-Summary/Requirement-Summary.xlsx")
            pivotSheet = pivotWorksheet.active

            pivotSheet.insert_rows(1,2)
            pivotSheet.insert_rows(6,2) # IGST/CGST Type
            
            
            df = pd.DataFrame(pivotSheet, index=None)
            rows = len(df.axes[0])
            cols = len(df.axes[1])
            pivotSheet.insert_rows(rows+1) # For Grand Total At bottom of the sheet

            

            for i in range(9,rows):
                pivotSheet.cell(i,cols-3).value = "=SUM(B"+str(i)+":"+openpyxl.utils.cell.get_column_letter(cols-4)+str(i)+")"
                pivotSheet.cell(i,cols-1).value = '='+openpyxl.utils.cell.get_column_letter(cols-2)+str(i)+'-'+openpyxl.utils.cell.get_column_letter(cols-3)+str(i)
                pivotSheet.cell(i,cols-2).value = "="+formulaSheet.cell(10,2).value.replace("#VAL#",str(i)) 

            
            pivotSheet.cell(6,1).value = 'IGST/CGST Type'
            pivotSheet.cell(7,1).value = 'Order No'
            pivotSheet.cell(rows+1,1).value = 'Grand Total'
            
            VAL = ''
            for j in range(2,cols-3):
                cellValue = "="+formulaSheet.cell(2,2).value.replace("#VAL#",openpyxl.utils.cell.get_column_letter(j))
                pivotSheet.cell(6,j).value = cellValue
                pivotSheet.cell(rows+1,j).value = "=SUM("+openpyxl.utils.cell.get_column_letter(j)+str(9)+":"+openpyxl.utils.cell.get_column_letter(j)+str(rows)+")"

            todayDate = datetime.today().strftime('%Y-%m-%d')
            pivotSheet.cell(2,3).value = 'Order Date'
            pivotSheet.cell(2,4).value = OrderDate

            pivotSheet.cell(2,5).value = 'ClientName'
            pivotSheet.cell(2,6).value = key_list[position]



            pivotWorksheet.save(RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/60-Requirement-Summary/Requirement-Summary.xlsx")
            logger.info('Generated requirement summary file for order date - '+OrderDate)
            print('Generated requirement summary file for order date - '+OrderDate)
            # formulaWorksheet.save(Formulasheet+'/FormulaSheet.xlsx')

            formulaWorksheet.close()
            return 'Generated Requirement Summary file'


    except Exception as e:
        logger.error("Error while generating Requirement-Summary file: "+str(e))
        print("Error while generating Requirement-Summary file: "+str(e))


def getFilesToProcess(RootFolder,POSource,OrderDate,ClientCode):
    #converting str to datetime
    OrderDate = datetime.strptime(OrderDate, '%Y-%m-%d')
    # extracting year from the order date
    year = OrderDate.strftime("%Y")
    # formatting order date {2022-00-00) format
    OrderDate = OrderDate.strftime('%Y-%m-%d')
    try:
        inputFolderPath = RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/10-Download-Files/"
        outputFolderPath = RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/40-Extract-Excel/"
        startedProcessing = time.time()
        
        if len(os.listdir(inputFolderPath)) == 0:
            logger.info("'"+inputFolderPath+"' Folder is empty, add pdf files to convert")
            print("'"+inputFolderPath+"' Folder is empty, add pdf files to convert")
            return
        else:
            for f in os.listdir(inputFolderPath):
                fOutputExtension = f.replace('.pdf', '.xlsx')
                
                pdfToTable(inputFolderPath+f,outputFolderPath+fOutputExtension,RootFolder,POSource,OrderDate,ClientCode,f)
            
            print("Successfully converted all Files in"+"{:.2f}".format(time.time() - startedProcessing,2)+ " seconds!")
            # print("Completed in "+"{:.2f}".format(time.time() - startedProcessing,2)+ " seconds!")
    except Exception as e:
        logger.error("Error while processing files: "+str(e))
        print("Error while processing files: "+str(e))


def pdfToTable(inputPath,outputPath,RootFolder,POSource,OrderDate,ClientCode,filecsv):

    logger.info("Converting PDF files to Excel...")
    print("Converting PDF files to Excel...")
    try:
        logger.info("Converting '"+ filecsv)# +"' to excel '"+filecsv.replace('.pdf', '.xlsx')+"'")
        print("Converting '"+ filecsv)# +"' to excel '"+filecsv.replace('.pdf', '.xlsx')+"'")
        #converting str to datetime
        OrderDate = datetime.strptime(OrderDate, '%Y-%m-%d')
        # extracting year from the order date
        year = OrderDate.strftime("%Y")
        # formatting order date {2022-00-00) format
        OrderDate = OrderDate.strftime('%Y-%m-%d')

        startedProcessing = time.time()

        intermediateCSV = RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/30-Extract-CSV/"+filecsv.replace('.pdf', '.csv')
        intermediateExcel = RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/20-Intermediate-Files/1_"+filecsv.replace('.pdf', '.xlsx')
        intermediateExcel2 = RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/20-Intermediate-Files/2_"+filecsv.replace('.pdf', '.xlsx')
        intermediateoutputPath = RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/20-Intermediate-Files/3_"+filecsv.replace('.pdf', '.xlsx')

        tabula.convert_into(input_path= inputPath , output_path= intermediateCSV , pages = 'all', lattice= True)

        # making a dataframe using csv
        df = pd.read_csv(filepath_or_buffer= intermediateCSV, skiprows =[1,2,3])


        # converting csv to excel
        writer = pd.ExcelWriter(intermediateExcel, engine='xlsxwriter')
        df.to_excel(writer, sheet_name= 'Sheet1', merge_cells= False)
        writer.save()


        # loading excel
        workbook = load_workbook(filename= intermediateExcel)
        sheet = workbook.active


        # finding the last row needed
        endrow = 1
        endcolumn = 1

        for i in range(1, 10000000):
            if sheet.cell(i,1).value != None and "Grand" in sheet.cell(i,1).value:
                endrow = i
                break

        for i in range(1, 10000000):
            if sheet.cell(2,i).value != None and "Total" in sheet.cell(2,i).value:
                endcolumn = i
                break

        # print(endrow, endcolumn)
            


        # Cleaning loop
        for i in range(1, endrow+2):
            for j in range(1, endcolumn+1):
                if sheet.cell(i,j).value != None and "Aditya" in sheet.cell(i,j).value:
                    sheet.cell(i,1).value = "remove"
                elif sheet.cell(i,j).value != None and "P. O. Number" in sheet.cell(i,j).value:
                    sheet.cell(i,1).value = "remove"
                elif sheet.cell(i,j).value != None and "PO" in sheet.cell(i,j).value and i != 2:
                        sheet.cell(i,1).value = "remove"
                elif sheet.cell(i,j).value != None and "_x000D_" in sheet.cell(i,j).value:
                    temp = sheet.cell(i,j).value.replace("_x000D_", "")
                    sheet.cell(i,j).value = temp


        # rearranging the data into a single row

        # column numbers
        CGSTRate = endcolumn+1
        CGSTAmt = endcolumn+2
        SGSTRate = endcolumn+3
        SGSTAmt = endcolumn+4
        UTGSTRate = endcolumn+5
        UTGSTAmt = endcolumn+6
        IGSTRate = endcolumn+7
        IGSTAmt = endcolumn+8
        VendorPenalty = endcolumn+9
        VendorName = endcolumn+8
        ReceivingLocation = endcolumn+9
        PONumber = endcolumn+10
        VendorCode = endcolumn+11
        for i in range(1, endrow):
            for j in range(1, endcolumn+1):
                if sheet.cell(i,j).value == "CGST":
                    sheet.cell(i-1,CGSTRate).value = sheet.cell(i,j+1).value
                    sheet.cell(i-1,CGSTAmt).value = sheet.cell(i,j+2).value
                    sheet.cell(i,j+1).value = ""
                    sheet.cell(i,j+2).value = ""
                    sheet.cell(i,j).value = ""
                elif sheet.cell(i,j).value == "SGST":
                    sheet.cell(i-2,SGSTRate).value = sheet.cell(i,j+1).value
                    sheet.cell(i-2,SGSTAmt).value = sheet.cell(i,j+2).value
                    sheet.cell(i,j+1).value = ""
                    sheet.cell(i,j+2).value = ""
                    sheet.cell(i,j).value = ""
                elif sheet.cell(i,j).value == "UTGST":
                    sheet.cell(i-3,UTGSTRate).value = sheet.cell(i,j+1).value
                    sheet.cell(i-3,UTGSTAmt).value = sheet.cell(i,j+2).value
                    sheet.cell(i,j+1).value = ""
                    sheet.cell(i,j+2).value = ""
                    sheet.cell(i,j).value = ""
                elif sheet.cell(i,j).value == "IGST":
                    sheet.cell(i-4,IGSTRate).value = sheet.cell(i,j+1).value
                    sheet.cell(i-4,IGSTAmt).value = sheet.cell(i,j+2).value
                    sheet.cell(i,j+1).value = ""
                    sheet.cell(i,j+2).value = ""
                    sheet.cell(i,j).value = ""
                elif sheet.cell(i,j).value == "VendorPenalty":
                    sheet.cell(i-4,VendorPenalty).value = sheet.cell(i,j+1).value
                    sheet.cell(i,j+1).value = ""
                    sheet.cell(i,j).value = ""


        # removing empty and unwanted rows
        counter = endrow
        while counter > 0:
            if sheet.cell(counter,1).value == None or sheet.cell(counter,1).value == "" or sheet.cell(counter, 1).value == "remove":
                sheet.delete_rows(counter)
            counter = counter -1

        sheet.delete_cols(endcolumn-1)
        sheet.delete_cols(endcolumn-2)

        sheet["A1"] = "POItem"
        sheet["B1"] = "ArticleEAN"
        sheet["C1"] = "Article Number"
        sheet["D1"] = "ArticleDescription"
        sheet["E1"] = "HSNCode"
        sheet["F1"] = "MRP"
        sheet["G1"] = "BasicCostPrice(TaxableValue)"
        sheet["H1"] = "Qty"
        sheet["I1"] = "UM"
        sheet["J1"] = "TaxableValue"
        # sheet["K1"] = "GSTRate"
        # sheet["L1"] = "GSTAmt"
        sheet["K1"] = "Total Amount"
        sheet["L1"] = "CGSTRate"
        sheet["M1"] = "CGSTAmt"
        sheet["N1"] = "SGSTRate"
        sheet["O1"] = "SGSTAmt"
        sheet["P1"] = "UTGSTRate"
        sheet["Q1"] = "UTGSTAmt"
        sheet["R1"] = "IGSTRate"
        sheet["S1"] = "IGSTAmt"
        sheet["T1"] = "Vendor Penalty"
        sheet["U1"] = "Vendor Name"
        sheet["V1"] = "Receiving Location"
        sheet["W1"] = "PO Number"
        sheet["X1"] = "Vendor Code"

        # workbook = xlsxwriter.Workbook(intermediateExcel, {'strings_to_numbers': True})
        df2 = pd.read_csv(filepath_or_buffer= intermediateCSV, on_bad_lines='skip')

        # print(df2)
        writer2 = pd.ExcelWriter(intermediateExcel2, engine='xlsxwriter')
        df2.to_excel(writer2, sheet_name= 'Sheet1', merge_cells= False)
        writer2.save()
        workbook2 = load_workbook(filename= intermediateExcel2)
        sheet2 = workbook2.active

        # Cleaning loop
        for i in range(1, 10):
            for j in range(1, 5):
                if sheet2.cell(i,j).value != None and "_x000D_" in sheet2.cell(i,j).value:
                    temp = sheet2.cell(i,j).value.replace("_x000D_", "%")
                    sheet2.cell(i,j).value = temp

        # Recalculating endrow
        for i in range(1, 10000000):
            if sheet.cell(i,1).value != None and "Grand" in sheet.cell(i,1).value:
                endrow = i
                break

        PONumberRow = 1
        for i in range(1, 25):
            if sheet2.cell(i,1).value != None and "P. O. Number" in sheet2.cell(i,1).value:
                PONumberRow = i
                break

        VendorNameValue = sheet2.cell(2, 1).value.replace(":","").split('%')[1]
        VendorCodeValue = sheet2.cell(2, 1).value.replace(":","").split('%')[3]
        ReceivingLocationValue = sheet2.cell(4, 1).value.replace(":","").split('%')[3]
        PONumValue = sheet2.cell(PONumberRow, 1).value.replace(":", "").split('%')[1]
        for i in range(2, endrow):
            sheet.cell(i, VendorName).value = VendorNameValue
            sheet.cell(i, ReceivingLocation).value = ReceivingLocationValue
            sheet.cell(i, PONumber).value = PONumValue
            sheet.cell(i, VendorCode).value = VendorCodeValue
        
        # get  Delete row count
        lastrow = 1
        for i in range(endrow, 1000000):
            if sheet.cell(i,1).value != None and "Other Conditions" in sheet.cell(i,1).value:
                lastrow = i+5
                break
            
        while endrow <= lastrow:
            sheet.delete_rows(lastrow)
            lastrow = lastrow -1

        workbook2.save(filename= intermediateoutputPath)
        workbook.save(filename= outputPath)

        # logger.info("Converted '"+ inputPath + "' to '" + outputPath+"'"+ " in "+ "{:.2f}".format(time.time() - startedProcessing,2)+ " seconds.")
        # print("Converted '"+ inputPath + "' to '" + outputPath+"'"+ " in "+ "{:.2f}".format(time.time() - startedProcessing,2)+ " seconds.")
        print("Converted '"+ inputPath + " in "+ "{:.2f}".format(time.time() - startedProcessing,2)+ " seconds.")
        
        return "Conversion Complete!"
    
    except Exception as e:
        logger.error("Error while processing file: "+str(e))
        print("Error while processing file: "+str(e))


def generatingPackaingSlip(RootFolder,ReqSource,OrderDate,ClientCode,Formulasheet,TemplateFiles):
    try:
        #converting str to datetime
        OrderDate = datetime.strptime(OrderDate, '%Y-%m-%d')
        # extracting year from the order date
        year = OrderDate.strftime("%Y")
        # formatting order date {2022-00-00) format
        OrderDate = OrderDate.strftime('%Y-%m-%d')

        startedTemplating = time.time()
        sourcePivot = ReqSource
        source = TemplateFiles+"/PackingSlip-Template.xlsx"
        destination = TemplateFiles+"/TemplateFile.xlsx"
        

        # Making Copy of template file
        # shutil.copy(source, destination)
        # print("File copied successfully.")

        # Load work vook and sheets
        InputWorkbook = load_workbook(sourcePivot,data_only=True)
        # TemplateWorkbook = load_workbook(destination)

        InputSheet = InputWorkbook.active
        # TemplateSheet = TemplateWorkbook['ORDER']

        # Get rows and Column count
        df = pd.DataFrame(InputSheet, index=None)
        rows = len(df.axes[0])
        cols = len(df.axes[1])


        formulaWorksheet = load_workbook(Formulasheet+'/FormulaSheet.xlsx',data_only=True)
        formulaSheet = formulaWorksheet['FormulaSheet']
        DBFformula = formulaWorksheet['DBF']

        # df_itemMaster = pd.read_excel(MasterFolderPath+'Item Master.xlsx',sheet_name='Item Master')
        # df_IGST = pd.read_excel(MasterFolderPath+'IGST Master.xlsx',sheet_name='DBF')
        # df_SGST = pd.read_excel(MasterFolderPath+'SGST Master.xlsx',sheet_name='DBF')

        # Opening Item Master Sheet
        df_IteamMaster = pd.read_excel(MasterFolderPath+'Item Master.xlsx', sheet_name='Item Master',index_col=False)
        # Opening IGST Master Sheet
        df_IGSTMaster = pd.read_excel(MasterFolderPath+'IGST Master.xlsx', sheet_name='DBF',index_col=False)
        # Opening SGST Master Sheet
        df_SGSTMaster = pd.read_excel(MasterFolderPath+'SGST Master.xlsx', sheet_name='DBF',index_col=False)
        # Opening Location2 Master Sheet
        df_Location2 = pd.read_excel(MasterFolderPath+'Location 2 Master.xlsx',sheet_name='Location2',index_col=False)
        
        
        
        for column in range(2,cols-3):
            startedTemplatingFile = time.time()
            # Making Copy of template file
            shutil.copy(source, destination)
            # logger.info("Template File copied successfully for generating packaging-slip")
            # print("Template File copied successfully for generating packaging-slip")

            # Load work vook and sheets
            TemplateWorkbook = load_workbook(destination, data_only=True)
            TemplateSheet = TemplateWorkbook['ORDER']
            dbfsheet = TemplateWorkbook['DBF']

            
            TemplateSheet.cell(6,4).value = InputSheet.cell(7,2).value # Order Name
            TemplateSheet.cell(5,1).value = InputSheet.cell(7,2).value # Order Name
            # PO Number
            filename = InputSheet.cell(4,column).value
            TemplateSheet.cell(5,2).value = InputSheet.cell(4,column).value
            TemplateSheet.cell(6,2).value = InputSheet.cell(4,column).value
            TemplateSheet.cell(6,1).value = InputSheet.cell(4,column).value
            TemplateSheet.cell(6,3).value = InputSheet.cell(4,column).value
            # Receving Location
            TemplateSheet.cell(5,3).value = InputSheet.cell(5,column).value
            TemplateSheet.cell(4,2).value = InputSheet.cell(5,column).value

            TemplateSheet.cell(1,1).value = 'DATE'
            TemplateSheet.cell(1,2).value = InputSheet.cell(2,4).value # Date

            TemplateSheet.cell(1,3).value = 'SGST/IGST'
            TemplateSheet.cell(1,4).value = InputSheet.cell(6,column).value # IGST/SGST Type
            if TemplateSheet.cell(1,4).value == None:
                print("IGST/SGST TYPE = None, Requirment Summary file is not saved. Open the file, save it then process")
                break
            

            # Copy EAN to template sheet
            Trows = 8
            Tcols = 5
            dbfrows = 2
            dbfcols = 57
            for row in range(7,rows):
                # if InputSheet.cell(row,column).value != None or InputSheet.cell(row,column).value != "":
                if str(InputSheet.cell(row,column).value).isnumeric():
                    
                        
                    # Copy Qty to template sheet
                    TemplateSheet.cell(Trows,Tcols).value = InputSheet.cell(row,column).value 
                    TemplateSheet.cell(Trows,Tcols+1).value = InputSheet.cell(row,column).value
                    TemplateSheet.cell(Trows,Tcols+2).value = "="+openpyxl.utils.cell.get_column_letter(Tcols+1)+str(Trows) # Actual Qty
                    
                    # Copy EAN to template sheet
                    TemplateSheet.cell(Trows,Tcols-3).value = InputSheet.cell(row,1).value

                    # VLOOKUP
                    # StyleName
                    # TemplateSheet.cell(Trows,Tcols-4).value = "="+formulaSheet.cell(3,2).value.replace("#VAL#",str(Trows))
                    TemplateSheet.cell(Trows,Tcols-4).value = "='Hidden Item Master'!A"+str(Trows-6)


                    # style
                    # TemplateSheet.cell(Trows,Tcols-2).value =  "="+formulaSheet.cell(4,2).value.replace("#VAL#",str(Trows))
                    TemplateSheet.cell(Trows,Tcols-2).value = "='Hidden Item Master'!C"+str(Trows-6)

                    # SADM SKU
                    # TemplateSheet.cell(Trows,Tcols-1).value = "="+formulaSheet.cell(5,2).value.replace("#VAL#",str(Trows))
                    TemplateSheet.cell(Trows,Tcols-1).value = "='Hidden Item Master'!D"+str(Trows-6)
                    
                    # Rate (in Rs.) Order file
                    # TemplateSheet.cell(Trows,Tcols+3).value = "="+formulaSheet.cell(6,2).value.replace("#VAL#",str(Trows))
                    TemplateSheet.cell(Trows,Tcols+3).value = "='Hidden Item Master'!E"+str(Trows-6)

                    #Cls stk vs order
                    # TemplateSheet.cell(Trows,Tcols+6).value = TemplateSheet.cell(Trows,Tcols+5).value - TemplateSheet.cell(Trows,Tcols+2).value
                    TemplateSheet.cell(Trows,Tcols+6).value = "="+openpyxl.utils.cell.get_column_letter(Tcols+5)+str(Trows) +'-'+openpyxl.utils.cell.get_column_letter(Tcols+2)+str(Trows) 

                    # LOCATION2
                    # TemplateSheet.cell(Trows,Tcols+7).value = "="+formulaSheet.cell(7,2).value.replace("#VAL#",str(Trows))
                    TemplateSheet.cell(Trows,Tcols+7).value = "='Hidden Item Master'!F"+str(Trows-6)


                    #BULK  / DTA  BULK  /  EOSS LOC
                    # TemplateSheet.cell(Trows,Tcols+8).value = "="+formulaSheet.cell(8,2).value.replace("#VAL#",str(Trows))
                    TemplateSheet.cell(Trows,Tcols+8).value = "='Hidden Item Master'!G"+str(Trows-6)

                    #MRP
                    # TemplateSheet.cell(Trows,Tcols+9).value = "="+formulaSheet.cell(9,2).value.replace("#VAL#",str(Trows))
                    TemplateSheet.cell(Trows,Tcols+9).value = "='Hidden Item Master'!E"+str(Trows-6)

                    # Closing stk
                    # TemplateSheet.cell(Trows,Tcols+5).value = "="+formulaSheet.cell(11,2).value.replace("#VAL#",str(Trows)) 
                    TemplateSheet.cell(Trows,Tcols+5).value = InputSheet.cell(row,cols-2).value

                    #SCAN
                    # TemplateSheet.cell(Trows,Tcols+10).value = "="+formulaSheet.cell(12,2).value.replace("#VAL#",str(Trows))
                    TemplateSheet.cell(Trows,Tcols+10).value = "=SUMIF('packing slip'!$B$3:$B$999,$B"+str(Trows)+",'packing slip'!$C$3:$C$999)"

                    #SCAN VS DIFF
                    # TemplateSheet.cell(Trows,Tcols+11).value = "="+formulaSheet.cell(13,2).value.replace("#VAL#",str(Trows))
                    TemplateSheet.cell(Trows,Tcols+11).value = "=O"+str(Trows)+"-G"+str(Trows)

                    # ERROR
                    # TemplateSheet.cell(Trows,Tcols+12).value = "="+formulaSheet.cell(14,2).value.replace("#VAL#",str(Trows))
                    TemplateSheet.cell(Trows,Tcols+12).value = "=IF(P"+str(Trows)+"<>0,1,0)"

                    # STYLE COLOR
                    # TemplateSheet.cell(Trows,Tcols+13).value = "="+formulaSheet.cell(15,2).value.replace("#VAL#",str(Trows))
                    TemplateSheet.cell(Trows,Tcols+13).value = "=IF(P"+str(Trows)+"<=0,0,1)"
                    
                    # Adding values to DBF
                    for i in range(1,dbfcols):
                        dbfsheet.cell(dbfrows, i).value = '='+DBFformula.cell(2,i+1).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows))
                    
                    Trows += 1
                    dbfrows += 1 

            
            TemplateWorkbook.save(RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/70-Packaging-Slip/"+"PackagingSlip_"+str(filename)+".xlsx")
            TemplateWorkbook.close()


            # Opening Packaging slip using openpyxl to check igst/sgst value
            sourcePackagingSlip = RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/70-Packaging-Slip/"+"PackagingSlip_"+str(filename)+".xlsx"
            TemplateWorkbook = load_workbook(sourcePackagingSlip, data_only=True)
            TemplateSheet = TemplateWorkbook['ORDER']
            # print(TemplateSheet.cell(1,4).value)
            # Getting data from IGST/ SGST Sheet DBF to Tempalate sheet DBF

            # If IGST then open IGST Master
            if TemplateSheet.cell(1,4).value == 'IGST':

                # Opening Packaging slip as df for second time to get EAN values
                df_TempalateWorkbook = pd.read_excel(sourcePackagingSlip,sheet_name='ORDER', skiprows=6,index_col=False)
                df_TempalateWorkbook.rename(columns = {'EAN':'EAN ID'}, inplace = True)
                
                # Temporary df to store EAN ID
                df_EAN_temp = df_TempalateWorkbook[['EAN ID']]
                # Applying join using EAN ID to get hidden_item_master to get SKU and Other fields
                df_hidden_item_master = df_EAN_temp.merge(df_IteamMaster,on='EAN ID',how='left')

                df_Location2_temp = df_hidden_item_master.merge(df_Location2,on='EAN ID',how='left')
                # print(df_Location2_temp)

                # Temporary df to store SKU
                df_SKU_temp = df_hidden_item_master[['SKU ID']]
                df_SKU_temp.rename(columns = {'SKU ID':'ITEMNAME'}, inplace = True)
                # Applying join using ITEMNAME to get hidden_dbf (from IGST/SGST sheet) to get SKU and Other fields
                df_GST_hidden = df_SKU_temp.merge(df_IGSTMaster,on='ITEMNAME',how='left')


                with pd.ExcelWriter(sourcePackagingSlip, mode='a', engine='openpyxl',if_sheet_exists='replace') as writer:
                    # Adding Hidden Item Master sheet to the RQ sheet with values 'Style Name', 'EAN', 'Style', 'SKU ID', 'MRP'
                    df_Location2_temp.to_excel(writer,sheet_name='Hidden Item Master',index=False, columns=['Style Name', 'EAN', 'Style', 'SKU', 'MRP','Location 2','BULK  / DTA  BULK  /  EOSS LOC'])
                    # # Adding Hidden DBF sheet to the RQ sheet with values 'Vouchertypename', 'CSNNO','DATE' ETC.
                    df_GST_hidden.to_excel(writer,sheet_name='Hidden DBF',index=False, columns=['Vouchertypename', 'CSNNO','DATE',
                    'REFERENCE', 'REF1','DEALNAME', 'PRICELEVEL', 'ITEMNAME', 'GODOWN', 'QTY', 'RATE', 'SUBTOTAL', 'DISCPERC',
                    'DISCAMT', 'ITEMVALUE', 'LedgerAcct', 'CATEGORY1', 'COSTCENT1', 'CATEGORY2', 'COSTCENT2', 'CATEGORY3', 'COSTCENT3',
                    'CATEGORY4', 'COSTCENT4', 'ITEMTOTAL', 'TOTALQTY', 'CDISCHEAD', 'CDISCPERC', 'COMMONDISC', 'BEFORETAX',
                    'TAXHEAD', 'TAXPERC', 'TAXAMT', 'STAXHEAD', 'STAXPERC', 'STAXAMT', 'ITAXHEAD', 'ITAXPERC' ,'ITAXAMT', 'NETAMT',
                    'ROUND', 'ROUND1', 'REFTYPE', 'Name', 'REFAMT', 'Narration', 'Transport','transmode', 'pymtterm', 'ordno',
                    'orddate', 'DANO', 'Delyadd1', 'Delyadd2', 'Delyadd3', 'Delyadd4'])
            
            # If SGST then open SGST Master
            if TemplateSheet.cell(1,4).value == 'SGST':
                
                # Opening Packaging slip as df for second time to get EAN values
                df_TempalateWorkbook = pd.read_excel(sourcePackagingSlip,sheet_name='ORDER', skiprows=6,index_col=False)
                df_TempalateWorkbook.rename(columns = {'EAN':'EAN ID'}, inplace = True)
                
                # Temporary df to store EAN ID
                df_EAN_temp = df_TempalateWorkbook[['EAN ID']]
                # Applying join using EAN ID to get hidden_item_master to get SKU and Other fields
                df_hidden_item_master = df_EAN_temp.merge(df_IteamMaster,on='EAN ID',how='left')

                df_Location2_temp = df_hidden_item_master.merge(df_Location2,on='EAN ID',how='left')
                # print(df_Location2_temp)

                # Temporary df to store SKU
                df_SKU_temp = df_hidden_item_master[['SKU ID']]
                df_SKU_temp.rename(columns = {'SKU ID':'ITEMNAME'}, inplace = True)
                # Applying join using ITEMNAME to get hidden_dbf (from IGST/SGST sheet) to get SKU and Other fields
                df_GST_hidden = df_SKU_temp.merge(df_SGSTMaster,on='ITEMNAME',how='left')


                with pd.ExcelWriter(sourcePackagingSlip, mode='a', engine='openpyxl',if_sheet_exists='replace') as writer:
                    # Adding Hidden Item Master sheet to the RQ sheet with values 'Style Name', 'EAN', 'Style', 'SKU ID', 'MRP'
                    df_Location2_temp.to_excel(writer,sheet_name='Hidden Item Master',index=False, columns=['Style Name', 'EAN', 'Style', 'SKU', 'MRP','Location 2','BULK  / DTA  BULK  /  EOSS LOC'])
                    # # Adding Hidden DBF sheet to the RQ sheet with values 'Vouchertypename', 'CSNNO','DATE' ETC.
                    df_GST_hidden.to_excel(writer,sheet_name='Hidden DBF',index=False, columns=['Vouchertypename', 'CSNNO','DATE',
                    'REFERENCE', 'REF1','DEALNAME', 'PRICELEVEL', 'ITEMNAME', 'GODOWN', 'QTY', 'RATE', 'SUBTOTAL', 'DISCPERC',
                    'DISCAMT', 'ITEMVALUE', 'LedgerAcct', 'CATEGORY1', 'COSTCENT1', 'CATEGORY2', 'COSTCENT2', 'CATEGORY3', 'COSTCENT3',
                    'CATEGORY4', 'COSTCENT4', 'ITEMTOTAL', 'TOTALQTY', 'CDISCHEAD', 'CDISCPERC', 'COMMONDISC', 'BEFORETAX',
                    'TAXHEAD', 'TAXPERC', 'TAXAMT', 'STAXHEAD', 'STAXPERC', 'STAXAMT', 'ITAXHEAD', 'ITAXPERC' ,'ITAXAMT', 'NETAMT',
                    'ROUND', 'ROUND1', 'REFTYPE', 'Name', 'REFAMT', 'Narration', 'Transport','transmode', 'pymtterm', 'ordno',
                    'orddate', 'DANO', 'Delyadd1', 'Delyadd2', 'Delyadd3', 'Delyadd4'])
                pass

            TemplateWorkbook.close()


            logger.info("Packing slip generated for: "+str(filename)+ " file in {:.2f}".format(time.time() - startedTemplatingFile,2)+ " seconds.")
            print("Packing slip generated for: "+str(filename)+ " file in {:.2f}".format(time.time() - startedTemplatingFile,2)+ " seconds.")
            
        logger.info("Total time taken for generation of packing-slips:  {:.2f}".format(time.time() - startedTemplating,2)+ " seconds.")
        print("Total time taken for generation of packing-slips:  {:.2f}".format(time.time() - startedTemplating,2)+ " seconds.")
        return 'Completed!'
    except Exception as e:
        logger.error("Error while generating packing-slip file: "+str(e))
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)
        print("Error while generating packing-slip file: "+str(e))