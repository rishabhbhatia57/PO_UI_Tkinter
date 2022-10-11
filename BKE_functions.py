import os
import shutil
from datetime import datetime
import pandas as pd
import xlsxwriter
import glob
import numpy as np
from openpyxl import load_workbook, Workbook
import openpyxl.utils.cell
import time
from config import ConfigFolderPath, ClientsFolderPath
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from openpyxl.utils import get_column_letter
# from openpyxl.styles import Style
import tabula
import csv
import json
import sys
from openpyxl.styles import PatternFill
from flask import jsonify

import BKE_log
# from UI_logscmd import PrintLogger
from config import MasterFolderPath
pd.options.mode.chained_assignment = None
# import warnings
# warnings.simplefilter(action='ignore', category=SettingWithCopyWarning)

logger = BKE_log.setup_custom_logger('root')


def validate_column_names(df_formula_cols, df_master, file_name):
    try:

        df_master_columns = pd.DataFrame(
            list(df_master.columns), columns=['Column Names'])
        df_validate_columns = df_formula_cols[df_formula_cols['Master Files'] == file_name]
        df_temp = df_validate_columns.assign(Check=df_validate_columns['Column Names'].isin
                                             (df_master_columns['Column Names']).astype(str))

        df_temp = df_temp[df_temp['Check'].str.contains('False')]

        if not df_temp.empty:
            cols_not_found_list = list(df_temp['Column Names'])
            results = {
                "cols_name": cols_not_found_list,
                "cols_not_found": True
            }
            logger.error('Process Stopped! Unable to find '+str(results['cols_name'])+' columns in '+str(
                file_name)+' file. Kindly check '+str(file_name)+'.')
            return results
        else:
            print('Validated '+str(file_name)+'.')
            results = {
                "cols_name": "",
                "cols_not_found": False
            }
            return results
    except Exception as e:
        print(e)
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)


def check_master_files(RootFolder, OrderDate, ClientCode, formulaWorksheet, TemplateFiles):

    valid_result = {
        'valid': True
    }

    invalid_result = {
        'valid': False
    }
    try:
        logger.info('Validating master files. This may take few minutes...')
        df_formula_cols = pd.read_excel(
            formulaWorksheet, sheet_name="Validate Column Names")

        with open(ConfigFolderPath+'config.json', 'r') as jsonFile:
            config = json.load(jsonFile)

            file_name = "Item Master"
            df_item_master = pd.read_excel(
                config['masterFolder']+'Item Master.xlsx', sheet_name='Item Master')
            # checking  Item Master 1
            result = validate_column_names(
                df_formula_cols, df_item_master, file_name)
            if result['cols_not_found'] == True:  # Problem exists
                return invalid_result

             # checking  Location Master 2
            file_name = "Location Master"
            df_location_master = pd.read_excel(
                config['masterFolder']+file_name+'.xlsx', sheet_name='Location Master')
            result1 = validate_column_names(
                df_formula_cols, df_location_master, file_name)
            if result1['cols_not_found'] == True:
                return invalid_result

            # checking  Location Master 3
            file_name = "WH Closing Stock"
            df_closing_stock = pd.read_excel(
                config['masterFolder']+'WH Closing Stock.xlsx', sheet_name='ClosingStock', skiprows=3)  # ClosingStock
            result2 = validate_column_names(
                df_formula_cols, df_closing_stock, file_name)
            if result2['cols_not_found'] == True:
                return invalid_result

            # checking  IGST Master 4
            file_name = "IGST Master"
            df_igst_master = pd.read_excel(
                config['masterFolder']+'IGST Master.xlsx', sheet_name='DBF')
            result3 = validate_column_names(
                df_formula_cols, df_igst_master, file_name)
            if result3['cols_not_found'] == True:
                return invalid_result

            # checking  SGST Master 5
            file_name = "SGST Master"
            df_sgst_master = pd.read_excel(
                config['masterFolder']+'SGST Master.xlsx', sheet_name='DBF')
            result4 = validate_column_names(
                df_formula_cols, df_sgst_master, file_name)
            if result4['cols_not_found'] == True:
                return invalid_result

            # checking  Location 2 Master 6
            file_name = "Location 2 Master"
            df_location_2_master = pd.read_excel(
                config['masterFolder']+'Location 2 Master.xlsx', sheet_name='Location2')
            result5 = validate_column_names(
                df_formula_cols, df_location_2_master, file_name)
            if result5['cols_not_found'] == True:
                return invalid_result

        return valid_result

    except Exception as e:
        print(e)
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)
        logger.error("Error while checking "+str(file_name)+" file: "+str(e))
        return invalid_result


def downloadFiles(RootFolder, POSource, OrderDate, ClientCode):
    # file_logger = BKE_log.setup_custom_logger_file('root',RootFolder,OrderDate,ClientCode)
    # converting str to datetime
    # print(OrderDate)
    OrderDate = datetime.strptime(OrderDate, '%Y-%m-%d')
    # extracting year from the order date
    year = OrderDate.strftime("%Y")
    # formatting order date {2022-00-00) format
    OrderDate = OrderDate.strftime('%Y-%m-%d')
    source_folder = POSource
    destination_folder = RootFolder+"/"+ClientCode+"-" + \
        year+"/"+OrderDate+"/99-Working/10-Download-Files/"
    try:
        # print("Copying PDF Files from '"+str(source_folder)+"' to '"+str(destination_folder)+"'")
        logger.info("Copying PDF Files from '"+str(source_folder) +
                    "' to '"+str(destination_folder)+"'")
        # file_logger.info("Copying PDF Files from '"+str(source_folder)+"' to '"+str(destination_folder)+"'")
        # Tab1.pl.write("Copying PDF Files from '"+str(source_folder)+"' to '"+str(destination_folder)+"'")

        for file_name in os.listdir(POSource):
            # construct full file path

            source = source_folder + "/" + file_name
            destination = destination_folder
            # copy only files
            if os.path.isfile(source):

                shutil.copy(source, destination)
                # from source '"+source_folder+"' to destination '"+destination_folder+"'")
                logger.info("File '"+file_name+"' copied")
                # file_logger.info("File '"+file_name+"' copied")
                # from source '"+source_folder+"' to destination '"+destination_folder+"'")
                print("File '"+file_name+"' copied")
    except Exception as e:
        logger.error("Error while copying files: "+str(e))
        # file_logger.info("Error while copying files: "+str(e))
        print("Error while copying files: "+str(e))


def scriptStarted():
    logger.info('Starting script')
    print('Starting script')
    print('Script is running...')
    print('Do not close this window while processing...')
    logger.info('Do not close this window while processing...')
    return "Script Started."


def scriptEnded():
    logger.info('Script Ended')
    print('Script Ended')
    return "Script Ended"


def checkFolderStructure(RootFolder, ClientCode, OrderDate, mode):

    try:
        # converting str to datetime
        OrderDate = datetime.strptime(OrderDate, '%Y-%m-%d')
        # extracting year from the order date
        year = OrderDate.strftime("%Y")
        # formatting order date {2022-00-00) format
        OrderDate = OrderDate.strftime('%Y-%m-%d')

        # Checking if the folder exists or not, if doesnt exists, then script will create one.
        # logger.info('Checking if the folder exists or not, if doesnt exists, then script will create one.')
        # print('Checking if the folder exists or not, if doesnt exists, then script will create one.')

        DatedPath = RootFolder + "/"+ClientCode+"-"+year+"/"+str(OrderDate)
        isExist = os.path.exists(DatedPath)

        internalDir = ["/99-Working/10-Download-Files", "/99-Working/20-Intermediate-Files",
                       "/99-Working/30-Extract-CSV", "/99-Working/40-Extract-Excel",
                       "/50-Consolidate-Orders", "/60-Requirement-Summary", "/70-Packing-Slip", "/80-Logs"]

        if mode == 'consolidation':
            if not isExist:
                logger.info("Creating the new directory...")
                # file_logger.info("Creating the new directory...")
                os.makedirs(DatedPath)
                for i in range(0, len(internalDir)):
                    if not os.path.exists(DatedPath+internalDir[i]):
                        os.makedirs(DatedPath+internalDir[i])

            if isExist:
                for i in range(0, len(internalDir)):
                    if not os.path.exists(DatedPath+internalDir[i]):
                        os.makedirs(DatedPath+internalDir[i])

        if mode == 'packing':
            if not isExist:
                logger.info("Creating the new directory...")
                # file_logger.info("Creating the new directory...")
                os.makedirs(DatedPath)
                for i in range(6, len(internalDir)):
                    if not os.path.exists(DatedPath+internalDir[i]):
                        os.makedirs(DatedPath+internalDir[i])
                # os.makedirs(DatedPath+"/70-Packing-Slip")
                # os.makedirs(DatedPath+"/80-Logs")
            if isExist:
                logger.info("Creating the new directory...")
                # file_logger.info("Creating the new directory...")
                print("Creating the new directory...")
                for i in range(6, len(internalDir)):
                    if not os.path.exists(DatedPath+internalDir[i]):
                        os.makedirs(DatedPath+internalDir[i])
                # os.makedirs(DatedPath+"/70-Packing-Slip")
                # if not log_isExist:
                #     os.makedirs(DatedPath+"/80-Logs")

            logger.info("Folder '"+ClientCode+"-"+year +
                        "/"+str(OrderDate)+"' is created.")
            # # file_logger.info("Folder '"+ClientCode+"-"+year+"/"+str(OrderDate)+"' is created.")
            print("Folder '"+ClientCode+"-"+year +
                  "/"+str(OrderDate)+"' is created.")
        else:
            logger.info("Folder '"+ClientCode+"-"+year +
                        "/"+str(OrderDate)+"' exists.")
            # # file_logger.info("Folder '"+ClientCode+"-"+year+"/"+str(OrderDate)+"' exists.")
            print("Folder '"+ClientCode+"-"+year +
                  "/"+str(OrderDate)+"' exists.")
        # file_logger = BKE_log.setup_custom_logger_file('root',RootFolder,OrderDate,ClientCode)

    except Exception as e:
        logger.error("Error while checking folder structure:  "+str(e))
        # file_logger.info("Folder '"+ClientCode+"-"+year+"/"+str(OrderDate)+"' exists.")
        print("Error while checking folder structure:  "+str(e))


def mergeExcelsToOne(RootFolder, POSource, OrderDate, ClientCode):
    # file_logger = BKE_log.setup_custom_logger_file('root',RootFolder,OrderDate,ClientCode)
    # converting str to datetime
    OrderDate = datetime.strptime(OrderDate, '%Y-%m-%d')
    # extracting year from the order date
    year = OrderDate.strftime("%Y")
    # formatting order date {2022-00-00) format
    OrderDate = OrderDate.strftime('%Y-%m-%d')
    inputpath = RootFolder+"/"+ClientCode+"-"+year + \
        "/"+OrderDate+"/99-Working/40-Extract-Excel/"
    outputpath = RootFolder+"/"+ClientCode+"-" + \
        year+"/"+OrderDate+"/50-Consolidate-Orders/"

    try:
        # logger.info('Checking 40-Extract-Excel directory exists or not.')
        file_list = glob.glob(inputpath + "/*.xlsx")
        if len(file_list) == 0:
            logger.info('No excel files found to merge.')
            # file_logger.info('No excel files found to merge.')
            print('No excel files found to merge.')
            return
        else:
            excl_list = []
            print("Merging files...")
            logger.info("Merging files...")
            # file_logger.info("Merging files...")
            for f in os.listdir(inputpath):
                # logger.info("Accessing '"+f+"' right now: ")
                # print("Accessing '"+f+"' right now: ")
                df = pd.read_excel(inputpath+"/"+f)
                df.insert(0, "file_name", f)
                excl_list.append(df)
            excl_merged = pd.concat(excl_list, ignore_index=True,)
            excl_merged.to_excel(
                outputpath+"/"+'Consolidate-Orders.xlsx', index=False)
            logger.info("Merged "+str(len(file_list)) +
                        " excel files into a single excel file 'Consolidate-Orders.xlsx'")
            # file_logger.info("Merged "+str(len(file_list))+ " excel files into a single excel file 'Consolidate-Orders.xlsx'")
            print("Merged "+str(len(file_list)) +
                  " excel files into a single excel file 'Consolidate-Orders.xlsx'")
            return 'All excels are merged into a single excel file'
    except Exception as e:
        logger.info("Error while merging files: "+str(e))
        # file_logger.info("Error while merging files: "+str(e))
        print("Error while merging files: "+str(e))


def mergeToPivotRQ(RootFolder, POSource, OrderDate, ClientCode, formulaWorksheet, TemplateFiles):
    with open(ClientsFolderPath, 'r') as jsonFile:
        config = json.load(jsonFile)
        ClientName = config

        key_list = list(ClientName.keys())
        val_list = list(ClientName.values())
        position = val_list.index(ClientCode)

    try:
        # converting str to datetime
        OrderDate = datetime.strptime(OrderDate, '%Y-%m-%d')
        # extracting year from the order date
        year = OrderDate.strftime("%Y")
        # formatting order date {2022-00-00) format
        OrderDate = OrderDate.strftime('%Y-%m-%d')

        # formulaWorksheet1 = load_workbook(formulaWorksheet,data_only=True)
        # formulaSheet = formulaWorksheet['FormulaSheet']
        df_formula_cols = pd.read_excel(
            formulaWorksheet, sheet_name="Validate Column Names")

        if not os.path.exists(RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/50-Consolidate-Orders/Consolidate-Orders.xlsx"):
            logger.info(
                "Could not find the consolidated order folder to generate requirement summary file")
            # file_logger.info("Could not find the consolidated order folder to generate requirement summary file")
            print(
                "Could not find the consolidated order folder to generate requirement summary file")
            return
        else:

            df_consolidated_order = pd.read_excel(
                RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/50-Consolidate-Orders/Consolidate-Orders.xlsx")

            with open(ConfigFolderPath+'config.json', 'r') as jsonFile:
                config = json.load(jsonFile)
                df_item_master = pd.read_excel(
                    config['masterFolder']+'Item Master.xlsx')
                df_location_master = pd.read_excel(
                    config['masterFolder']+'Location Master.xlsx')
                df_closing_stock = pd.read_excel(
                    config['masterFolder']+'WH Closing Stock.xlsx', skiprows=3)

            # --------------------
            # Renaming EAN as Article number to perform join using ArticleEAN
            df_item_master.rename(columns={'EAN': 'ArticleEAN'}, inplace=True)

            # common_columns = df_validate_item_master.merge(df_item_master, on= )

            # Merge on EAN from Item master
            df_SKU = df_consolidated_order.merge(
                df_item_master, on='ArticleEAN', how='left')

            df_SKU.rename(columns={'MRP_x': 'MRP'}, inplace=True)
            df_SKU_nodups = df_SKU.drop_duplicates()  # dropping duplicates

            df_SKU_nodups.to_excel(RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/50-Consolidate-Orders/df_join_SKU.xlsx", columns=[
                                   'ArticleEAN', 'SKU', 'Qty', 'MRP', 'Receiving Location', 'Style', 'Style Name', 'PO Number'], index=False)
            # Opening df_SKU excel as df_SKU dataframe
            df_SKU = pd.read_excel(RootFolder+"/"+ClientCode+"-"+year +
                                   "/"+OrderDate+"/50-Consolidate-Orders/df_join_SKU.xlsx")

            # Perfoming join to get values of GST
            df_gst_type = df_SKU.merge(
                df_location_master, on='Receiving Location', how='left')
            df_gst_type_nodups = df_gst_type.drop_duplicates()  # dropping duplicates
            df_gst_type_nodups.to_excel(RootFolder+"/"+ClientCode+"-"+year+"/" +
                                        OrderDate+"/50-Consolidate-Orders/df_join_gst_type.xlsx", index=False)
            # Opening df_gst_type excel as df_gst_type dataframe
            df_gst_type = pd.read_excel(
                RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/50-Consolidate-Orders/df_join_gst_type.xlsx")

            # Perfoming join to get values of closing stock
            df_join_cl_stk = df_gst_type.merge(
                df_closing_stock, on='SKU', how='left')

            df_join_cl_stk_nodups = df_join_cl_stk.drop_duplicates()  # dropping duplicates

            df_join_cl_stk_nodups['Actual qty'] = df_join_cl_stk_nodups['Actual qty'].fillna(
                0)

            df_join_cl_stk_nodups.rename(
                columns={'Style_x': 'Style', 'Style Name_x': 'Style Name'}, inplace=True)

            df_join_cl_stk_nodups.to_excel(RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/50-Consolidate-Orders/df_join_cl_stk.xlsx", index=False, columns=['ArticleEAN', 'SKU', 'Qty', 'MRP',
                                                                                                                                                                'Receiving Location', 'Style', 'Style Name',	'PO Number', 'SGST/IGST Type', 'Actual qty'])

            # Opening df_gst_type excel as df_gst_type dataframe
            df_join_cl_stk = pd.read_excel(
                RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/50-Consolidate-Orders/df_join_cl_stk.xlsx")

            df_join_cl_stk['Order No.'] = ''  # adding order number as col
            df_join_cl_stk['Grand Total'] = ''  # adding Grand Total as col

            # final file used by requirement summary to make pivot
            df_join_cl_stk.to_excel(RootFolder+"/"+ClientCode+"-"+year+"/" +
                                    OrderDate+"/50-Consolidate-Orders/df_join_pivot.xlsx", index=False)
            df_pivot_final_join = pd.read_excel(
                RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/50-Consolidate-Orders/df_join_pivot.xlsx")
            # --------------------

            # Constant Variables used in loops
            workbook_path = RootFolder+"/"+ClientCode+"-"+year+"/" + \
                OrderDate+"/60-Requirement-Summary/Requirement-Summary.xlsx"
            workbook_sheet = 'Requirement Summary'
            color = "00FFCC99"
            start_rows = 3
            start_cols = 8

            df_pivot = pd.pivot_table(df_join_cl_stk, index=["ArticleEAN", 'Actual qty', 'MRP', "SKU"], values='Qty',
                                      columns=['PO Number', 'Order No.', 'Grand Total', 'SGST/IGST Type', 'Receiving Location'], aggfunc='sum')

            df_pivot['Grand Total'] = 0
            df_pivot['Closing Stock'] = 0
            df_pivot['Diff CS - GT'] = 0
            df_pivot['Rate'] = 0

            df_pivot.to_excel(workbook_path, sheet_name=workbook_sheet)

            # open pivot sheet again
            df_temp_p = pd.read_excel(workbook_path, sheet_name=workbook_sheet)
            df_temp = pd.read_excel(
                workbook_path, sheet_name=workbook_sheet, skiprows=5)
            df_temp.to_excel(RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/50-Consolidate-Orders/"+"df_temp.xlsx",
                             columns=['MRP', 'Actual qty'], index=False)
            tempWorkbook = load_workbook(
                RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/50-Consolidate-Orders/"+"df_temp.xlsx")
            tempSheet = tempWorkbook.active
            # Inserting 5 rows to handle the gap between closing stock header and starting(entry) rows
            tempSheet.insert_rows(2, 5)
            # Saving closing stock and MRP in this temp sheet for later use
            tempWorkbook.save(RootFolder+"/"+ClientCode+"-"+year +
                              "/"+OrderDate+"/50-Consolidate-Orders/"+"df_temp.xlsx")

            df_temp = pd.read_excel(RootFolder+"/"+ClientCode+"-"+year +
                                    "/"+OrderDate+"/50-Consolidate-Orders/"+"df_temp.xlsx")
            # Copying Closing stock from temp file to requirement summary file
            df_temp_p['Closing Stock'] = df_temp['Actual qty']
            print('Fetching Rate values from Item Master...')
            # Copying MRP from temp file to requirement summary file
            df_temp_p['Rate'] = df_temp['MRP']

            rq_template_source = TemplateFiles+"Requirement-Summary-Template.xlsx"
            rq_summary_destination = RootFolder+"/"+ClientCode+"-"+year+"/" + \
                OrderDate+"/60-Requirement-Summary/Temp-Requirement-Summary.xlsx"

            shutil.copy(rq_template_source, rq_summary_destination)

            with pd.ExcelWriter(rq_summary_destination, mode='a', engine='openpyxl', if_sheet_exists='replace') as rq_sum_writer:
                # Saving RQ Sum after adding CS, MRP
                df_temp_p.to_excel(
                    rq_sum_writer, sheet_name=workbook_sheet, index=False)

            pivotWorksheet = load_workbook(rq_summary_destination)
            pivotSheet = pivotWorksheet[workbook_sheet]
            # Deleting the cols for Actual Oty and MRP_y
            pivotSheet.delete_cols(2, 2)
            # inerting rows for date and cient name
            pivotSheet.insert_rows(1, 1)

            pivotSheet.cell(1, 1).value = 'Order Date'
            pivotSheet.cell(1, 1).font = Font(bold=True)
            pivotSheet.cell(1, 2).value = OrderDate
            pivotSheet.cell(2, 1).value = ''  # Removing Unnamed: 0

            pivotSheet.cell(1, 3).value = 'ClientName'
            pivotSheet.cell(1, 3).font = Font(bold=True)
            pivotSheet.cell(1, 4).value = key_list[position]

            rows = pivotSheet.max_row  # get max rows
            cols = pivotSheet.max_column  # get max rows

            for j in range(start_rows, cols-3):  # Rows Grand total
                pivotSheet.cell(start_rows+1, j).value = "=SUM("+openpyxl.utils.cell.get_column_letter(j)+str(
                    start_cols)+":"+openpyxl.utils.cell.get_column_letter(j)+str(rows)+")"  # Grand Total on second last col
                pivotSheet.cell(start_rows, j).fill = PatternFill(
                    start_color=color, end_color=color, fill_type="solid")  # Color to order field

            for i in range(start_cols, rows+1):  # Cols Grand total
                pivotSheet.cell(i, cols-3).value = "=SUM(B"+str(i)+":" + \
                    openpyxl.utils.cell.get_column_letter(cols-4)+str(i)+")"
                pivotSheet.cell(i, cols-1).value = "="+openpyxl.utils.cell.get_column_letter(
                    cols-2)+str(i)+"-"+openpyxl.utils.cell.get_column_letter(cols-3)+str(i)

            pivotSheet.cell(3, 2).font = Font(bold=True)
            pivotSheet.cell(4, 2).font = Font(bold=True)
            pivotSheet.cell(5, 2).font = Font(bold=True)
            pivotSheet.cell(6, 2).font = Font(bold=True)
            pivotSheet.cell(7, 2).font = Font(bold=True)
            pivotSheet.cell(7, 1).font = Font(bold=True)

            thin_border = Border(left=Side(style='thin'),
                                 right=Side(style='thin'),
                                 top=Side(style='thin'),
                                 bottom=Side(style='thin'))

            for i in range(1, rows+1):
                for j in range(1, cols+1):
                    pivotSheet.cell(row=i, column=j).border = thin_border
                    pivotSheet.cell(row=i, column=j).alignment = Alignment(
                        horizontal='center', vertical='center')

            dim_holder = DimensionHolder(worksheet=pivotSheet)
            for col in range(pivotSheet.min_column, pivotSheet.max_column + 1):
                dim_holder[get_column_letter(col)] = ColumnDimension(
                    pivotSheet, min=col, max=col, width=20)
            pivotSheet.column_dimensions = dim_holder

            for r in range(8, rows+1):
                pivotSheet[f'A{r}'].number_format = '0'

            pivotWorksheet.save(workbook_path)
            os.remove(rq_summary_destination)

            print('Requirements summary sheet generated.')
            logger.info('Requirements summary sheet generated.')
            # file_logger.info('Requirements summary sheet generated.')

            return 'Generated Requirement Summary file'

    except Exception as e:
        logger.error(
            "Error while generating Requirement-Summary file: "+str(e))
        # file_logger.error("Error while generating Requirement-Summary file: "+str(e))
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)
        print("Error while generating Requirement-Summary file: "+str(e))


def getFilesToProcess(RootFolder, POSource, OrderDate, ClientCode):
    # file_logger = BKE_log.setup_custom_logger_file('root',RootFolder,OrderDate,ClientCode)
    # converting str to datetime
    OrderDate = datetime.strptime(OrderDate, '%Y-%m-%d')
    # extracting year from the order date
    year = OrderDate.strftime("%Y")
    # formatting order date {2022-00-00) format
    OrderDate = OrderDate.strftime('%Y-%m-%d')
    try:
        inputFolderPath = RootFolder+"/"+ClientCode+"-" + \
            year+"/"+OrderDate+"/99-Working/10-Download-Files/"
        outputFolderPath = RootFolder+"/"+ClientCode+"-" + \
            year+"/"+OrderDate+"/99-Working/40-Extract-Excel/"
        startedProcessing = time.time()

        if len(os.listdir(inputFolderPath)) == 0:
            logger.info("'"+inputFolderPath +
                        "' Folder is empty, add pdf files to convert")
            # file_logger.info("'"+inputFolderPath+"' Folder is empty, add pdf files to convert")
            print("'"+inputFolderPath+"' Folder is empty, add pdf files to convert")
            return
        else:
            logger.info("Converting PDF files to Excel...")
            # file_logger.info("Converting PDF files to Excel...")
            print("Converting PDF files to Excel...")
            count = 0

            for f in os.listdir(inputFolderPath):
                fOutputExtension = f.replace('.pdf', '.xlsx')
                pdfToTable(inputFolderPath+f, outputFolderPath+fOutputExtension,
                           RootFolder, POSource, OrderDate, ClientCode, f)
                count += 1

            print("Successfully converted "+str(count)+" Files in " +
                  "{:.2f}".format(time.time() - startedProcessing, 2) + " seconds!")
            logger.info("Successfully converted "+str(count)+" Files in " +
                        "{:.2f}".format(time.time() - startedProcessing, 2) + " seconds!")
            # file_logger.info("Successfully converted "+str(count)+" Files in "+"{:.2f}".format(time.time() - startedProcessing,2)+ " seconds!")
            # print("Completed in "+"{:.2f}".format(time.time() - startedProcessing,2)+ " seconds!")
    except Exception as e:
        logger.error("Error while processing files: "+str(e))
        # file_logger.error("Error while processing files: "+str(e))
        print("Error while processing files: "+str(e))


def pdfToTable(inputPath, outputPath, RootFolder, POSource, OrderDate, ClientCode, filecsv):
    # file_logger = BKE_log.setup_custom_logger_file('root',RootFolder,OrderDate,ClientCode)

    try:
        # logger.info("Converting '"+ filecsv)# +"' to excel '"+filecsv.replace('.pdf', '.xlsx')+"'")
        # file_logger.info("Converting '"+ filecsv)
        # print("Converting '"+ filecsv)# +"' to excel '"+filecsv.replace('.pdf', '.xlsx')+"'")
        # converting str to datetimedatetime
        OrderDate = datetime.strptime(OrderDate, '%Y-%m-%d')
        # extracting year from the order date
        year = OrderDate.strftime("%Y")
        # formatting order date {2022-00-00) format
        OrderDate = OrderDate.strftime('%Y-%m-%d')

        startedProcessing = time.time()

        intermediateCSV = RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate + \
            "/99-Working/30-Extract-CSV/"+filecsv.replace('.pdf', '.csv')
        intermediateExcel = RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate + \
            "/99-Working/20-Intermediate-Files/1_" + \
            filecsv.replace('.pdf', '.xlsx')
        intermediateExcel2 = RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate + \
            "/99-Working/20-Intermediate-Files/2_" + \
            filecsv.replace('.pdf', '.xlsx')
        intermediateoutputPath = RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate + \
            "/99-Working/20-Intermediate-Files/3_" + \
            filecsv.replace('.pdf', '.xlsx')

        tabula.convert_into(
            input_path=inputPath, output_path=intermediateCSV, pages='all', lattice=True)

        # making a dataframe using csv
        df = pd.read_csv(filepath_or_buffer=intermediateCSV,
                         skiprows=[1, 2, 3])

        # converting csv to excel
        writer = pd.ExcelWriter(intermediateExcel, engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1', merge_cells=False)
        writer.save()

        # loading excel
        workbook = load_workbook(filename=intermediateExcel)
        sheet = workbook.active

        # finding the last row needed
        endrow = 1
        endcolumn = 1

        for i in range(1, 10000000):
            if sheet.cell(i, 1).value != None and "Grand" in sheet.cell(i, 1).value:
                endrow = i
                break

        for i in range(1, 10000000):
            if sheet.cell(2, i).value != None and "Total" in sheet.cell(2, i).value:
                endcolumn = i
                break

        # print(endrow, endcolumn)

        # Cleaning loop
        for i in range(1, endrow+2):
            for j in range(1, endcolumn+1):
                if sheet.cell(i, j).value != None and "Aditya" in sheet.cell(i, j).value:
                    sheet.cell(i, 1).value = "remove"
                elif sheet.cell(i, j).value != None and "P. O. Number" in sheet.cell(i, j).value:
                    sheet.cell(i, 1).value = "remove"
                elif sheet.cell(i, j).value != None and "PO" in sheet.cell(i, j).value and i != 2:
                    sheet.cell(i, 1).value = "remove"
                elif sheet.cell(i, j).value != None and "_x000D_" in sheet.cell(i, j).value:
                    temp = sheet.cell(i, j).value.replace("_x000D_", "")
                    sheet.cell(i, j).value = temp

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
                if sheet.cell(i, j).value == "CGST":
                    sheet.cell(i-1, CGSTRate).value = sheet.cell(i, j+1).value
                    sheet.cell(i-1, CGSTAmt).value = sheet.cell(i, j+2).value
                    sheet.cell(i, j+1).value = ""
                    sheet.cell(i, j+2).value = ""
                    sheet.cell(i, j).value = ""
                elif sheet.cell(i, j).value == "SGST":
                    sheet.cell(i-2, SGSTRate).value = sheet.cell(i, j+1).value
                    sheet.cell(i-2, SGSTAmt).value = sheet.cell(i, j+2).value
                    sheet.cell(i, j+1).value = ""
                    sheet.cell(i, j+2).value = ""
                    sheet.cell(i, j).value = ""
                elif sheet.cell(i, j).value == "UTGST":
                    sheet.cell(i-3, UTGSTRate).value = sheet.cell(i, j+1).value
                    sheet.cell(i-3, UTGSTAmt).value = sheet.cell(i, j+2).value
                    sheet.cell(i, j+1).value = ""
                    sheet.cell(i, j+2).value = ""
                    sheet.cell(i, j).value = ""
                elif sheet.cell(i, j).value == "IGST":
                    sheet.cell(i-4, IGSTRate).value = sheet.cell(i, j+1).value
                    sheet.cell(i-4, IGSTAmt).value = sheet.cell(i, j+2).value
                    sheet.cell(i, j+1).value = ""
                    sheet.cell(i, j+2).value = ""
                    sheet.cell(i, j).value = ""
                elif sheet.cell(i, j).value == "VendorPenalty":
                    sheet.cell(
                        i-4, VendorPenalty).value = sheet.cell(i, j+1).value
                    sheet.cell(i, j+1).value = ""
                    sheet.cell(i, j).value = ""

        # removing empty and unwanted rows
        counter = endrow
        while counter > 0:
            if sheet.cell(counter, 1).value == None or sheet.cell(counter, 1).value == "" or sheet.cell(counter, 1).value == "remove":
                sheet.delete_rows(counter)
            counter = counter - 1

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
        df2 = pd.read_csv(filepath_or_buffer=intermediateCSV,
                          on_bad_lines='skip')

        # print(df2)
        writer2 = pd.ExcelWriter(intermediateExcel2, engine='xlsxwriter')
        df2.to_excel(writer2, sheet_name='Sheet1', merge_cells=False)
        writer2.save()
        workbook2 = load_workbook(filename=intermediateExcel2)
        sheet2 = workbook2.active

        # Cleaning loop
        for i in range(1, 10):
            for j in range(1, 5):
                if sheet2.cell(i, j).value != None and "_x000D_" in sheet2.cell(i, j).value:
                    temp = sheet2.cell(i, j).value.replace("_x000D_", "%")
                    sheet2.cell(i, j).value = temp

        # Recalculating endrow
        for i in range(1, 10000000):
            if sheet.cell(i, 1).value != None and "Grand" in sheet.cell(i, 1).value:
                endrow = i
                break

        PONumberRow = 1
        for i in range(1, 25):
            if sheet2.cell(i, 1).value != None and "P. O. Number" in sheet2.cell(i, 1).value:
                PONumberRow = i
                break

        VendorNameValue = sheet2.cell(
            2, 1).value.replace(":", "").split('%')[1]
        VendorCodeValue = sheet2.cell(
            2, 1).value.replace(":", "").split('%')[3]
        ReceivingLocationValue = sheet2.cell(
            4, 1).value.replace(":", "").split('%')[3]
        PONumValue = sheet2.cell(PONumberRow, 1).value.replace(
            ":", "").split('%')[1]
        for i in range(2, endrow):
            sheet.cell(i, VendorName).value = VendorNameValue
            sheet.cell(i, ReceivingLocation).value = ReceivingLocationValue
            sheet.cell(i, PONumber).value = PONumValue
            sheet.cell(i, VendorCode).value = VendorCodeValue

        # get  Delete row count
        lastrow = 1
        for i in range(endrow, 1000000):
            if sheet.cell(i, 1).value != None and "Other Conditions" in sheet.cell(i, 1).value:
                lastrow = i+5
                break

        while endrow <= lastrow:
            sheet.delete_rows(lastrow)
            lastrow = lastrow - 1

        workbook2.save(filename=intermediateoutputPath)
        workbook.save(filename=outputPath)

        # logger.info("Converted '"+ inputPath + "' to '" + outputPath+"'"+ " in "+ "{:.2f}".format(time.time() - startedProcessing,2)+ " seconds.")
        # print("Converted '"+ inputPath + "' to '" + outputPath+"'"+ " in "+ "{:.2f}".format(time.time() - startedProcessing,2)+ " seconds.")
        print("Converted '" + filecsv + "' to '"+filecsv.replace('pdf', 'xlsx') +
              "' in " + "{:.2f}".format(time.time() - startedProcessing, 2) + " seconds.")
        logger.info("Converted '" + filecsv + "' to '"+filecsv.replace('pdf', 'xlsx') +
                    "' in " + "{:.2f}".format(time.time() - startedProcessing, 2) + " seconds.")
        # file_logger.info("Converted '"+ filecsv + "' to '"+filecsv.replace('pdf','xlsx')+"' in "+ "{:.2f}".format(time.time() - startedProcessing,2)+ " seconds.")
        return "Conversion Complete!"

    except Exception as e:
        logger.error("Error while processing file: "+str(e))
        # file_logger.error("Error while processing file: "+str(e))
        print("Error while processing file: "+str(e))


def generatingPackingSlip(RootFolder, ReqSource, OrderDate, ClientCode, formulaWorksheet, TemplateFiles):
    # file_logger = BKE_log.setup_custom_logger_file('root',RootFolder,OrderDate,ClientCode)
    try:
        # converting str to datetime
        OrderDate = datetime.strptime(OrderDate, '%Y-%m-%d')
        # extracting year from the order date
        year = OrderDate.strftime("%Y")
        # formatting order date {2022-00-00) format
        OrderDate = OrderDate.strftime('%Y-%m-%d')

        startedTemplating = time.time()
        sourcePivot = ReqSource
        source = TemplateFiles+"PackingSlip-Template.xlsx"
        destination = TemplateFiles+"TemplateFile.xlsx"

        # Making Copy of template file
        # shutil.copy(source, destination)
        # print("File copied successfully.")

        # Load work vook and sheets
        InputWorkbook = load_workbook(sourcePivot, data_only=True)
        # TemplateWorkbook = load_workbook(destination)

        InputSheet = InputWorkbook.active
        # TemplateSheet = TemplateWorkbook['ORDER']

        # Get rows and Column count
        df = pd.DataFrame(InputSheet, index=None)
        rows = len(df.axes[0])
        cols = len(df.axes[1])

        formulaWorksheet = load_workbook(formulaWorksheet, data_only=True)
        formulaSheet = formulaWorksheet['FormulaSheet']
        DBFformula = formulaWorksheet['DBF']

        # df_itemMaster = pd.read_excel(MasterFolderPath+'Item Master.xlsx',sheet_name='Item Master')
        # df_IGST = pd.read_excel(MasterFolderPath+'IGST Master.xlsx',sheet_name='DBF')
        # df_SGST = pd.read_excel(MasterFolderPath+'SGST Master.xlsx',sheet_name='DBF')
        print('Loading Master files for processing...')
        logger.info('Loading Master files for processing...')
        # file_logger.info('Loading Master files for processing...')
        # Opening Item Master Sheet
        df_item_master = pd.read_excel(
            MasterFolderPath+'Item Master.xlsx', sheet_name='Item Master', index_col=False)
        # if df_item_master != True:
        #     logger.info('<Column name> does not exist, please check item master')
        #     return

        print('Item Master loaded.')
        logger.info('Item Master loaded.')
        # file_logger.info('Item Master loaded.')
        # Opening IGST Master Sheet
        df_IGSTMaster = pd.read_excel(
            MasterFolderPath+'IGST Master.xlsx', sheet_name='DBF', index_col=False)
        print('IGST Master loaded.')
        logger.info('IGST Master loaded.')
        # file_logger.info('IGST Master loaded.')
        # Opening SGST Master Sheet
        df_SGSTMaster = pd.read_excel(
            MasterFolderPath+'SGST Master.xlsx', sheet_name='DBF', index_col=False)
        print('SGST Master loaded.')
        logger.info('SGST Master loaded.')
        # file_logger.info('SGST Master loaded.')
        # Opening Location2 Master Sheet
        df_Location2 = pd.read_excel(
            MasterFolderPath+'Location 2 Master.xlsx', sheet_name='Location2', index_col=False)
        print('Location 2 Master loaded.')
        logger.info('Location 2 Master loaded.')
        # file_logger.info('Location 2 Master loaded.')

        start_cols = 3

        for column in range(start_cols, cols-3):
            startedTemplatingFile = time.time()
            # Making Copy of template file
            shutil.copy(source, destination)
            # logger.info("Template File copied successfully for generating packing-slip")
            # print("Template File copied successfully for generating packing-slip")

            # Load work vook and sheets
            TemplateWorkbook = load_workbook(destination)
            TemplateSheet = TemplateWorkbook['ORDER']
            dbfsheet = TemplateWorkbook['DBF']

            TemplateSheet.cell(5, 1).value = InputSheet.cell(
                start_cols, column).value  # Order Name
            # TemplateSheet.cell(5,1).value = InputSheet.cell(7,column).value # Order Name
            # PO Number
            filename = InputSheet.cell(start_cols-1, column).value  # (2,col-3)
            TemplateSheet.cell(5, 2).value = InputSheet.cell(
                start_cols-1, column).value

            # Receving Location
            TemplateSheet.cell(5, 3).value = InputSheet.cell(
                start_cols+3, column).value

            TemplateSheet.cell(1, 1).value = 'Order Date'
            TemplateSheet.cell(1, 2).value = InputSheet.cell(
                1, 2).value  # Date

            TemplateSheet.cell(1, 3).value = 'SGST/IGST'
            TemplateSheet.cell(1, 4).value = InputSheet.cell(
                start_cols+2, column).value  # IGST/SGST Type (5,cols-3)
            if TemplateSheet.cell(1, 4).value == None:
                print(
                    "IGST/SGST TYPE = None, Requirment Summary file is not saved. Open the file, save it then process")
                logger.info(
                    "IGST/SGST TYPE = None, Requirment Summary file is not saved. Open the file, save it then process")
                # file_logger.info("IGST/SGST TYPE = None, Requirment Summary file is not saved. Open the file, save it then process")
                break

            # Copy EAN to template sheet
            Trows = 8
            Tcols = 5
            dbfrows = 2
            dbfcols = 57
            for row in range(8, rows):
                # if InputSheet.cell(row,column).value != None or InputSheet.cell(row,column).value != "":
                if str(InputSheet.cell(row, column).value).isnumeric():

                    # Copy Qty to template sheet
                    TemplateSheet.cell(Trows, Tcols).value = InputSheet.cell(
                        row, column).value
                    TemplateSheet.cell(
                        Trows, Tcols+1).value = InputSheet.cell(row, column).value
                    TemplateSheet.cell(
                        Trows, Tcols+2).value = "="+openpyxl.utils.cell.get_column_letter(Tcols+1)+str(Trows)  # Actual Qty

                    # Copy EAN to template sheet
                    TemplateSheet.cell(
                        Trows, Tcols-3).value = InputSheet.cell(row, 1).value

                    # VLOOKUP
                    # StyleName
                    # TemplateSheet.cell(Trows,Tcols-4).value = "="+formulaSheet.cell(3,2).value.replace("#VAL#",str(Trows))
                    TemplateSheet.cell(
                        Trows, Tcols-4).value = "='Hidden Item Master'!A"+str(Trows-6)

                    # style
                    # TemplateSheet.cell(Trows,Tcols-2).value =  "="+formulaSheet.cell(4,2).value.replace("#VAL#",str(Trows))
                    TemplateSheet.cell(
                        Trows, Tcols-2).value = "='Hidden Item Master'!C"+str(Trows-6)

                    # SADM SKU
                    # TemplateSheet.cell(Trows,Tcols-1).value = "="+formulaSheet.cell(5,2).value.replace("#VAL#",str(Trows))
                    TemplateSheet.cell(
                        Trows, Tcols-1).value = "='Hidden Item Master'!D"+str(Trows-6)

                    # Rate (in Rs.) Order file
                    # TemplateSheet.cell(Trows,Tcols+3).value = "="+formulaSheet.cell(6,2).value.replace("#VAL#",str(Trows))
                    TemplateSheet.cell(
                        Trows, Tcols+3).value = "='Hidden Item Master'!E"+str(Trows-6)

                    # Cls stk vs order
                    # TemplateSheet.cell(Trows,Tcols+6).value = TemplateSheet.cell(Trows,Tcols+5).value - TemplateSheet.cell(Trows,Tcols+2).value
                    TemplateSheet.cell(Trows, Tcols+6).value = "="+openpyxl.utils.cell.get_column_letter(
                        Tcols+5)+str(Trows) + '-'+openpyxl.utils.cell.get_column_letter(Tcols+2)+str(Trows)

                    # LOCATION2
                    # TemplateSheet.cell(Trows,Tcols+7).value = "="+formulaSheet.cell(7,2).value.replace("#VAL#",str(Trows))
                    TemplateSheet.cell(
                        Trows, Tcols+7).value = "='Hidden Item Master'!F"+str(Trows-6)

                    # BULK  / DTA  BULK  /  EOSS LOC
                    # TemplateSheet.cell(Trows,Tcols+8).value = "="+formulaSheet.cell(8,2).value.replace("#VAL#",str(Trows))
                    TemplateSheet.cell(
                        Trows, Tcols+8).value = "='Hidden Item Master'!G"+str(Trows-6)

                    # MRP
                    # TemplateSheet.cell(Trows,Tcols+9).value = "="+formulaSheet.cell(9,2).value.replace("#VAL#",str(Trows))
                    TemplateSheet.cell(
                        Trows, Tcols+9).value = "='Hidden Item Master'!E"+str(Trows-6)

                    # Closing stk
                    # TemplateSheet.cell(Trows,Tcols+5).value = "="+formulaSheet.cell(11,2).value.replace("#VAL#",str(Trows))
                    TemplateSheet.cell(
                        Trows, Tcols+5).value = InputSheet.cell(row, cols-2).value

                    # SCAN
                    # TemplateSheet.cell(Trows,Tcols+10).value = "="+formulaSheet.cell(12,2).value.replace("#VAL#",str(Trows))
                    TemplateSheet.cell(
                        Trows, Tcols+10).value = "=SUMIF('packing slip'!$B$3:$B$999,$B"+str(Trows)+",'packing slip'!$C$3:$C$999)"

                    # SCAN VS DIFF
                    # TemplateSheet.cell(Trows,Tcols+11).value = "="+formulaSheet.cell(13,2).value.replace("#VAL#",str(Trows))
                    TemplateSheet.cell(
                        Trows, Tcols+11).value = "=O"+str(Trows)+"-G"+str(Trows)

                    # ERROR
                    # TemplateSheet.cell(Trows,Tcols+12).value = "="+formulaSheet.cell(14,2).value.replace("#VAL#",str(Trows))
                    TemplateSheet.cell(
                        Trows, Tcols+12).value = "=IF(P"+str(Trows)+"<>0,1,0)"

                    # STYLE COLOR
                    # TemplateSheet.cell(Trows,Tcols+13).value = "="+formulaSheet.cell(15,2).value.replace("#VAL#",str(Trows))
                    TemplateSheet.cell(
                        Trows, Tcols+13).value = "=IF(P"+str(Trows)+"<=0,0,1)"

                    # Adding values to DBF
                    for i in range(1, dbfcols):
                        dbfsheet.cell(dbfrows, i).value = '='+DBFformula.cell(2, i+1).value.replace(
                            "#VAL#", str(Trows)).replace("#DBFROWS#", str(dbfrows))

                    Trows += 1
                    dbfrows += 1

            TemplateSheet.cell(5, 5).value = "=SUM(E8:E"+str(Trows-1)+")"
            TemplateSheet.cell(5, 6).value = "=SUM(F8:F"+str(Trows-1)+")"
            TemplateSheet.cell(5, 7).value = "=SUM(G8:G"+str(Trows-1)+")"

            TemplateWorkbook.save(RootFolder+"/"+ClientCode+"-"+year+"/" +
                                  OrderDate+"/70-Packing-Slip/"+"PackingSlip_"+str(filename)+".xlsx")
            TemplateWorkbook.close()

            # Opening Packing slip using openpyxl to check igst/sgst value
            sourcePackingSlip = RootFolder+"/"+ClientCode+"-"+year+"/" + \
                OrderDate+"/70-Packing-Slip/" + \
                "PackingSlip_"+str(filename)+".xlsx"
            TemplateWorkbook = load_workbook(sourcePackingSlip, data_only=True)
            TemplateSheet = TemplateWorkbook['ORDER']
            # print(TemplateSheet.cell(1,4).value)
            # Getting data from IGST/ SGST Sheet DBF to Tempalate sheet DBF

            # If IGST then open IGST Master
            if TemplateSheet.cell(1, 4).value == 'IGST':

                # Opening Packing slip as df for second time to get EAN values
                df_TempalateWorkbook = pd.read_excel(
                    sourcePackingSlip, sheet_name='ORDER', skiprows=6, index_col=False)
                # df_TempalateWorkbook.rename(columns = {'EAN':'EAN ID'}, inplace = True)

                # Temporary df to store EAN
                df_EAN_temp = df_TempalateWorkbook[['EAN']]
                # Applying join using EAN ID to get hidden_item_master to get SKU and Other fields
                df_hidden_item_master = df_EAN_temp.merge(
                    df_item_master, on='EAN', how='left')

                df_Location2_temp = df_hidden_item_master.merge(
                    df_Location2, on='EAN', how='left')
                # print(df_Location2_temp.head(10))
                df_Location2_temp.rename(
                    columns={'SKU ID': 'SKU'}, inplace=True)

                # Temporary df to store SKU
                df_SKU_temp = df_hidden_item_master[['SKU']]
                df_SKU_temp.rename(columns={'SKU': 'ITEMNAME'}, inplace=True)
                # Applying join using ITEMNAME to get hidden_dbf (from IGST/SGST sheet) to get SKU and Other fields
                df_GST_hidden = df_SKU_temp.merge(
                    df_IGSTMaster, on='ITEMNAME', how='left')

                with pd.ExcelWriter(sourcePackingSlip, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    # Adding Hidden Item Master sheet to the RQ sheet with values 'Style Name', 'EAN', 'Style', 'SKU', 'MRP'
                    df_Location2_temp.to_excel(writer, sheet_name='Hidden Item Master', index=False, columns=[
                                               'Style Name', 'EAN', 'Style', 'SKU', 'MRP', 'Location 2', 'BULK  / DTA  BULK  /  EOSS LOC'])
                    # # Adding Hidden DBF sheet to the RQ sheet with values 'Vouchertypename', 'CSNNO','DATE' ETC.
                    df_GST_hidden.to_excel(writer, sheet_name='Hidden DBF', index=False, columns=['Vouchertypename', 'CSNNO', 'DATE',
                                                                                                  'REFERENCE', 'REF1', 'DEALNAME', 'PRICELEVEL', 'ITEMNAME', 'GODOWN', 'QTY', 'RATE', 'SUBTOTAL', 'DISCPERC',
                                                                                                  'DISCAMT', 'ITEMVALUE', 'LedgerAcct', 'CATEGORY1', 'COSTCENT1', 'CATEGORY2', 'COSTCENT2', 'CATEGORY3', 'COSTCENT3',
                                                                                                  'CATEGORY4', 'COSTCENT4', 'ITEMTOTAL', 'TOTALQTY', 'CDISCHEAD', 'CDISCPERC', 'COMMONDISC', 'BEFORETAX',
                                                                                                  'TAXHEAD', 'TAXPERC', 'TAXAMT', 'STAXHEAD', 'STAXPERC', 'STAXAMT', 'ITAXHEAD', 'ITAXPERC', 'ITAXAMT', 'NETAMT',
                                                                                                  'ROUND', 'ROUND1', 'REFTYPE', 'Name', 'REFAMT', 'Narration', 'Transport', 'transmode', 'pymtterm', 'ordno',
                                                                                                  'orddate', 'DANO', 'Delyadd1', 'Delyadd2', 'Delyadd3', 'Delyadd4'])

            # If SGST then open SGST Master
            if TemplateSheet.cell(1, 4).value == 'SGST':

                # Opening Packing slip as df for second time to get EAN values
                df_TempalateWorkbook = pd.read_excel(
                    sourcePackingSlip, sheet_name='ORDER', skiprows=6, index_col=False)
                # df_TempalateWorkbook.rename(columns = {'EAN':'EAN ID'}, inplace = True)

                # Temporary df to store EAN ID
                df_EAN_temp = df_TempalateWorkbook[['EAN']]
                # Applying join using EAN ID to get hidden_item_master to get SKU and Other fields
                df_hidden_item_master = df_EAN_temp.merge(
                    df_item_master, on='EAN', how='left')

                df_Location2_temp = df_hidden_item_master.merge(
                    df_Location2, on='EAN', how='left')
                # print(df_Location2_temp)

                # Temporary df to store SKU
                df_SKU_temp = df_hidden_item_master[['SKU']]
                df_SKU_temp.rename(columns={'SKU': 'ITEMNAME'}, inplace=True)
                # Applying join using ITEMNAME to get hidden_dbf (from IGST/SGST sheet) to get SKU and Other fields
                df_GST_hidden = df_SKU_temp.merge(
                    df_SGSTMaster, on='ITEMNAME', how='left')

                with pd.ExcelWriter(sourcePackingSlip, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    # Adding Hidden Item Master sheet to the RQ sheet with values 'Style Name', 'EAN', 'Style', 'SKU', 'MRP'
                    df_Location2_temp.to_excel(writer, sheet_name='Hidden Item Master', index=False, columns=[
                                               'Style Name', 'EAN', 'Style', 'SKU', 'MRP', 'Location 2', 'BULK  / DTA  BULK  /  EOSS LOC'])
                    # # Adding Hidden DBF sheet to the RQ sheet with values 'Vouchertypename', 'CSNNO','DATE' ETC.
                    df_GST_hidden.to_excel(writer, sheet_name='Hidden DBF', index=False, columns=['Vouchertypename', 'CSNNO', 'DATE',
                                                                                                  'REFERENCE', 'REF1', 'DEALNAME', 'PRICELEVEL', 'ITEMNAME', 'GODOWN', 'QTY', 'RATE', 'SUBTOTAL', 'DISCPERC',
                                                                                                  'DISCAMT', 'ITEMVALUE', 'LedgerAcct', 'CATEGORY1', 'COSTCENT1', 'CATEGORY2', 'COSTCENT2', 'CATEGORY3', 'COSTCENT3',
                                                                                                  'CATEGORY4', 'COSTCENT4', 'ITEMTOTAL', 'TOTALQTY', 'CDISCHEAD', 'CDISCPERC', 'COMMONDISC', 'BEFORETAX',
                                                                                                  'TAXHEAD', 'TAXPERC', 'TAXAMT', 'STAXHEAD', 'STAXPERC', 'STAXAMT', 'ITAXHEAD', 'ITAXPERC', 'ITAXAMT', 'NETAMT',
                                                                                                  'ROUND', 'ROUND1', 'REFTYPE', 'Name', 'REFAMT', 'Narration', 'Transport', 'transmode', 'pymtterm', 'ordno',
                                                                                                  'orddate', 'DANO', 'Delyadd1', 'Delyadd2', 'Delyadd3', 'Delyadd4'])
                pass

            TemplateWorkbook.close()

            # Opening Sheet again to hide hidden dbf and item master sheets
            TemplateWorkbook = load_workbook(sourcePackingSlip)
            hidden_item_master = TemplateWorkbook['Hidden Item Master']
            hidden_dbf = TemplateWorkbook['Hidden DBF']

            hidden_item_master.sheet_state = 'hidden'
            hidden_dbf.sheet_state = 'hidden'

            TemplateWorkbook.close()
            TemplateWorkbook.save(sourcePackingSlip)

            logger.info("Packing slip generated for: "+str(filename) +
                        " file in {:.2f}".format(time.time() - startedTemplatingFile, 2) + " seconds.")
            # file_logger.info("Packing slip generated for: "+str(filename)+ " file in {:.2f}".format(time.time() - startedTemplatingFile,2)+ " seconds.")
            print("Packing slip generated for: "+str(filename) +
                  " file in {:.2f}".format(time.time() - startedTemplatingFile, 2) + " seconds.")

        logger.info("Total time taken for generation of packing-slips:  {:.2f}".format(
            time.time() - startedTemplating, 2) + " seconds.")
        # file_logger.info("Total time taken for generation of packing-slips:  {:.2f}".format(time.time() - startedTemplating,2)+ " seconds.")
        print("Total time taken for generation of packing-slips:  {:.2f}".format(
            time.time() - startedTemplating, 2) + " seconds.")
        return 'Completed!'
    except Exception as e:
        logger.error("Error while generating packing-slip file: "+str(e))
        # file_logger.info("Error while generating packing-slip file: "+str(e))
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)
        print("Error while generating packing-slip file: "+str(e))
