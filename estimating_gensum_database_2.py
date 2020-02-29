#NEXT: copy brs file template to prj folder then replace paths in program to point to new location
#NEXT: transfer user input information to gensum
#next: transfer brs first column company to gensum
#next: transfer final brs results to historical database 
#import libraries
import numpy as np
import pandas as pd
import urllib
import os
import openpyxl
#import winshell
from win32com.client import Dispatch
from shutil import copy as cp
import datetime
#from uszipcode import SearchEngine, SimpleZipcode, Zipcode


def import_csv(file):
    #import csv file
    df=pd.read_csv(file)
    return df
def get_download_path():
    #return default download path for linux or windows
        if os.name == 'nt':
            import winreg
            sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
            downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
                location = winreg.QueryValueEx(key, downloads_guid)[0]
            return location
        else:
            return os.path.join(os.path.expanduser('~'), 'downloads')


def master_program():

    # GET PATH: users download directory on windows computer
    download_path = get_download_path()
    
    # CREATE PATH: creats folder path of default download path with default folder where gensum file data is to be extracted from is to be placed
    project_setup_directory= download_path +r'\_02-estimating_gensum_database'
    
    # CREATE PATH:Creates folder path with a name of the time program was ran
    time_now = datetime.datetime.now()
    current_time =str(time_now)[0:19].replace(":","-")
    print(current_time)
    prj_folder = project_setup_directory+r"\\"+current_time
    
   
    # CREATE FOLDER: create project folder with name of time program ran
    #  This avoids confusion if ran several times
    #  This is where all files resulting from running program will be stored
    if os.path.exists(prj_folder):
        return
    else:
        os.makedirs(prj_folder)
    
    # CREATE PATH: Location of file that data is extracted from 
    gensum_run_path=project_setup_directory +r'\gensum_run.xlsx'
    
    # OPEN FILE: gensum_run to extract info from
    try:
        wb_exist_gensum=openpyxl.load_workbook(gensum_run_path,data_only=True)
    except:
        print("Error when opening workbook.\nTry deleting all other worksheets in excel document and resaving file.\nMake sure file is in correct location and named 'gensum_run.xlsx'\nIf still does not work contact: Jonathancascioli@gmail.com")
    else:
        print("workbook successfully imported.")

    # TELL USER: what worksheets are in excel document
    print("The following worksheets are in the excel document.")
    
    for sheet in wb_exist_gensum:
       print(sheet.title)
    
    # INPUT: Ask user which worksheet he wants to open
    worksheet_open_question=input("Which worksheet would you like to open?:\n")
    
    # TRY: To open user input worksheet
    try:
        ws = wb_exist_gensum[worksheet_open_question]
    except:
        print("That Worksheet does not exist\n")
    else:
        print("Worksheet successfully opened\n")
    
    #FIND CELL: Find cell with "start" in it to set the beginning of a range to pull data from
    for row in ws.iter_rows():
        for cell in row:
            if cell.value =="START":
                start_loc=cell.coordinate
                start_offset=ws[start_loc].offset(1,0).coordinate
    print("Start Location:"+start_loc+"\n")
    print("Start offset:"+start_offset+"\n")

    #FIND CELL: Find cell with "STOP" in it to set the end of a range to pull data from
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == "STOP":
                stop_loc=cell.coordinate
                stop_offset=ws[stop_loc].offset(-1,0).coordinate
    print("Stop location:"+stop_loc+"\n")
    print("Stop Offset:"+stop_offset+"\n")



    # CREATE EXCEL: document to place extracted data into
    new_wb=openpyxl.Workbook()
    
    # CREATE SHEET: create workbook sheet
    ws2=new_wb.create_sheet("BID_PACKAGES",0)

    #SET HEADERS: Set headers for each column
    ws2.cell(row=1,column=1,value="Time Ran")
    ws2.cell(row=1,column=2,value="Project")
    ws2.cell(row=1,column=3,value="BP#")
    ws2.cell(row=1,column=4,value="Bid Package Description")
    ws2.cell(row=1,column=5,value="Package_Total")

        
   
    counter=2
    # PULL RANGE: pull Bid Division
    division_cell_range=ws[start_offset:stop_offset]
    for cell, in division_cell_range:
        division=cell.value
        print(division)
        ws2.cell(row=counter,column=3,value=division)
        counter+=1
            
    #PULL RANGE: pull bid package description     
    counter=2
    bid_package_start=ws[start_offset].offset(0,1).coordinate
    bid_package_stop=ws[stop_offset].offset(0,1).coordinate
    bid_package_cell_range=ws[bid_package_start:bid_package_stop]
    for cell, in bid_package_cell_range:
        bid_package=cell.value
        print(bid_package)
        ws2.cell(row=counter,column=4,value=bid_package)
        counter+=1


    #PULL RANGE:  pull total cost column
    counter=2
    total_col_start=ws[start_offset].offset(0,2).coordinate
    total_col_stop=ws[stop_offset].offset(0,2).coordinate
    total_col_range=ws[total_col_start:total_col_stop]
    for cell, in total_col_range:
        tol_value=cell.value
        print(tol_value)
        ws2.cell(row=counter,column=5,value=tol_value)
        counter+=1


    #FIND CELL: FIND Project Name Cell, then offset over 1 column to retrieve project name info
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == "Project Name:":
                prj_n_coord=cell.coordinate
                print(prj_n_coord)
                prj_n_value=ws[prj_n_coord].offset(0,1).value
    print("Project Name is:"+prj_n_value+"\n\n\n")

    #PULL RANGE: Pull project name and apply over every row of bid package length
    counter=2
    for x in range(len(total_col_range)):
        ws2.cell(row=counter,column=2,value=prj_n_value)
        counter+=1

    #ITTERATE DATE: itterate current time(got earlier) over every row in first column to ref when run
    counter=2
    for x in range(len(total_col_range)):
        ws2.cell(row=counter,column=1,value=current_time)
        counter+=1

    # CREATE PATH: path to use for resulting file of program
    gensum_prj_file = prj_folder+ r"\\"+current_time+"__-__gensum.xlsx"

    # SAVE EXCEL: save excel file to path location
    new_wb.save(gensum_prj_file)



    #####ENTER PROJECT INFORMATION INTO BRS WORKSHEET#####
    ws_p_info = wb_brs.worksheets[0]
    ws_p_info["B2"] = prj_name
    ws_p_info["B3"] = prj_numb
    ws_p_info["B4"] = bid_date
    ws_p_info["B5"] = location
    ws_p_info["B6"] = t_office
    ws_p_info["B7"] = client
    ws_p_info["B8"] = lead_est

    #save brs
    new_brs = project_setup_directory + r'\new_brs.xlsx'
    wb_brs.save(new_brs)


    #import newly saved excel file with openpyXL for manipulation
    wb = openpyxl.load_workbook(file_name)

    #test if open py excel has properly imported excel doc
    try:
        for sheet in wb:
            print(sheet.title)
    except:
        print("Error when opening workbook with OpenPyXL")
    else:
        print("workbook successfully imported with OpenPyXl")

    for x in range(len(df2)):
        sheet_name = df2.iloc[x,0] +' - '+df2.iloc[x,1]
        sheet_name_rep = sheet_name.replace('=', '').replace('"','').replace('(','').replace(')','').replace('and','&').replace(',','')
        wb.create_sheet(sheet_name_rep)

    for sheet in wb:
        print(sheet.title)

    wb.save(file_name)


master_program()
