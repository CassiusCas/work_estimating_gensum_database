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
    
    # OPEN FILE: gensumw
    try:
        wb_exist_gensum=openpyxl.load_workbook(gensum_run_path,data_only=True)
    except:
        print("Error when opening workbook.\nTry deleting all other worksheets in excel document and resaving file.\nIf still does not work contact: Jonathancascioli@gmail.com")
    else:
        print("workbooko successfully imported.")

    print("The following worksheets are in the excel document.")
    for sheet in wb_exist_gensum:
       print(sheet.title)
    
    worksheet_open_question=input("Which worksheet would you like to open?:\n")
    
    try:
        ws = wb_exist_gensum[worksheet_open_question]
    except:
        print("That Worksheet does not exist")
    else:
        return
    
    # CREATE EXCEL: document to place extracted data into
    create_wb=openpyxl.Workbook()
    
    # CREATE PATH: path to use for resulting file of program
    gensum_prj_file = prj_folder+ r"\\"+current_time+"__-__gensum.xlsx"

    # SAVE EXCEL: save excel file to path location
    create_wb.save(gensum_prj_file)

    #Write cell with value of stop at the bottom of the 
    # Pull project information from top of sheet and 

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
