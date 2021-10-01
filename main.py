import os

import numpy as np
import openpyxl
import pandas as pd
import os.path
import shutil
import logging

from datetime import datetime
from pathlib import Path
from openpyxl.styles import Border, Side
# from Utils import Utils
from os import path


def folderStructureCreation(dirPath):
    print("Checking Folder structure...", end="")

    folderList = ['in', 'out', 'log', 'template']

    for folder in folderList:
        folderPath = os.path.join(dirPath, folder)

        # Check if folder exist
        if not path.exists(folderPath):
            os.makedirs(folderPath)

    print("Done!")

    pass


def fixHeaderColumn(df):
    df.columns = map(str.lower, df.columns)

    if 'indigenous' in df.columns.tolist():
        arr = {'indigenous': 'company'}
        df.rename(columns=arr, inplace=True)

    if 'vaccine site' in df.columns.tolist():
        arr = {'vaccine site': 'vaccination site'}
        df.rename(columns=arr, inplace=True)

    if 'time' not in df.columns.tolist():
        df['time'] = ''

    return df[['priority group*', 'sub-priority group*', 'last_name*', 'first_name*',
               'middle_name*', 'suffix', 'mobile number\n(format:9170123456)*',
               'current_residence:_region*', 'current_residence:\nprovince*',
               'current_residence:\nmunicipality/city*',
               'current_residence:\nbarangay*', 'sex*', 'birthdate_mm/dd/yyyy_*',
               'occupation*', 'allergy to vaccines or components of vaccines*',
               'with_comorbidity?*', 'email*', 'employee number*', 'company', 'time',
               'vaccination site']]


def getData(fileName):
    filePath = os.path.join(inPath, fileName)
    fileName = os.path.splitext(fileName)[0]
    tagging = fileName.split("_AZ_")[0].split("_")[1]
    df = pd.read_excel(filePath, dtype=str, na_filter=False)

    df = fixHeaderColumn(df)

    df['Tagging'] = tagging
    df['File Name'] = fileName

    return df


def duplicateTemplateLTGC(tempLTGC_Path, out, outputFilename):
    companyDir = out + "/"
    srcFile = companyDir + outputFilename + ".xlsx"

    if not os.path.isfile(srcFile):
        shutil.copy(tempLTGC_Path, srcFile)

    return companyDir + outputFilename + ".xlsx"


if __name__ == '__main__':
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)
    pd.set_option('display.width', None)

    today = datetime.today()
    dateTime = today.strftime("%m%d%y%H%M%S")

    dirPath = r"/Users/Ran/Documents/Vaccine/LTGC_OffSite"

    inPath = os.path.join(dirPath, "in")
    outPath = os.path.join(dirPath, "out")
    logPath = os.path.join(dirPath, "log")
    backupPath = os.path.join(dirPath, "backup")
    templateFilePath = os.path.join(dirPath, "template/VIMSTemplate.xlsx")
    outFilename = 'Consolidated_LTGC_Offsite_' + dateTime

    # Folder Structure Creation
    folderStructureCreation(dirPath)

    # Get all Files in
    arrFilenames = os.listdir(inPath)

    arrdfFrames = []

    for inFile in arrFilenames:
        if not inFile == ".DS_Store":
            print("Reading: " + inFile + "......")

            arrdfFrames.append(getData(inFile))

            # arrdf.append(getData(inFile))

    # Merge all df
    df_master = pd.concat(arrdfFrames)

    # Create copy of template file and save it to out folder
    templateFile = duplicateTemplateLTGC(templateFilePath, outPath, outFilename)

    # Write df_master(consolidated/append data) to excel
    writer = pd.ExcelWriter(templateFile, engine='openpyxl', mode='a')
    writer.book = openpyxl.load_workbook(templateFile)
    writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
    df_master.to_excel(writer, sheet_name="Eligible Population", startrow=1, header=False, index=False)
    writer.save()

    print(df_master)
