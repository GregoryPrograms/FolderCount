import os
import tkinter
from xlrd import open_workbook
from tkinter import filedialog
from pathlib import Path

#First, we want to get the folder location of all the naming documents
tkinter.Tk().withdraw()
excelDirPath = Path(filedialog.askdirectory(title = 'Select directory containing excel sheets...'))
folderDirPath = Path(filedialog.askdirectory(title = 'Select directory containing folders...'))
errorFile = open(folderDirPath / 'Missing File List.txt', 'a+')
#For each excel doc in the directory, we want to go through and combine the values into one complete data struct
#Otherwise, we would have to repetitively iterate through every excel document to look for the ones we want.
for excelPath in os.listdir(excelDirPath):
    #We don't want to try opening any non-excel documents, though.
    if(excelPath.endswith(".xlsx")):
        #Operating on the per - workbook level
        excelDoc = open_workbook(excelDirPath / excelPath)
        for sheet in excelDoc.sheets():
            #Operating on the per - sheet level
            #For each naming sheet (box), we check the folder name column and make sure every one of them is a file.
            for cell in sheet.col(2):
                if(cell.value == "BARCODE"):
                    continue
                else:
                    if(Path.is_file(folderDirPath / cell.value)):
                        continue
                    else:
                    #If we are at this section of code, we've found a folder name that isn't in the directory we specified. 
                    #We print this to a file, as well as the name of the previous file
                        errorFile.write(cell.value + "\n")
errorFile.close()