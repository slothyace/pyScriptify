import glob
import os
import pyFunctions.baseFunctions.baseFunctions as psBaseFunc
import pyFunctions.askDeleteFolder as askDeleteFolder

def single(console, folderPath):
    excelFiles = glob.glob(os.path.join(folderPath, f"*.xlsx"))
    if not excelFiles:
        psBaseFunc.consolePrint(console, f"No .xlsx files found in {folderPath}.")
    else:
        psBaseFunc.mergeExcel(console, folderPath)
        askDeleteFolder.single(console, folderPath)

def multi(console, foldersPath):
    folders = os.listdir(foldersPath)
    foldersMerged = []
    for folder in folders:
        folderPath = os.path.join(foldersPath, folder)
        if os.path.isdir(folderPath):
            excelFiles = glob.glob(os.path.join(folderPath, f"*.xlsx"))
            if not excelFiles:
                psBaseFunc.consolePrint(console, f"No .xlsx files found in {folderPath}.")
            else:
                psBaseFunc.mergeExcel(console, folderPath)
                foldersMerged.append(folderPath)
    askDeleteFolder.multi(console, foldersMerged)
    