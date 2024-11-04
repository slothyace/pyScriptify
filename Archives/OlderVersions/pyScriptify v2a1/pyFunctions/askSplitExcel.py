import glob
import os
import pyFunctions.baseFunctions.baseFunctions as psBaseFunc
import pyFunctions.askDeleteFile as askDeleteFile

def single(console, filePath):
    psBaseFunc.splitExcel(console, filePath)
    askDeleteFile.single(console, filePath)

def multi(console, folderPath):
    excelFiles = glob.glob(os.path.join(folderPath, f"*.xlsx"))
    splitFiles = []
    if not excelFiles:
        psBaseFunc.consolePrint(console, f"No .xlsx files found in {folderPath}.")
    else:
        for xlsxFile in excelFiles:
            psBaseFunc.splitExcel(console, xlsxFile)
            splitFiles.append(xlsxFile)
        askDeleteFile.multi(console, splitFiles)