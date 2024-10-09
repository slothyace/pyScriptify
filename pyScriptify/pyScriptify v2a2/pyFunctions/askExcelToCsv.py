import glob
import os
import pyFunctions.baseFunctions.baseFunctions as psBaseFunc
import pyFunctions.askDeleteFile as askDeleteFile

def single(console, filePath):
    psBaseFunc.excelToCsv(console, filePath)
    askDeleteFile.single(console, filePath)

def multi(console, folderPath):
    excelFiles = glob.glob(os.path.join(folderPath, f"*.xlsx"))
    convertedFiles = []
    if not excelFiles:
        psBaseFunc.consolePrint(console, f"No .xlsx files found in {folderPath}.")
    else:
        for xlsxFile in excelFiles:
            psBaseFunc.excelToCsv(console, xlsxFile)
            convertedFiles.append(xlsxFile)
        askDeleteFile.multi(console, convertedFiles)