import glob
import os
import pyFunctions.baseFunctions.baseFunctions as psBaseFunc
import pyFunctions.askDeleteFile as askDeleteFile

def single(console, filePath):
    psBaseFunc.csvToExcel(console, filePath)
    askDeleteFile.single(console, filePath)

def multi(console, folderPath):
    csvFiles = glob.glob(os.path.join(folderPath, f"*.csv"))
    convertedFiles = []
    if not csvFiles:
        psBaseFunc.consolePrint(console, f"No .csv files found in {folderPath}.")
    else:
        for csvFile in csvFiles:
            psBaseFunc.csvToExcel(console, csvFile)
            convertedFiles.append(csvFile)
        askDeleteFile.multi(console, convertedFiles)
