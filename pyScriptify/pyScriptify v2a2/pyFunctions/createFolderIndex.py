import pyFunctions.baseFunctions.baseFunctions as psBaseFunc
import os

def single(console, folderPath, extension):
    psBaseFunc.listFilesInFolder(console, folderPath, extension)

def multi(console, foldersPath, extension):
    folders = os.listdir(foldersPath)
    for folder in folders:
        folderPath = os.path.join(foldersPath, folder)
        if os.path.isdir(folderPath):
            psBaseFunc.listFilesInFolder(console, folderPath, extension)