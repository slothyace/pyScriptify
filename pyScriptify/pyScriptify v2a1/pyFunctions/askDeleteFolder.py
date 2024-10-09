import tkinter as tk
import pyFunctions.baseFunctions.baseFunctions as psBaseFunc

def single(console, folderPath):
    deleteChoice = tk.messagebox.askyesno(f"Delete Folder", f"Delete Original Folder?")
    if deleteChoice:
        psBaseFunc.deleteFolder(console, folderPath)
    else:
        psBaseFunc.consolePrint(console, f"Original folder not deleted.")
    psBaseFunc.consoleSee(console)

def multi(console, foldersPath):
    deleteChoice = tk.messagebox.askyesno(f"Delete Folder", f"Delete Original Folders?")
    if deleteChoice:
        for folderPath in foldersPath:
            psBaseFunc.deleteFolder(console, folderPath)
    else:
        psBaseFunc.consolePrint(console, f"Original folders not deleted.")
    psBaseFunc.consoleSee(console)