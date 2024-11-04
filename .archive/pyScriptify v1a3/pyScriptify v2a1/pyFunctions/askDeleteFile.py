import tkinter as tk
import pyFunctions.baseFunctions.baseFunctions as psBaseFunc

def single(console, filePath):
    deleteChoice = tk.messagebox.askyesno(f"Delete File", f"Delete Original File?")
    if deleteChoice:
        psBaseFunc.deleteFile(console, filePath)
    else:
        psBaseFunc.consolePrint(console, f"Original file not deleted.")
    psBaseFunc.consoleSee(console)

def multi(console, folderPath):
    deleteChoice = tk.messagebox.askyesno(f"Delete File", f"Delete Original Files?")
    if deleteChoice:
        for filePath in folderPath:
            psBaseFunc.deleteFile(console, filePath)
    else:
        psBaseFunc.consolePrint(console, f"Original files not deleted.")
    psBaseFunc.consoleSee(console)