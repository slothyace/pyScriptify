# StdLib
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import os
import subprocess
import sys
import time
import json
import webbrowser

# CtmLib
import pyFunctions
import pyFunctions.askCsvToExcel
import pyFunctions.baseFunctions.baseFunctions as psBaseFunc
import pyFunctions.baseFunctions.modMan as modMan

# ExtLib
extLib = ["ttkbootstrap", "ttkbootstrap.constants", "pyfiglet"]
modMan.modMan(extLib)
import ttkbootstrap as ttk
import ttkbootstrap.constants as ttkconstants
import pyfiglet as pyfig

settings = json.load(open(f"pyAssets/config.json", f"r"))
os.system(f"cls")
print(f"{pyfig.figlet_format(settings["config"]["appName"])}")
print(settings[f"config"][f"version"])

if settings["debugger"]["pythonTerminal"] == "False":
    if not sys.executable.endswith(f"pythonw.exe"):
        os.system(f"start pythonw.exe pyScriptify.py")
        print(f"Relaunching with pythonw...")
        time.sleep(1)
        sys.exit()

def fBrowser(console, function, scale):
    match function + scale:
        case "csvToExcelsingle":
            filePath=filedialog.askopenfilename(filetypes=[("CSV", "*.csv")])
            if filePath:
                psBaseFunc.consoleClear(console)
                pyFunctions.askCsvToExcel.single(console, filePath)

        case "csvToExcelmulti":
            folderPath=filedialog.askdirectory()
            if folderPath:
                psBaseFunc.consoleClear(console)
                pyFunctions.askCsvToExcel.multi(console, folderPath)

        case "excelToCsvsingle":
            filePath=filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
            if filePath:
                psBaseFunc.consoleClear(console)
                pyFunctions.askExcelToCsv.single(console, filePath)

        case "excelToCsvmulti":
            folderPath=filedialog.askdirectory()
            if folderPath:
                psBaseFunc.consoleClear(console)
                pyFunctions.askExcelToCsv.multi(console, folderPath)

        case "splitExcelsingle":
            filePath=filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
            if filePath:
                psBaseFunc.consoleClear(console)
                pyFunctions.askSplitExcel.single(console, filePath)

        case "splitExcelmulti":
            folderPath=filedialog.askdirectory()
            if folderPath:
                psBaseFunc.consoleClear(console)
                pyFunctions.askSplitExcel.multi(console, folderPath)

        case "mergeExcelsingle":
            folderPath=filedialog.askdirectory()
            if folderPath:
                psBaseFunc.consoleClear(console)
                pyFunctions.askMergeExcel.single(console, folderPath)
        case "mergeExcelmulti":
            foldersPath=filedialog.askdirectory()
            if foldersPath:
                psBaseFunc.consoleClear(console)
                pyFunctions.askMergeExcel.multi(console, foldersPath)

        case "folderIndexsingle":
            folderPath=filedialog.askdirectory()
            extension=simpledialog.askstring("Extension filter", "csv, xlsx, mp3, etc. | * for wildcard").strip(f".")
            if folderPath and extension:
                psBaseFunc.consoleClear(console)
                pyFunctions.createFolderIndex.single(console, folderPath, extension)
        case "folderIndexmulti":
            foldersPath=filedialog.askdirectory()
            extension=simpledialog.askstring("Extension filter", "csv, xlsx, mp3, etc. | * for wildcard").strip(f".")
            if foldersPath and extension:
                psBaseFunc.consoleClear(console)
                pyFunctions.createFolderIndex.multi(console, foldersPath, extension)

winY = int(settings[f"config"][f"winY"])
winX = int(1.5*winY)



def GUI():
    GUImain = ttk.Window(themename=f"darkly")
    GUImain.title(f"{settings[f"config"][f"appName"]} Ver. {settings[f"config"][f"version"]}")
    GUImain.geometry(f"{winX}x{winY}")
    GUImain.resizable(True, True)
    GUImain.minsize(winX, winY)
    ttk.Separator(GUImain, bootstyle="success", orient="horizontal").pack(fill="x")

    def GUImainBanner():
        global topFrame
        topFrame = ttk.LabelFrame(GUImain, bootstyle="success")
        topFrame.pack(pady=(0,10), padx=10, fill="x")
        ttk.Label(topFrame, text=pyfig.figlet_format(settings["config"]["appName"]), font=(settings['config']['font'], 8)).pack(padx=10)
        ttk.Label(topFrame, text=(f"Version: {settings["config"]["version"]}"), font=(settings['config']['font'], 8)).pack(side="right", padx=10)
    GUImainBanner()

    ttk.Separator(GUImain, bootstyle="success", orient="horizontal").pack(fill="x")

    def GUImainButtons():
        buttonFrame = ttk.Frame(GUImain)
        buttonFrame.pack(pady=8, padx=8, fill="none")

        csvToExcelBtn = ttk.Menubutton(buttonFrame, text="CSV -> Excel", bootstyle="success", width=15)
        csvToExcelMenu = ttk.Menu(csvToExcelBtn)
        csvToExcelBtn["menu"] = csvToExcelMenu
        csvToExcelMenu.add_command(label="File", command=lambda: fBrowser(console, function="csvToExcel", scale="single"))
        csvToExcelMenu.add_command(label="Folder", command=lambda: fBrowser(console, function="csvToExcel", scale="multi"))
        csvToExcelMenu.add_command(label="Info", command= lambda: (psBaseFunc.consoleClear(console), psBaseFunc.consolePrint(console, string=f"Converts a CSV file to an Excel file.")))

        excelToCsvBtn = ttk.Menubutton(buttonFrame, text="Excel -> CSV", bootstyle="success", width=15)
        excelToCsvMenu = ttk.Menu(excelToCsvBtn)
        excelToCsvBtn["menu"] = excelToCsvMenu
        excelToCsvMenu.add_command(label="File", command=lambda: fBrowser(console, function="excelToCsv", scale="single"))
        excelToCsvMenu.add_command(label="Folder", command=lambda: fBrowser(console, function="excelToCsv", scale="multi"))
        excelToCsvMenu.add_command(label="Info", command=lambda: (psBaseFunc.consoleClear(console), psBaseFunc.consolePrint(console, string=f"Converts an excel file to a CSV file.")))

        splitExcelBtn = ttk.Menubutton(buttonFrame, text="Split Excel", bootstyle="success", width=15)
        splitExcelMenu = ttk.Menu(splitExcelBtn)
        splitExcelBtn["menu"] = splitExcelMenu
        splitExcelMenu.add_command(label="File", command=lambda: fBrowser(console, function="splitExcel", scale="single"))
        splitExcelMenu.add_command(label="Folder", command=lambda: fBrowser(console, function="splitExcel", scale="multi"))
        splitExcelMenu.add_command(label="Info", command=lambda: (psBaseFunc.consoleClear(console), psBaseFunc.consolePrint(console, string=f"Splits an excel workbook's sheets into individual excel files.")))

        mergeExcelBtn = ttk.Menubutton(buttonFrame, text="Merge Excel", bootstyle="success", width=15)
        mergeExcelMenu = ttk.Menu(mergeExcelBtn)
        mergeExcelBtn["menu"] = mergeExcelMenu
        mergeExcelMenu.add_command(label="Single", command=lambda: fBrowser(console, function="mergeExcel", scale="single"))
        mergeExcelMenu.add_command(label="Multi", command=lambda: fBrowser(console, function="mergeExcel", scale="multi"))
        mergeExcelMenu.add_command(label="Info", command=lambda: (psBaseFunc.consoleClear(console), psBaseFunc.consolePrint(console, string=f"Merges multiple excel files into sheets in 1 excel workbook.")))

        folderIndexBtn = ttk.Menubutton(buttonFrame, text="List Folder Files", bootstyle="success", width=15)
        folderIndexMenu = ttk.Menu(folderIndexBtn)
        folderIndexBtn["menu"] = folderIndexMenu
        folderIndexMenu.add_command(label="Single", command=lambda: fBrowser(console, function="folderIndex", scale="single"))
        folderIndexMenu.add_command(label="Multi", command=lambda: fBrowser(console, function="folderIndex", scale="multi"))
        folderIndexMenu.add_command(label="Info", command=lambda: (psBaseFunc.consoleClear(console), psBaseFunc.consolePrint(console, string=f"Produces a .txt file containing the names of the files inside the folder.")))

        csvToExcelBtn.grid(row=0, column=0, padx=2, pady=2)
        excelToCsvBtn.grid(row=1, column=0, padx=2, pady=2)
        splitExcelBtn.grid(row=0, column=1, padx=2, pady=2)
        mergeExcelBtn.grid(row=1, column=1, padx=2, pady=2)
        folderIndexBtn.grid(row=0, column=3, rowspan=2, sticky="NESW", padx=2, pady=2)
    GUImainButtons()

    ttk.Separator(GUImain, bootstyle="success", orient="horizontal").pack(fill="x")

    def GUImainConsole():
        global console
        consoleFrame = ttk.LabelFrame(GUImain, bootstyle="success", text=" Console ")
        consoleFrame.pack(pady=(0,10), padx=10, fill="both", expand=True)

        console = ttk.Text(consoleFrame, wrap=tk.WORD, width=1, height=1, font=(settings['config']['font'], 12))
        console.pack(pady=0, padx=10, fill="both", expand=True)
        console.insert(tk.END, f"Welcome to {settings["config"]["appName"]}.")

        consoleClearBtn = ttk.Button(consoleFrame, text="Clear Console", command=lambda: psBaseFunc.consoleClear(console), bootstyle="success", width=20)
        consoleClearBtn.pack(pady=(0,10), padx=10, fill="x", expand=False)
    GUImainConsole()

    ttk.Separator(GUImain, bootstyle="success", orient="horizontal").pack(fill="x")

    def GUIdonateButton():
        ttk.Button(GUImain, text="Like this project? Sponsor me!", command=lambda: webbrowser.open("https://ko-fi.com/slothyacedia"), bootstyle="success").pack(padx=10, pady=10)
    
    if settings["hidden"]["donate"] != "True":
        GUIdonateButton()
    

    def GUImainBindings():
        console.bind("<Key>", lambda e: "break")
        console.bind("<Button-1>", lambda e: "break")
        console.bind("<Button-2>", lambda e: "break")
        console.bind("<Button-3>", lambda e: "break")
    GUImainBindings()

    GUImain.mainloop()

GUI()