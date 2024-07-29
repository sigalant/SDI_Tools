import openpyxl as opx
import tkinter as tk
from datetime import date
from tkinter import filedialog
from PIL import Image, ImageTk

inputFilepath = ""
outputFilepath = ""

root = tk.Tk()

root.title("SDI Bulk Loads Formatting Tool")
root.geometry("800x400")

#TODO: Change this filepath to something that makes sense
ico= Image.open("V:\\Budget\\AutoQuotes Budget Script\\SDI Logo.jpg")
photo = ImageTk.PhotoImage(ico)
root.wm_iconphoto(False, photo)

errorFrame = tk.Frame(root)
errorFrame.pack(side=tk.BOTTOM)
errorMsg = tk.Label(errorFrame, text="")
errorMsg.pack(pady=50)

def formatFile():
    if inputFilepath == '':
        errorMsg.config(text= "Error: No input file selected")
        return
    if outputFilepath =='':
        errorMsg.config(text="Error: No output file selected")
        return
    errorMsg.config(text="Something Broke...")

    wb = opx.load_workbook(inputFilepath)
    sheet = wb.active

#directly copy and paste(formatted)
#Save indexes of amp, kw, gph, btu, cfm(exh & supply), and volts (for formatting, processing, or validating)

    #search for headers

    #Fill 2D array with info

    #Create new xlsx

    #Fill xlsx with info
    
    #TODO: Read xl file, format new xl file
    pass

frame = tk.Frame(root)
frame.pack(padx=40, pady=40)

fileFrame = tk.Frame(root)
fileFrame.pack()

bottomFrame = tk.Frame(root)
bottomFrame.pack()

in_text = tk.Label(fileFrame, text="The input file is: " + inputFilepath)
in_text.pack(side=tk.TOP)

out_text = tk.Label(fileFrame, text = "The output folder is: " + outputFilepath)
out_text.pack(side=tk.TOP)

format_button = tk.Button(bottomFrame, text="format file", command=formatFile)
format_button.pack(padx=10, pady=30, side=tk.BOTTOM)

def getFilepath():
    global inputFilepath
    inputFilepath = filedialog.askopenfilename(filetypes=(("Micorsoft Excel Worksheet", "*.xlsx"),))
    in_text.config(text="The input file is: " + inputFilepath)

def getOutputFolder():
    global outputFilepath
    outputFilepath = filedialog.askdirectory()
    out_text.config(text="The output folder is: " + outputFilepath)

in_file = tk.Button(frame, text="select input file", command=getFilepath)
in_file.pack(padx=10, pady=15, side=tk.LEFT)

out_folder = tk.Button(frame, text="select output folder", command=getOutputFolder)
out_folder.pack(padx=10, pady=15, side=tk.LEFT)

root.mainloop()
    
