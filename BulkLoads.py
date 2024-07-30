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

    temp = inputFilepath.split('/')
    filename = temp[(len(temp)-1)].split('.')[0] + "_formatted.xlsx"
    newFile = outputFilepath+"/"+filename
    wbNew=opx.Workbook()
    sheetNew=wbNew.active

    indexDict = {}
    sheetData = []

    #Copy all data to a 2D-List, and save indexes of important values
    for row in sheet.rows:
        rowData = []
        for col in range(sheet.max_column):
            cellData = row[col].value
            #Get indexes from the first row
            match cellData:
                case 'AMPS':
                    indexDict['AMPS'] = col
                case 'KW':
                    indexDict['KW'] = col
                case 'GPH':
                    indexDict['GPH'] = col
                case 'BTUS':
                    indexDict['BTUS'] = col
                case 'EXH CFM':
                    indexDict['EXH CFM'] = col
                case 'SUPPLY CFM':
                    indexDict['SUPPLY CFM'] = col
                case 'VOLTS':
                    indexDict['VOLTS'] = col
                case 'PH':
                    indexDict['PH'] = col
                case _:
                    pass
            rowData.append(cellData)
        sheetData.append(rowData)
    #Format Sheet
    
    #Add data to sheet
    sheetNew.append(sheetData[0])
    for row in range(1,len(sheetData)):
        if row==0:
            row=1
            print("Printed a row ig")
        if sheetData[row][indexDict['AMPS']] != None and sheetData[row][indexDict['AMPS']] != '':
            sheetData[row][indexDict['AMPS']] = sheetData[row][indexDict['AMPS']].split('A')[0]
        if sheetData[row][indexDict['PH']] != None and sheetData[row][indexDict['PH']] != '':
            sheetData[row][indexDict['PH']] = sheetData[row][indexDict['PH']].split('PH')[0]
        if sheetData[row][indexDict['VOLTS']] != None and sheetData[row][indexDict['VOLTS']] != '':
            sheetData[row][indexDict['VOLTS']] = sheetData[row][indexDict['VOLTS']].split('V')[0]
        sheetData[row][indexDict['KW']] = "=IF("+chr((indexDict['KW']-2)+65)+str(row+1)+">1,(1.732*"+chr((indexDict['KW']-3)+65)+str(row+1)+"*"+chr((indexDict['KW']-1)+65)+str(row+1)+")/1000,("+chr((indexDict['KW']-3)+65)+str(row+1)+"*"+chr((indexDict['KW']-1)+65)+str(row+1)+")/1000)"
        #sheetData[row][indexDict['KW']].number_format="0.00"
        print(sheetData[row])
        sheetNew.append(sheetData[row])
        sheetNew[row][indexDict['KW']].number_format="0.00"
    wbNew.save(newFile)

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
    
