#Imports
import LogErrors
import openpyxl as opx
import tkinter as tk
import re
import sys
from datetime import date
from tkinter import filedialog
from PIL import Image, ImageTk

#sys.excepthook = LogErrors.handle_exception

#Holds I/O filepaths
inputFilepath = ""
outputFilepath = ""



#Root window
root = tk.Tk()
root.report_callback_exception = LogErrors.handle_exception
root.title("SDI Bulk Loads Formatting Tool")
root.geometry("800x450")

#TODO: Change this filepath to something that makes sense
ico= Image.open("V:\\Budget\\AutoQuotes Budget Script\\SDI Logo.jpg")
photo = ImageTk.PhotoImage(ico)
root.wm_iconphoto(False, photo)

#For displaying errors to user
errorFrame = tk.Frame(root)
errorFrame.pack(side=tk.BOTTOM)
errorMsg = tk.Label(errorFrame, text="")
errorMsg.pack(pady=40)

def formatFile(voltList):
    #Stop if I/O paths not provided
    if inputFilepath == '':
        errorMsg.config(text= "Error: No input file selected")
        return
    if outputFilepath == '':
        errorMsg.config(text="Error: No output file selected")
        return
    #message returned if unhandled exception
    errorMsg.config(text="There was an error...")

    wb = opx.load_workbook(inputFilepath)
    sheet = wb.active

    temp = inputFilepath.split('/')
    filename = temp[(len(temp)-1)].split('.')[0] + "_formatted.xlsx"
    newFile = outputFilepath+"/"+filename
    wbNew=opx.Workbook()
    sheetNew=wbNew.active

    indexDict = {}
    sheetData = []
    summingIndexes = []

    #Copy all data to a 2D-List, and save indexes of important values
    for row in sheet.rows:
        rowData = []
        double = False
        for col in range(sheet.max_column):
            cellData = row[col].value
            
            
            #Get indexes from the first row
            match cellData:
                case str(x) if 'AMPS' in x and 'AMPS' not in indexDict:
                    indexDict['AMPS'] = col
                    summingIndexes.append(col)
                case str(x) if 'KW' in x and 'KW' not in indexDict:
                    indexDict['KW'] = col
                case str(x) if 'GPH' in x or 'LPH' in x and 'GPH' not in indexDict:
                    indexDict['GPH'] = col
                    summingIndexes.append(col)
                case str(x) if 'BTUS' in x and 'BTUS' not in indexDict:
                    indexDict['BTUS'] = col
                    summingIndexes.append(col)
                case str(x) if 'EXH CFM' in x and 'EXH CFM' not in indexDict:
                    indexDict['EXH CFM'] = col
                    summingIndexes.append(col)
                case str(x) if 'SUPPLY CFM' in x and 'SUPPLY CFM' not in indexDict:
                    indexDict['SUPPLY CFM'] = col
                    summingIndexes.append(col)
                case str(x) if 'VOLTS' in x and 'VOLTS' not in indexDict:
                    indexDict['VOLTS'] = col
                case str(x) if 'PH' in x and 'PH' not in indexDict:
                    indexDict['PH'] = col
                case str(x) if 'HEAT REJECTION' in x and 'HEAT REJECTION' not in indexDict:
                    indexDict['HEAT REJECTION'] = col
                    summingIndexes.append(col)
                case _:
                    pass

            #Remove commas for better number processing
            if type(cellData) == str:
                cellData = cellData.replace(',','')
                
            # Adjust for 'newlines' in cell
            if "_x000D_" in str(cellData):
                if (col in summingIndexes and '(' not in str(cellData)):
                    cellData = ' '.join(cellData.split("_x000D_"))
                    cellData = sum([float(s) for s in re.findall(r'\d+\.?\d*',cellData)])
                else:
                    cellData = cellData.split("_x000D_")[0]

            #Adjust for utility quantity x: (x)...A   (Should this work for other fields?)
            if ')' in str(cellData) and col == indexDict['AMPS']:
                cellData = [float(s) for s in re.findall(r'\d+\.?\d*',cellData)]
                cellData = cellData[0] * cellData[1]
                #cellData = str(float(cellData.split('(')[1].split(')')[0]) * float(cellData.split(')')[1].split('A')[0]))
            
            
            if type(cellData) == str and row[col].row>1 and col in summingIndexes:
                try:
                    if indexDict['GPH'] == col:
                        cellData = float(re.findall(r'\d+',cellData)[0])
                except Exception as e:
                    
                    print(cellData)
                    print("Error: " + str(e))
                    print("GPH column not found")
                    pass
                try:
                    if indexDict['BTUS'] == col:
                        cellData = re.findall(r'\d+',cellData)
                        if len(cellData) > 1:
                            cellData = int(cellData[0])*int(cellData[1])
                        else:
                            cellData = int(cellData[0])
                except Exception as e:
                    print(cellData)
                    print("Error: " + str(e))
                    print("BTUS column not found")
                    pass
                try:
                    if indexDict['EXH CFM'] == col:
                        cellData = float(re.findall(r'\d+',cellData)[0])
                except Exception as e:
                    print(cellData)
                    print("Error: " + str(e))
                    print("EXH CFM column not found")
                    pass
                try:
                    if indexDict['SUPPLY CFM'] == col:
                        cellData = float(re.findall(r'\d+',cellData)[0])
                except Exception as e:
                    print(cellData)
                    print("Error: " + str(e))
                    print("SUPPLY CFM column not found")
                try:
                    if indexDict['HEAT REJECTION'] == col:
                        cellData = float(re.findall(r'\d+',cellData)[0])
                except Exception as e:
                    print(cellData)
                    print("Error: " + str(e))
                    print("HEAT REJECTION column not found")
                    pass
                
            rowData.append(cellData)
                
        sheetData.append(rowData)

    #File Heading
    sheetNew.row_dimensions[1].height=30
    sheetNew.row_dimensions[2].height=30.75
    sheetNew.row_dimensions[3].height=23.25
    sheetNew.row_dimensions[4].height=27
    sheetNew.row_dimensions[5].height=30
    sheetNew.row_dimensions[6].height=15.75
    sheetNew.row_dimensions[7].height=5

    sheetNew.column_dimensions['A'].width = 10
    sheetNew.column_dimensions['B'].width = 5
    sheetNew.column_dimensions['C'].width = 40.1
    sheetNew.column_dimensions['D'].width = 10
    sheetNew.column_dimensions['E'].width = 10
    sheetNew.column_dimensions['F'].width = 10
    sheetNew.column_dimensions['G'].width = 10
    sheetNew.column_dimensions['H'].width = 10

    img = opx.drawing.image.Image("V:\\Budget\\Autoquotes Budget Script\\SDI Logo.jpg")
    img.height=40
    img.width=65
    sheetNew.add_image(img, "A1")

    sheetNew['A3'] = "________ Preliminary Utility Schedule"
    sheetNew['A3'].font = opx.styles.Font(size=24, bold=True)
    sheetNew['A4'] = str(date.today().strftime("%B %d, %Y"))
    sheetNew['A4'].font = opx.styles.Font(size=18, bold=True)
    sheetNew.append([''])
    #Add data to sheet
        
    sheetNew.append(sheetData[0])
    sheetNew.freeze_panes = sheetNew['C7']
    
    for row in range(1,len(sheetData)):
        #Check if area header
        if sheetData[row][0] != None and sheetData[row][0] == str(sheetData[row][0]).upper() and sheetData[row][1] == None:
            sheetNew.append([sheetData[row][0]])
            sheetNew[sheetNew.max_row][0].font = opx.styles.Font(size=14, color='FFFFFF', bold=True)
            for i in range(sheetNew.max_column):
                sheetNew[sheetNew.max_row][i].fill = opx.styles.PatternFill(fgColor="002060", fill_type="solid")
            continue
        
        #Add data
        if type(sheetData[row][indexDict['AMPS']]) == str and sheetData[row][indexDict['AMPS']] != '':
            sheetData[row][indexDict['AMPS']] = [float(s) for s in re.findall(r'\d+\.?\d*',sheetData[row][indexDict['AMPS']])][0]#float(str(sheetData[row][indexDict['AMPS']]).split('A')[0])
        if sheetData[row][indexDict['PH']] != None and sheetData[row][indexDict['PH']] != '':
            sheetData[row][indexDict['PH']] = float(sheetData[row][indexDict['PH']].split('PH')[0])
        if type(sheetData[row][indexDict['VOLTS']]) == str and sheetData[row][indexDict['VOLTS']] != '':
            if '/' in sheetData[row][indexDict['VOLTS']]:
                sheetData[row][indexDict['VOLTS']] = sheetData[row][indexDict['VOLTS']].split('/')[0]
            v = float(sheetData[row][indexDict['VOLTS']].split('V')[0])
            sheetData[row][indexDict['VOLTS']] = v
            try:
                sheetData[row][indexDict['KW']] = "=IF("+chr((indexDict['KW']-2)+65)+str(row+7)+">1,(1.732*"+chr((indexDict['KW']-3)+65)+str(row+7)+"*"+chr((indexDict['KW']-1)+65)+str(row+7)+")/1000,("+chr((indexDict['KW']-3)+65)+str(row+7)+"*"+chr((indexDict['KW']-1)+65)+str(row+7)+")/1000)"       
            except:
                pass
        sheetNew.append(sheetData[row])
        try:
            if float(sheetData[row][indexDict['VOLTS']]) not in voltList:
                sheetNew[row+7][indexDict['VOLTS']].fill = opx.styles.PatternFill(start_color='00FFFF00', end_color='00FFFF00', fill_type='solid')
        except:
            pass
        try:
            sheetNew[sheetNew.max_row][indexDict['AMPS']].number_format="0.0"
        except:
            pass
        try:
            sheetNew[sheetNew.max_row][indexDict['KW']].number_format="0.00"
        except:
            pass

    #Sums
    sheetNew.append([""])
    sheetNew.append(["Total"])
    sheetNew[sheetNew.max_row][0].font = opx.styles.Font(bold=True)

    try:
        sheetNew[sheetNew.max_row][indexDict['KW']].value = "=SUM("+chr(indexDict['KW']+65)+str(7)+":"+chr(indexDict['KW']+65)+str(sheetNew.max_row-1)+")"
        sheetNew[sheetNew.max_row][indexDict['KW']].number_format="#,##0"
        sheetNew[sheetNew.max_row][indexDict['KW']].font = opx.styles.Font(bold=True)
    except:
        pass

    try:
        sheetNew[sheetNew.max_row][indexDict['GPH']].value = "=SUM("+chr(indexDict['GPH']+65)+str(7)+":"+chr(indexDict['GPH']+65)+str(sheetNew.max_row-1)+")"
        sheetNew[sheetNew.max_row][indexDict['GPH']].number_format="#,##0"
        sheetNew[sheetNew.max_row][indexDict['GPH']].font = opx.styles.Font(bold=True)
    except:
        pass

    try:
        sheetNew[sheetNew.max_row][indexDict['BTUS']].value = "=SUM("+chr(indexDict['BTUS']+65)+str(7)+":"+chr(indexDict['BTUS']+65)+str(sheetNew.max_row-1)+")"
        sheetNew[sheetNew.max_row][indexDict['BTUS']].number_format="#,##0"
        sheetNew[sheetNew.max_row][indexDict['BTUS']].font = opx.styles.Font(bold=True)
    except:
        pass

    try:
        sheetNew[sheetNew.max_row][indexDict['EXH CFM']].value = "=SUM("+chr(indexDict['EXH CFM']+65)+str(7)+":"+chr(indexDict['EXH CFM']+65)+str(sheetNew.max_row-1)+")"
        sheetNew[sheetNew.max_row][indexDict['EXH CFM']].number_format="#,##0"
        sheetNew[sheetNew.max_row][indexDict['EXH CFM']].font = opx.styles.Font(bold=True)
    except:
        pass

    try:
        sheetNew[sheetNew.max_row][indexDict['SUPPLY CFM']].value = "=SUM("+chr(indexDict['SUPPLY CFM']+65)+str(7)+":"+chr(indexDict['SUPPLY CFM']+65)+str(sheetNew.max_row-1)+")"
        sheetNew[sheetNew.max_row][indexDict['SUPPLY CFM']].number_format="#,##0"
        sheetNew[sheetNew.max_row][indexDict['SUPPLY CFM']].font = opx.styles.Font(bold=True)
    except:
        pass

    try:
        sheetNew[sheetNew.max_row][indexDict['HEAT REJECTION']].value = "=SUM("+chr(indexDict['HEAT REJECTION']+65)+str(7)+":"+chr(indexDict['HEAT REJECTION']+65)+str(sheetNew.max_row-1)+")"
        sheetNew[sheetNew.max_row][indexDict['HEAT REJECTION']].number_format="#,##0"
        sheetNew[sheetNew.max_row][indexDict['HEAT REJECTION']].font = opx.styles.Font(bold=True)
    except:
        pass
    
    wbNew.save(newFile)


frame = tk.Frame(root)
frame.pack(padx=40, pady=20)

voltFrame = tk.Frame(root)
voltFrame.pack(pady=20)

fileFrame = tk.Frame(root)
fileFrame.pack()


bottomFrame = tk.Frame(root)
bottomFrame.pack()

in_text = tk.Label(fileFrame, text="The input file is: " + inputFilepath)
in_text.pack(side=tk.TOP)

out_text = tk.Label(fileFrame, text = "The output folder is: " + outputFilepath)
out_text.pack(side=tk.TOP)

voltLabel = tk.Label(voltFrame, text="Enter acceptable voltages:")
voltLabel.pack(side=tk.TOP)
voltEntries = []

def tab_pressed(event):
    if voltEntries.index(event.widget)+1<len(voltEntries):
        voltEntries[voltEntries.index(event.widget)+1].focus_set()
    else:
        voltEntries[0].focus_set()
    return "break"

def addVolt(butts):
    voltText = tk.Text(voltFrame, height = 1, width = 10)
    voltText.bind('<Tab>', tab_pressed)
    voltText.bind('<Return>', tab_pressed)
    voltText.pack(padx=10, pady=10, side = tk.LEFT)
    voltEntries.append(voltText)
    if(len(voltEntries) >= 5):
        butts[0].pack_forget()
    butts[1].pack(side=tk.RIGHT)
def removeVolt(butts):
    voltText = voltEntries.pop()
    voltText.destroy()
    if(len(voltEntries) <= 1):
        butts[1].pack_forget()
    butts[0].pack(side=tk.RIGHT)
voltButtons=[None,None]
addVoltButton = tk.Button(voltFrame, text = "+", command=lambda:addVolt(voltButtons))
removeVoltButton = tk.Button(voltFrame, text = "-", command=lambda:removeVolt(voltButtons))
voltButtons[0]=addVoltButton
voltButtons[1]=removeVoltButton



for i in range(3):
    voltText = tk.Text(voltFrame, height = 1, width = 10)
    voltText.bind('<Tab>', tab_pressed)
    voltText.bind('<Return>', tab_pressed)
    voltText.pack(padx=10, pady=10, side = tk.LEFT)
    voltEntries.append(voltText)

removeVoltButton.pack(padx=1,pady=1,side=tk.RIGHT)
addVoltButton.pack(padx=1,pady=1, side =tk.RIGHT)


def form():
    try:
        voltList = [float(t.get("1.0", 'end-1c')) for t in voltEntries]
    except:
        errorMsg.config(text="Error: All acceptable voltages should be numbers")
        return

    formatFile(voltList)

format_button = tk.Button(bottomFrame, text="format file", command=form)
format_button.pack(padx=10, pady=10, side=tk.BOTTOM)

def getFilepath():
    global inputFilepath
    inputFilepath = filedialog.askopenfilename(filetypes=(("Microsoft Excel Worksheet", "*.xlsx"),))
    in_text.config(text="The input file is: " + inputFilepath)

def getOutputFolder():
    global outputFilepath
    outputFilepath = filedialog.askdirectory()
    out_text.config(text="The output folder is: " + outputFilepath)

in_file = tk.Button(frame, text="select input file", command=getFilepath)
in_file.pack(padx=20, pady=15, side=tk.LEFT)

out_folder = tk.Button(frame, text="select output folder", command=getOutputFolder)
out_folder.pack(padx=20, pady=15, side=tk.LEFT)

root.mainloop()
    
