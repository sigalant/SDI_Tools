import openpyxl as opx
import tkinter as tk
import re
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
        double = False
        for col in range(sheet.max_column):
            cellData = row[col].value
            
                
            #Get indexes from the first row
            match cellData:
                case str(x) if 'AMPS' in x:
                    indexDict['AMPS'] = col
                case str(x) if 'KW' in x:
                    indexDict['KW'] = col
                case str(x) if 'GPH' in x or 'LPH' in x:
                    indexDict['GPH'] = col
                case str(x) if 'BTUS' in x:
                    indexDict['BTUS'] = col
                case str(x) if 'EXH CFM' in x:
                    indexDict['EXH CFM'] = col
                case str(x) if 'SUPPLY CFM' in x:
                    indexDict['SUPPLY CFM'] = col
                case str(x) if 'VOLTS' in x:
                    indexDict['VOLTS'] = col
                case str(x) if 'PH' in x:
                    indexDict['PH'] = col
                case str(x) if 'HEAT REJECTION' in x:
                    indexDict['HEAT REJECTION'] = col
                case _:
                    pass
            # Adjust for 'newlines' in cell
            if "_x000D_" in str(cellData):
                if col == indexDict['AMPS'] and '(' not in str(cellData):
                    cellData = cellData.split("_x000D_")
                    cellData = str(float(cellData[0].split('A')[0]) + float(cellData[1].split('A')[0]))
                else:
                    double = True
                    cellData = cellData.split("_x000D_")[0]
            #Adjust for utility quantity x: (x)...A
            if ')' in str(cellData) and col == indexDict['AMPS']:
                cellData = str(float(cellData.split('(')[1].split(')')[0]) * float(cellData.split(')')[1].split('A')[0]))
            if type(cellData) == str:
                cellData = cellData.replace(',','')
            if double and cellData != None:
                try:
                    if indexDict['GPH'] == col:
                        cellData = float(re.findall(r'\d+',cellData)[0])*2
                except:
                    pass
                try:
                    if indexDict['BTUS'] == col:
                        cellData = float(re.findall(r'\d+',cellData)[0])*2
                except:
                    pass
                try:
                    if indexDict['EXH CFM'] == col:
                        cellData = float(re.findall(r'\d+',cellData)[0])*2
                except:
                    pass
                try:
                    if indexDict['SUPPLY CFM'] == col:
                        cellData = float(re.findall(r'\d+',cellData)[0])*2
                except:
                    pass
                try:
                    if indexDict['HEAT REJECTION'] == col:
                        cellData = float(re.findall(r'\d+',cellData)[0])*2
                except:
                    pass
            elif cellData != None:
                try:
                    if indexDict['GPH'] == col:
                        cellData = float(re.findall(r'\d+',cellData)[0])
                except:
                    pass
                try:
                    if indexDict['BTUS'] == col:
                        cellData = float(re.findall(r'\d+',cellData)[0])
                except:
                    pass
                try:
                    if indexDict['EXH CFM'] == col:
                        cellData = float(re.findall(r'\d+',cellData)[0])
                except:
                    pass
                try:
                    if indexDict['SUPPLY CFM'] == col:
                        cellData = float(re.findall(r'\d+',cellData)[0])
                except:
                    pass
                try:
                    if indexDict['HEAT REJECTION'] == col:
                        cellData = float(re.findall(r'\d+',cellData)[0])
                except:
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
        if sheetData[row][indexDict['AMPS']] != None and sheetData[row][indexDict['AMPS']] != '':
            print(sheetData[row][indexDict['AMPS']])
            sheetData[row][indexDict['AMPS']] = float(str(sheetData[row][indexDict['AMPS']]).split('A')[0])
        if sheetData[row][indexDict['PH']] != None and sheetData[row][indexDict['PH']] != '':
            sheetData[row][indexDict['PH']] = float(sheetData[row][indexDict['PH']].split('PH')[0])
        if sheetData[row][indexDict['VOLTS']] != None and sheetData[row][indexDict['VOLTS']] != '':
            if '/' in sheetData[row][indexDict['VOLTS']]:
                sheetData[row][indexDict['VOLTS']] = sheetData[row][indexDict['VOLTS']].split('/')[0]
            sheetData[row][indexDict['VOLTS']] = float(sheetData[row][indexDict['VOLTS']].split('V')[0])
            try:
                sheetData[row][indexDict['KW']] = "=IF("+chr((indexDict['KW']-2)+65)+str(row+7)+">1,(1.732*"+chr((indexDict['KW']-3)+65)+str(row+7)+"*"+chr((indexDict['KW']-1)+65)+str(row+7)+")/1000,("+chr((indexDict['KW']-3)+65)+str(row+7)+"*"+chr((indexDict['KW']-1)+65)+str(row+7)+")/1000)"       
            except:
                pass
        sheetNew.append(sheetData[row])
        sheetNew[sheetNew.max_row][indexDict['AMPS']].number_format="0.0"
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
    
