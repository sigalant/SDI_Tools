#SDI Specs Formatting Tool
#Antonio Sigala
#06/26/2024

#Imports
import docx as d
import openpyxl as opx
import num2words as n2m
import re
import tkinter as tk

from os import listdir
from os.path import isfile, join


#input/output locations
inputFilepath = ""
outputFilepath = ""

#tkinter root window
root = tk.Tk()

root.title("SDI Specs Formatting Tool")
root.geometry("800x400")

def findSpecs():
    wb = opx.load_workbook("./_ALL INCLUSIVE.xlsx", read_only=True)
    wbNew = opx.Workbook()
    newSheet = wbNew.active
    sheet = wb.active
    headerIndexes = [-1,-1,-1,-1]
    onlyfiles = [f for f in listdir("V:\\Temp\\Antonio\\Template Specs_Word Files") if isfile(join("V:\\Temp\\Antonio\\Template Specs_Word Files", f))]        
    yellowFill = opx.styles.PatternFill(start_color = 'FFFF00', end_color = 'FFFF00', fill_type = 'solid')
    redFill = opx.styles.PatternFill(start_color = 'FF0000', end_color = 'FF0000', fill_type = 'solid')
    noFill = opx.styles.PatternFill(start_color = 'FFFFFF', end_color = 'FFFFFF', fill_type = 'solid')
    
    #Iterate every row in Revit output sheet
    for row in sheet.rows:
        #Find header locations
        if row[0].value == None or row[0].value == 'NO' or row[0].value == 'EQUIPMENT':  
            if row[0].value == 'EQUIPMENT':
                for i in range(len(row)):
                    match str(row[i].value):
                        case "EQUIPMENT":
                            headerIndexes[0] = i
                        case "VENTILATION":
                            headerIndexes[1] = i
                        case "PLUMBING":
                            headerIndexes[2] = i
                        case "ELECTRICAL":
                            headerIndexes[3] = i
                        case _:
                            pass
                if -1 in headerIndexes:
                    print("Header missing?!??!?!")
        elif row[1].value == None or "spare" in row[4].value.lower() or (row[5].value != None and ("by" in row[5].value.lower() or "exist" in row[5].value.lower())):
            continue
        else:
            
            specData = []
            if row[headerIndexes[0]+5].value != None and "CUSTOM" in row[headerIndexes[0]+5].value:
                specData = [row[headerIndexes[0]+4].value, "Custom Fabrication", ""]
            else:
                specData = [row[headerIndexes[0]+4].value, row[headerIndexes[0]+2].value, row[headerIndexes[0]+3].value]
            
            newSheet.append([specData[0], specData[1], specData[2], (str(specData[0])+"_"+ str(specData[1]) +"_"+ str(specData[2]))])
            rowIndex = newSheet.max_row
            filled = False
            for file in onlyfiles:
                if (specData[1] == "Custom Fabrication" and str(specData[0])+ "_Custom Fabrication" in file) or (str(specData[2]) in file and specData[2] != "" and str(specData[1]).lower() in file.lower() and specData[1] != ""):
                    newSheet[rowIndex][3].value= "=HYPERLINK(\"[V:\\Temp\\Antonio\\Template Specs_Word Files\\" + str(file.split(".docx")[0]) + ".docx]\",\""+str(file.split(".docx")[0])+"\")" 
                    for i in range(0,4):
                        newSheet[rowIndex][i].fill = noFill
                    print(file)
                    break
                elif str(specData[0]).lower() + "_" + str(specData[1]).lower() in file.lower():
                    newSheet[rowIndex][3].value= "=HYPERLINK(\"[V:\\Temp\\Antonio\\Template Specs_Word Files\\" + str(file.split(".docx")[0]) + ".docx]\",\""+str(file.split(".docx")[0])+"\")" 
                    for i in range(0,4):
                        newSheet[rowIndex][i].fill = yellowFill
                elif not filled:
                    newSheet[rowIndex][3].value = ""

                    for i in range(0,4):
                        newSheet[rowIndex][i].fill = redFill
                    filled = True
    
    wbNew.save("TestSheet.xlsx")

def writeSpecs():

    #Create doc and style
    doc = d.Document()

    doc_styles = doc.styles
    header_style = doc_styles.add_style('Spec_Header', d.enum.style.WD_STYLE_TYPE.PARAGRAPH)
    header_font = header_style.font
    header_font.size = d.shared.Pt(10)
    header_font.name = 'Univers LT Std 55'


    #Open Revit output
    wb = opx.load_workbook("./_ALL INCLUSIVE.xlsx", read_only=True)
    sheet = wb.active

    #Open Specs reference file 
    wbr = opx.load_workbook("./All AutoText Index2.xlsx", read_only=True)
    refSheet = wbr.active

    headerIndexes = [-1,-1,-1,-1] #Holds index values for [Equipment, Ventilation, Plumbing, Electrical] Respectively from Revit output

    specRefs = [[],[],[],[]]  #Holds [Desc., Manufacturer, Model#, refFile] from xl spec ref file
    for row in refSheet.rows:
        specRefs[0].append(str(row[1].value).lower())
        specRefs[1].append(str(row[3].value).lower())
        specRefs[2].append(str(row[4].value).lower())
        specRefs[3].append(row[10].value)

    #Iterate every row in Revit output sheet
    for row in sheet.rows:
        #Find header locations
        if row[0].value == None or row[0].value == 'NO' or row[0].value == 'EQUIPMENT':
            #If first row, get header indexes  
            if row[0].value == 'EQUIPMENT':
                for i in range(len(row)):
                    match str(row[i].value):
                        case "EQUIPMENT":
                            headerIndexes[0] = i
                        case "VENTILATION":
                            headerIndexes[1] = i
                        case "PLUMBING":
                            headerIndexes[2] = i
                        case "ELECTRICAL":
                            headerIndexes[3] = i
                        case _:
                            pass
                if -1 in headerIndexes:
                    print("Header missing?!??!?!")        
            

        #Add section header (location/area)
        elif row[1].value == None:
            p = doc.add_paragraph('', style = 'Spec_Header')
            p.alignment = 1
            p.add_run(row[0].value + "\n").bold = True
        
        #Create header with info from Revit output
        else:
            p = doc.add_paragraph('', style = 'Spec_Header')
            p.alignment = 0

        
            tab_stops = p.paragraph_format.tab_stops
            tab_stop = tab_stops.add_tab_stop(d.shared.Inches(1.31), d.enum.text.WD_TAB_ALIGNMENT.LEFT)
            tab_stop = tab_stops.add_tab_stop(d.shared.Inches(1.69))
        

            run = "" #To hold all text for each header
        
            #Item Number and Description
            run = run + ("ITEM #" + str(row[headerIndexes[0]].value) + ":")
            run = run + ("\t" + str(row[headerIndexes[0]+4].value))
        
            #Check if not in contract
            if row[headerIndexes[0]+5].value != None and ("by vendor" in row[headerIndexes[0]+5].value.lower() or "by os&e" in row[headerIndexes[0]+5].value.lower() or "by general contractor" in row[headerIndexes[0]+5].value.lower()):
                run = run + " (NOT IN CONTRACT)"
        
            #Check if existing equipment
            if row[headerIndexes[0]+5].value != None and "EXIST" in str(row[headerIndexes[0]+5].value):
                run = run + " (EXISTING EQUIPMENT)"
            
            #Highlight if shelving unit    
            if "shelv" in str(row[headerIndexes[0]+4].value).lower():
                r = p.add_run(run)
                run = ""
                r.font.highlight_color = d.enum.text.WD_COLOR_INDEX.YELLOW
        
            #Skip rest if Spare Number
            if "SPARE NUMBER" in row[headerIndexes[0]+4].value:
                p.add_run(run)
                continue

            #Quantity
            run = run + ("\nQuantity:\t")
            run = run + (n2m.num2words(row[headerIndexes[0]+1].value).capitalize() + " (" + str(row[headerIndexes[0]+1].value) + ")")
        
            #If not in contract, add pert. data and skip rest
            if row[headerIndexes[0]+5].value != None and ("by vendor" in row[headerIndexes[0]+5].value.lower() or "by os&e" in row[headerIndexes[0]+5].value.lower() or "by general contractor" in row[headerIndexes[0]+5].value.lower()):
                run = run+ ("\nPertinent Data:\t")
                if(row[headerIndexes[0]+5].value == None):
                    run = run + add_run("---")
                else:
                    run = run + (str(row[headerIndexes[0]+5].value))
                p.add_run(run)
                continue
        
            #Manufacturer
            run = run + ("\nManufacturer:\t")
            
            #Check if custom fab
            if(row[headerIndexes[0]+5].value != None and "CUSTOM FABRICATION" in row[headerIndexes[0]+5].value):
                run = run + "Custom Fabrication"
        
            else:
                if(row[headerIndexes[0]+2].value == None):
                    run = run + ("---")
                else:
                    run = run + (str(row[headerIndexes[0]+2].value))
        
            #Model Number
            run = run + ("\nModel No.:\t")
            if(row[headerIndexes[0]+3].value == None):
                run = run + ("---")
            else: 
                run = run + (str(row[headerIndexes[0]+3].value))


            #Pertinent Data
            run = run + ("\nPertinent Data:\t")
        
            #Remove "Custom Fabrication"
            if(row[headerIndexes[0]+5].value != None and "CUSTOM FABRICATION" in row[headerIndexes[0]+5].value):
                temp = "".join(row[headerIndexes[0]+5].value.split("CUSTOM FABRICATION"))
                if ", " in temp:
                    run = run + temp[2:]
                else:
                    run = run + "---"
            else:
                if(row[headerIndexes[0]+5].value == None):
                    run = run + ("---")
                else:
                    run = run + (str(row[headerIndexes[0]+5].value))
        
            #Utilities
            run = run + ("\nUtilities Req'd:\t")
            is_empty = True

        #Plumbing    
            if(row[headerIndexes[2]+9].value != None and row[headerIndexes[2]+9].value[0] == "-" and row[headerIndexes[2]+4].value == None):
                run = run + (str(row[headerIndexes[2]+8].value + " drain recessed " + str(row[headerIndexes[2]+9].value)[1:]))
                is_empty = False
        
        #Electrical
            if(row[headerIndexes[3]+3].value != None):
                a=row[headerIndexes[3]+3].value.split("_x000D_\n")
                v=row[headerIndexes[3]+4].value.split("_x000D_\n")
                ph=row[headerIndexes[3]+5].value.split("_x000D_\n")
                c=None
                if(row[headerIndexes[3]+6].value != None):
                    c=str(row[headerIndexes[3]+6].value).split("_x000D_\n")
                for i in range(len(a)):
                    if not is_empty:
                        run = run + "; "
                    if "(" in str(a[i]):
                        run = run + str(a[i])[:3]+ " " 
                
                    run = run + str(v[i]) + "/" + str(ph[i])+", "
                
                    if "(" in str(a[i]):
                        run = run + str(a[i])[3:] 
                    else:
                        run = run + str(a[i])       
            
                    if (c != None):
                        if i<len(c):
                            run = run + " (" + str(c[i]) + ")"
                    is_empty = False
            
        #Ventilation
            cfm = [] 

            if row[headerIndexes[1]+3].value != None:
                if type(row[headerIndexes[1]+3].value) == int:
                    cfm.append(str(row[headerIndexes[1]+3].value)+ " CFM Exhaust")
                else:
                    values = row[headerIndexes[1]+3].value.split()
                    for value in values:
                        cfm.append(value[:4] + " CFM Exhaust")

            if row[headerIndexes[1]+6].value != None:
                values = row[headerIndexes[1]+6].value.split()
                for value in values:
                    cfm.append(value[:4] + " CFM Supply")
        
            if cfm:
                if not is_empty:
                    run = run + "; "
                run = run + ", ".join(cfm)
                is_empty = False

            tempList = []
        
        #Plumbing    
            #Water
            if row[headerIndexes[2]+4].value != None:
                tempList.append(str(row[headerIndexes[2]+4].value) + " CW")
            if row[headerIndexes[2]+5].value != None:
                tempList.append(str(row[headerIndexes[2]+5].value) + " HW")
        
            #Waste
            if row[headerIndexes[2]+7].value != None:
                tempList.append(str(row[headerIndexes[2]+7].value) + " IW")
            if row[headerIndexes[2]+8].value != None and tempList:
                tempList.append(str(row[headerIndexes[2]+8].value) + " DW")
            

            if not is_empty and tempList:
                run = run + "; " 
            
            run = run + (", ".join(tempList))

            if tempList:
                is_empty = False

            #Gas
            if row[headerIndexes[2]+11].value != None:
                if not is_empty:
                    run = run + "; "
                run = run + row[headerIndexes[2]+11].value + " Gas @ " + str(row[headerIndexes[2]+12].value) + " BTU; " + str(row[headerIndexes[2]+13].value) + " WC"
                is_empty = False
    
            #Chilled Water
            if row[headerIndexes[2]+16].value != None:
                if not is_empty:
                    run = run + "; "
                else:
                    is_empty = False
                run = run + str(row[headerIndexes[2]+15].value) + " Chilled Water Supply, " + str(row[headerIndexes[2]+16].value) + " Chilled Water Return"

        
            if is_empty:
                run = run +("---")
        
            #Add Header to doc
            p.add_run(run)

            ambiguousModels = ["custom", "custom design"]
            if row[headerIndexes[0]+5].value != None and "EXIST" in str(row[headerIndexes[0]+5].value):
                            



                p = doc.add_paragraph('', style = 'Spec_Header')
                remaining = (row[headerIndexes[0]+5].value != None and "REMAIN" in str(row[headerIndexes[0]+5].value))           
                name = ""
                if row[headerIndexes[0]+4].value.lower() != None:
                    name = row[headerIndexes[0]+4].value.lower()
                specText = ""
                if(remaining):
                    specText = specText + "Remain in place existing unit as follows:\n"
                else:
                    specText = specText + "Relocate existing unit as follows:\n"
                specList = []
                if remaining:
                    specList.append("Existing unit is located in existing kitchen; unit should be thoroughly cleaned and remaing where shown on plan")
                else:
                    specList.append("Existing unit is located in existing kitchen; unit should be thoroughly cleaned and relocated where shown on plan")
                    specList.append("Schedule time with Owner for relocating unit")
                if "shelv" in name.lower():
                    specList.append("Replace shelves where corrosion spots appear; clean, sand, polish and repaint if necessary")
                elif "trash" not in name.lower() and "bin" not in name.lower():
                    specList.append("Repair where corrosion spots appear; clean, sand, polish and repain if necessary")
                i = 1
                for spec in specList:
                    specText = specText + str(i) + ".\t" + spec + "\n"
                    i = i+1
                p.add_run(specText)
                rr = p.add_run(str(i) + ".\tVerify all existing utility requirements and conditions\n"+ str(i+1) + ".\tThoroughly clean and sanitize unit\n")
                rr.font.color.rgb = d.shared.RGBColor(0xFF,0x00,0x00)
                p.add_run(str(i+2) + ".\tMust meet all applicable federal, state, and local laws, rules, regulations and codes")

                '''
                if row[headerIndexes[0]+5].value != None and "REMAIN" in str(row[headerIndexes[0]+5].value):           
                    p.add_run("Remain in place existing unit as follows:\n")
                else:
                    p.add_run("Relocate existing units as follows:\n")
                p.add_run("1.\tExisting unit is located in ")
                temp = p.add_run("existing kitchen; ")
                temp.font.color.rgb = d.shared.RGBColor(0xFF,0x00,0x00)
                p.add_run("unit should be thoroughly cleaned and ")
                if row[headerIndexes[0]+5].value != None and "REMAIN" in str(row[headerIndexes[0]+5].value):           
                    p.add_run("remain where shown on plan\n")
                else:
                    p.add_run("relocated where shown on plan\n2.\tSchedule time with Owner for relocating unit\n")
                temp.font.color.rgb = d.shared.RGBColor(0xFF,0x00,0x00)
                temp = p.add_run("3.\tRepair where corrosion spots appear; clean, sand, polish and repaint if necessary\n4.\tVerify all existing utility requirements and conditions.\n5.\tThoroughly clean and sanitize the unit\n")
                temp.font.color.rgb = d.shared.RGBColor(0xFF,0x00,0x00)
                p.add_run("6.\t Must meet all applicable federal, state, and local laws, rules, regulations, and codes")

                #print(str(row[0].value) + " is existing")
                '''


            #if specs exist, copy and paste (Unless existing)
            elif (row[headerIndexes[0]+3].value != None and str(row[headerIndexes[0]+3].value).lower() in specRefs[2] 
                and str(row[headerIndexes[0]+3].value).lower() not in ambiguousModels):
                #COPY AND PASTE FROM ASSOCIATED DOC
                temp = d.Document("V:\\Temp\\Antonio\\Template Specs_Word Files\\" 
                                  + specRefs[3][specRefs[2].index(str(row[headerIndexes[0]+3].value).lower())] + ".docx")
                fullText = []
                #print(specRefs[3][index])
                #====Item# with specs====#
                print("Specs found for item# " + str(row[0].value))
            

                p = doc.add_paragraph('', style = 'Spec_Header')
                i = 0
                while "Utilities" not in temp.paragraphs[i].text:
                    i = i + 1

                for para in temp.paragraphs[i+1:]:
                    p_runs = []
                    addRuns = False
                    beginning = True
                    for runS in para.runs:
                        if runS.font.color.rgb != d.shared.RGBColor(0x00, 0x00, 0x00) and runS.font.color.rgb != None:
                            if fullText:
                                p.add_run('\n'.join(fullText) + '\n')
                                fullText = []
                            if beginning:
                                beginning = False
                                #p.add_run('\n')
                            if p_runs:
                                p.add_run(''.join(p_runs))
    
                            p_runs = []
                            addRuns = True
                            redRun = p.add_run(runS.text)
                            redRun.font.color.rgb = runS.font.color.rgb
                        else:
                            p_runs.append(runS.text)
                    if addRuns:
                        p.add_run(''.join(p_runs))
                        #p.add_run('\n')
                        addRuns = False
                    else:
                        fullText.append(para.text)
                    if not fullText:
                        p.add_run('\n')
                    p_runs = []

                p.add_run('\n'.join(fullText))
                

            # If Manufacturer and Desc. match, copy and highlight specs
            elif (row[headerIndexes[0]+4].value != None and row[headerIndexes[0]+4].value.lower() in specRefs[0]):
                manuf = ""
                highlight =False
                if(row[headerIndexes[0]+2].value != None and str(row[headerIndexes[0]+2].value).lower() in specRefs[1]):
                    #print("Manufac")
                    manuf = str(row[headerIndexes[0]+2].value).lower()
                    highlight = True
                elif(row[headerIndexes[0]+5].value != None and "custom fabrication" in str(row[headerIndexes[0]+5].value).lower()):
                    #print("Custom Fab")
                    manuf = "custom fabrication"
                else:
                    continue
                for index in [i for i,e in enumerate(specRefs[1]) if e == manuf]:
                    if index in [j for j,s in enumerate(specRefs[0]) if s == row[headerIndexes[0]+4].value.lower()]: #== specRefs[0].index(row[headerIndexes[0]+4].value.lower()):
                
                        #COPY AND PASTE FROM ASSOCIATED DOC
                    
                        temp = d.Document("V:\\Temp\\Antonio\\Template Specs_Word Files\\" 
                                  + specRefs[3][index] + ".docx")
                        fullText = []
                        

                        p = doc.add_paragraph('', style = 'Spec_Header')
                        i = 0
                        while "Utilities" not in temp.paragraphs[i].text:
                            i = i + 1
                        for para in temp.paragraphs[i+1:]:
                            p_runs = []
                            addRuns = False
                            beginning = True
                            for runS in para.runs:
                                if runS.font.color.rgb != d.shared.RGBColor(0x00, 0x00, 0x00):
                                    if fullText:
                                        if highlight:
                                            p.add_run('\n'.join(fullText) + '\n').font.highlight_color = d.enum.text.WD_COLOR_INDEX.YELLOW
                                        else:
                                            p.add_run('\n'.join(fullText) + '\n')
                                        fullText = []
                                    if beginning:
                                        beginning = False
                                        #p.add_run('\n')
                                    if p_runs:
                                        if highlight:
                                            p.add_run(''.join(p_runs)).font.highlight_color = d.enum.text.WD_COLOR_INDEX.YELLOW
                                        else:
                                            p.add_run(''.join(p_runs))
                                    p_runs = []
                                    addRuns = True
                                    redRun = p.add_run(runS.text)
                                    redRun.font.color.rgb = runS.font.color.rgb
                                    if highlight:
                                        redRun.font.highlight_color = d.enum.text.WD_COLOR_INDEX.YELLOW
                                else:
                                    p_runs.append(runS.text)
                            if addRuns:
                                if highlight:
                                    p.add_run(''.join(p_runs)).font.highlight_color = d.enum.text.WD_COLOR_INDEX.YELLOW
                                else:
                                    p.add_run(''.join(p_runs))
                                #p.add_run('\n')
                                addRuns = False
                            else:
                                fullText.append(para.text)
                            if not fullText:
                                p.add_run('\n')
                            p_runs = []
                        if highlight:
                            p.add_run('\n'.join(fullText)).font.highlight_color = d.enum.text.WD_COLOR_INDEX.YELLOW
                        else:
                            p.add_run('\n'.join(fullText))
                        break
    
                                           
                #print("Maybe Specs found for item# " + str(row[0].value))
    doc.save("temp.docx")
frame = tk.Frame(root)
frame.pack(padx=40, pady=40)


fileFrame = tk.Frame(root)
fileFrame.pack()

bottomFrame = tk.Frame(root)
bottomFrame.pack()

in_text = tk.Label(fileFrame, text="The input file is: " + inputFilepath)
in_text.pack(side=tk.TOP)

out_text = tk.Label(fileFrame, text="The output folder is: " + outputFilepath)
out_text.pack(side=tk.TOP)

format_button = tk.Button(bottomFrame, text="format file", command=findSpecs)
format_button.pack(padx=10, pady=30, side=tk.BOTTOM)


def getFilepath():
    global inputFilepath
    inputFilepath = filedialog.askopenfilename(filetypes = (("Microsoft Excel Worksheet", "*.xlsx"),))
    in_text.config(text= "The input file is: " + inputFilepath)

def getOutputFolder():
    global outputFilepath 
    outputFilepath = filedialog.askdirectory()
    out_text.config(text= "The output folder is: " + outputFilepath)

in_file = tk.Button(frame, text="select input file", command=getFilepath)
in_file.pack(padx=10, pady=15, side=tk.LEFT)

out_folder = tk.Button(frame, text="select output folder", command=getOutputFolder)
out_folder.pack(padx=10, pady=15, side=tk.LEFT)


root.mainloop()
