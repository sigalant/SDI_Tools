#SDI Specs Formatting Tool
#Antonio Sigala
#06/26/2024

#Imports
import SpecDB as db
import sqlite3
#from tinydb import TinyDB, Query

import docx as d
import openpyxl as opx
import num2words as n2m
import re
import tkinter as tk
from tkinter import filedialog,ttk
from PIL import Image, ImageTk

from os import listdir
from os.path import isfile, join


#input/output locations
inputFilepath = ""
outputFilepath = ""
excelFilepath = ""

#tkinter root window
root = tk.Tk()

root.title("SDI Specs Formatting Tool")
root.geometry("800x500")
ico = Image.open("V:\\Specs\\Specs Script\\SDI Logo.jpg")
photo = ImageTk.PhotoImage(ico)
root.wm_iconphoto(False, photo)

def copySpecs(tempDocPath, p, highlight, cur):
    #COPY AND PASTE FROM ASSOCIATED DOC
    temp = d.Document(tempDocPath)
    fullText = []
    i = 0
    
    specText = cur.execute("SELECT text FROM spec WHERE doc='" + tempDocPath + "'").fetchone()[0].split("\n")
    
    while i<len(specText) and "Utilities" not in specText[i]:
        i = i + 1
    if (len(("\n".join(specText[i+1:])).split('~')) - 1)%3 != 0:
        specText = ["\n".join(specText[i+1:])]
    else:
        specText = ("\n".join(specText[i+1:])).split('~')
    j=0
    while j < len(specText):
        if j%3 == 0:
            if(highlight):
                p.add_run(specText[j]).font.highlight_color = d.enum.text.WD_COLOR_INDEX.YELLOW
            else:
                p.add_run(specText[j])
        else:
            r=int(specText[j+1][:2], 16)
            g=int(specText[j+1][2:4],16)
            b=int(specText[j+1][4:6],16)
            
            redRun = p.add_run(specText[j])
            redRun.font.color.rgb = d.shared.RGBColor(r,g,b)
            if highlight:
                redRun.font.highlight_color = d.enum.text.WD_COLOR_INDEX.YELLOW
            j= j+1
        j=j+1
    '''
    #Add everything after the header
    while i< len(temp.paragraphs) and "Utilities" not in temp.paragraphs[i].text:
        i = i + 1
        
    #Go through each paragraph looking for alternately colored text
    for para in temp.paragraphs[i+1:]:
        p_runs = []
        addRuns = False
        for runS in para.runs:
            if runS.font.color.rgb != d.shared.RGBColor(0x00, 0x00, 0x00) and runS.font.color.rgb != None:
                if fullText:
                    if highlight:
                        p.add_run('\n'.join(fullText) + '\n').font.highlight_color = d.enum.text.WD_COLOR_INDEX.YELLOW
                    else:
                        p.add_run('\n'.join(fullText) + '\n')
                    fullText = []
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
    #break
    '''
def findSpecs(msgLabel):
    con = sqlite3.connect("SpecDB.db")
    cur = con.cursor()
    msgLabel.config(text="Error: Specs Not Found. Please check input file is correct.")
    try:
        wb = opx.load_workbook(inputFilepath, read_only=True)
    except Exception as e:
        print("Input File Not Found... Please Check Input Filepath")
        msgLabel.config(text="Error: Input File not found")
        return
    if outputFilepath == "":
        msgLabel.config(text="Error: Select Output Location")
        return
    wbNew = opx.Workbook()
    newSheet = wbNew.active
    sheet = wb.active
    headerIndexes = [-1,-1,-1,-1]
    
    #Find more permanent place for these files
    #onlyfiles = [f for f in listdir("V:\\Specs\\Specs Script\\Template Specs_Word Files") if isfile(join("V:\\Specs\\Specs Script\\Template Specs_Word Files", f))]        
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
                    msgLabel.config(text="Warning: One of the headers is missing from the input file")
            continue
        try:
            nothing = row[1].value == None or "spare" in row[4].value.lower()
        except:
            return
        #Skip if location header, spare number, existing item, or by OS&E/Manufacturer/etc.
        if row[1].value == None or "spare" in row[4].value.lower() or (row[5].value != None and ("by" in row[5].value.lower() or "exist" in row[5].value.lower())):
            continue
        #Collect Name, Manufacturer, and Model No. for finding/matching a Spec ".docx" file
        else:
            ambiguousModels = ["custom", "custom design"]
            specData = []
            #Manually fill field for custom fab
            if row[headerIndexes[0]+5].value != None and "CUSTOM FABRICATION" in row[headerIndexes[0]+5].value:
                specData = [row[headerIndexes[0]+4].value, "Custom Fabrication", ""]
            #Fill fields with Excel values
            else:
                specData = [row[headerIndexes[0]+4].value, row[headerIndexes[0]+2].value, str(row[headerIndexes[0]+3].value).replace('/', '-').replace('|','-')]

            #Add row to doc specific Excel file (Name, Manufacturer, Model No., Expected ".docx" filename)
            newSheet.append([specData[0], specData[1], specData[2], ''])#(str(specData[0])+"_"+ str(specData[1]) +"_"+ str(specData[2]))])
            rowIndex = newSheet.max_row #current row
            #filled = False #To avoid highlighting same row multiple times
            matches = []
            if specData[1] == "Custom Fabrication":
                matches = cur.execute("SELECT doc FROM item WHERE desc='" + str(specData[0]).replace("'","''").replace('"','""') + "' COLLATE NOCASE AND manu = 'Custom Fabrication' COLLATE NOCASE").fetchall()#db.itemTable.search((Query()['Description'].matches('(?i)'+str(specData[0])) & (Query()['Manufacturer'].matches('Custom Fabrication'))))
            else:
                matches = cur.execute("SELECT doc FROM item WHERE model='" + str(specData[2]).replace("'","''").replace('"','""') + "'").fetchall()#db.itemTable.search(Query()['Model_No.'].search(specData[2]))
            #if "22CG" in specData[2]:
                #print(matches)
            if matches:

                newSheet[rowIndex][3].value = "=HYPERLINK(\"[" + matches[0][0] + "]\",\""+ matches[0][0].split('\\')[len(matches[0][0].split('\\'))-1].split('.docx')[0] +"\")"
                for i in range(0,4):
                    newSheet[rowIndex][i].fill = noFill
                #print("Something is working...")
                if len(matches) > 1:
                    pass#print(matches)
            else:
                #Check partial matches
                if (specData[1] != "Custom Fabrication" and str(specData[2]).lower() not in ambiguousModels):
                    matches = cur.execute("SELECT doc FROM item WHERE model LIKE '%" + str(specData[2]).replace("'","''").replace('"','""') + "%' AND manu LIKE '%" + str(specData[1]).replace("'","''").replace('"','""') + "%'").fetchall()
                if not matches: 
                    matches = cur.execute("SELECT doc FROM item WHERE desc LIKE '%" + str(specData[0]).replace("'","''").replace('"','""')+ "%' AND manu LIKE '%"+ str(specData[1]).replace("'","''").replace('"','""') +"%'").fetchall()#db.itemTable.search((Query()['Description'].search('(?i)'+str(specData[0])) & (Query()['Manufacturer'].search('(?i)'+str(specData[1])))))
                else:
                    print(matches)
                if matches:
                    newSheet[rowIndex][3].value = "=HYPERLINK(\"[" + matches[0][0] + "]\",\""+ matches[0][0].split('\\')[len(matches[0][0].split('\\'))-1].split('.docx')[0] +"\")"
                    for i in range(0,4):
                        newSheet[rowIndex][i].fill = yellowFill
                else:
                    for i in range(0,4):
                        newSheet[rowIndex][i].fill = redFill
            
            '''
            for file in onlyfiles:
                #If model number and manufacturer match or is custom fab and name matches, row background is white
                if (specData[1] == "Custom Fabrication" and str(specData[0])+ "_Custom Fabrication" in file) or (str(specData[2]) in file and specData[2] != "" and str(specData[1]).lower() in file.lower() and specData[1] != ""):
                    newSheet[rowIndex][3].value= "=HYPERLINK(\"[V:\\Specs\\Specs Script\\Template Specs_Word Files\\" + str(file.split(".docx")[0]) + ".docx]\",\""+str(file.split(".docx")[0])+"\")" 
                    for i in range(0,4):
                        newSheet[rowIndex][i].fill = noFill
                    #print(file)
                    break
                #Elif manufacturer and name match, row background is yellow, and doc link is changed to the first (or last?) possible match
                elif str(specData[0]).lower() + "_" + str(specData[1]).lower() in file.lower():
                    newSheet[rowIndex][3].value= "=HYPERLINK(\"[V:\\Specs\\Specs Script\\Template Specs_Word Files\\" + str(file.split(".docx")[0]) + ".docx]\",\""+str(file.split(".docx")[0])+"\")" 
                    for i in range(0,4):
                        newSheet[rowIndex][i].fill = yellowFill
                    filled = True
                #Fill with red and remove ".docx" link on first non-match (will be overwritten if a match is found after the first non-match)
                elif not filled:
                    newSheet[rowIndex][3].value = ""

                    for i in range(0,4):
                        newSheet[rowIndex][i].fill = redFill
                    filled = True
            '''
    #Save new Specs Worksheet
    try:
        wbNew.save(outputFilepath+"\\SpecRefSheet.xlsx")
        global excelFilepath
        excelFilepath = outputFilepath+"/SpecRefSheet.xlsx"
    except Exception as e:
        print(e)
        msgLabel.config(text="Error: Cannot Save SpecRefSheet.xlsx while file is open")
        return
    msgLabel.config(text="Successfully Created Spec Ref Sheet")
        
def writeSpecs(msgLabel):

    global excelFilepath
    
    #Create doc and style
    doc = d.Document()

    doc_styles = doc.styles
    header_style = doc_styles.add_style('Spec_Header', d.enum.style.WD_STYLE_TYPE.PARAGRAPH)
    header_font = header_style.font
    header_font.size = d.shared.Pt(10)
    header_font.name = 'Univers LT Std 55'

    msgLabel.config(text="Error: Please confirm input file is correct")
    #Open Revit output
    try:
        wb = opx.load_workbook(inputFilepath, read_only=True)
    except:
        print("Input File Not Found... Please Check Input Filepath")
        msgLabel.config(text="Error: Input File not found")
        return
    if outputFilepath == '':
        msgLabel.config(text="Error: Select Output location")
        return
    sheet = wb.active

    #Open Specs reference file (optional?)
    try:
        wbr = opx.load_workbook(excelFilepath, read_only=True)
    except:
        #msgLabel.config(text="Warning: Spec Ref Sheet not found (Specs found manually)")
        print("The Excel File was not found... But that's OK")
        excelFilepath = ""
    headerIndexes = [-1,-1,-1,-1] #Holds index values for [Equipment, Ventilation, Plumbing, Electrical] Respectively from Revit output
    yellowFill = opx.styles.PatternFill(start_color = 'FFFF00', end_color = 'FFFF00', fill_type = 'solid')
    specRefs = [[],[],[],[],[]]  #Holds [Desc.[], Manufacturer[], Model#[], refFile[], exactMatch?[]] from xl spec ref file
    onlyfiles = []
    refSheet = None
    if excelFilepath:
        refSheet = wbr.active
        for row in refSheet.rows:
            specRefs[0].append(str(row[0].value).lower())
            specRefs[1].append(str(row[1].value).lower())
            specRefs[2].append(str(row[2].value).lower())
            #print(row[3].value)
            if row[3].value == None:
                specRefs[3].append("")
                specRefs[4].append(False)
            else:
                specRefs[3].append(str(row[3].value).split('\"')[3])
                if row[0].fill == yellowFill:
                    specRefs[4].append(False)
                else:
                    specRefs[4].append(True)
    else:
        onlyfiles = [f for f in listdir("V:\\Specs\\Specs Script\\Template Specs_Word Files") if isfile(join("V:\\Specs\\Specs Script\\Template Specs_Word Files", f))]

    #for l in specRefs:
        #print(l)
    #Iterate every row in Revit output sheet
    for row in sheet.rows:
        #Find header locations
        if row[0].value == None or row[0].value == 'NO' or row[0].value == 'EQUIPMENT':
            #If first row, get header indexes, or skip if no item information available
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
            #print(str(row[headerIndexes[0]].value))
            #Item Number and Description
            run = run + ("ITEM #" + str(row[headerIndexes[0]].value) + ":")
            run = run + ("\t" + str(row[headerIndexes[0]+4].value))
            try:
                row[headerIndexes[0]+5].value != None
                "SPARE NUMBER" in row[headerIndexes[0]+4].value
            except:
                return
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

            customFab = False
            
            #Check if custom fab
            if(row[headerIndexes[0]+5].value != None and "CUSTOM FABRICATION" in row[headerIndexes[0]+5].value):
                run = run + "Custom Fabrication"
                customFab=True
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
            if(row[headerIndexes[2]+9].value != None and str(row[headerIndexes[2]+9].value)[0] == "-" and row[headerIndexes[2]+4].value == None):
                run = run + (str(row[headerIndexes[2]+8].value) + " drain recessed " + str(row[headerIndexes[2]+9].value)[1:])
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
                run = run + str(row[headerIndexes[2]+11].value) + " Gas @ " + str(row[headerIndexes[2]+12].value) + " BTU; " + str(row[headerIndexes[2]+13].value) + " WC"
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

            
            ambiguousModels = ["custom", "custom design"] #Model No. that aren't specific to a model

            #Make/find specs for item 

            
                
            
            #Existing Items Specs
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

                #Pretty much every possible spec for existing units without any logic to determine what's right (all red text instead)
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
                
            #If excel sheet is not provided, search through docs for a match
            elif excelFilepath == "":
                con = sqlite3.connect("specDB.db")
                cur = con.cursor()
                specData = []
                #Manually fill field for custom fab
                if row[headerIndexes[0]+5].value != None and "CUSTOM FABRICATION" in row[headerIndexes[0]+5].value:
                    specData = [row[headerIndexes[0]+4].value, "Custom Fabrication", ""]
                #Fill fields with Excel values
                else:
                    specData = [row[headerIndexes[0]+4].value, row[headerIndexes[0]+2].value, str(row[headerIndexes[0]+3].value).replace('/', '-').replace('|','-')]

                matches = []
                if specData[1] == "Custom Fabrication":
                    matches = cur.execute("SELECT doc FROM item WHERE desc='" + str(specData[0]).replace("'","''").replace('"','""') + "' COLLATE NOCASE AND manu = 'Custom Fabrication' COLLATE NOCASE").fetchall()#db.itemTable.search((Query()['Description'].matches('(?i)'+str(specData[0])) & (Query()['Manufacturer'].matches('Custom Fabrication'))))
                else:
                    matches = cur.execute("SELECT doc FROM item WHERE model='" + str(specData[2]).replace("'","''").replace('"','""') + "'").fetchall()#db.itemTable.search(Query()['Model_No.'].search(specData[2]))
                #if "22CG" in specData[2]:
                    #print(matches)
                if matches:
                    
                    copySpecs(matches[0][0], doc.add_paragraph('', style = 'Spec_Header'), False, cur)
                    
                    if len(matches) > 1:
                        pass#print(matches)
                else:
                    #Check partial matches
                    if (specData[1] != "Custom Fabrication" and str(specData[2]).lower() not in ambiguousModels):
                        matches = cur.execute("SELECT doc FROM item WHERE model LIKE '%" + str(specData[2]).replace("'","''").replace('"','""') + "%' AND manu LIKE '%" + str(specData[1]).replace("'","''").replace('"','""') + "%'").fetchall()
                    if not matches: 
                        matches = cur.execute("SELECT doc FROM item WHERE desc LIKE '%" + str(specData[0]).replace("'","''").replace('"','""')+ "%' AND manu LIKE '%"+ str(specData[1]).replace("'","''").replace('"','""') +"%'").fetchall()#db.itemTable.search((Query()['Description'].search('(?i)'+str(specData[0])) & (Query()['Manufacturer'].search('(?i)'+str(specData[1])))))
                    if matches:
                        copySpecs(matches[0][0], doc.add_paragraph('', style = 'Spec_Header'), True, cur)
                
            else:
                con = sqlite3.connect("specDB.db")
                cur = con.cursor()
                #if specs exist, copy and paste    
                if (row[headerIndexes[0]+3].value != None and str(row[headerIndexes[0]+3].value).lower() in specRefs[2] and specRefs[4][specRefs[2].index(str(row[headerIndexes[0]+3].value).lower())]
                      and str(row[headerIndexes[0]+3].value).lower() not in ambiguousModels) and specRefs[3][specRefs[2].index(str(row[headerIndexes[0]+3].value).lower())] != "":
                    
                    copySpecs("V:\\Specs\\Specs Script\\Template Specs_Word Files\\" + specRefs[3][specRefs[2].index(str(row[headerIndexes[0]+3].value).lower())] + ".docx", doc.add_paragraph('', style = 'Spec_Header'), False, cur)
                    #print("Specs found for item# " + str(row[0].value))
                    
                # If Manufacturer and Desc. match, copy and highlight specs
                elif (row[headerIndexes[0]+4].value != None and row[headerIndexes[0]+4].value.lower() in specRefs[0]):
                    manuf = ""
                    highlight = False
                    #Check if specs exist for matching Manufacturer and Desc., and turn on highlighting
                    if(row[headerIndexes[0]+2].value != None and str(row[headerIndexes[0]+2].value).lower() in specRefs[1]):
                        #print("Manufac")
                        manuf = str(row[headerIndexes[0]+2].value).lower()
                        highlight = True
                    #Check if specs for custom fabrication item exists
                    elif(row[headerIndexes[0]+5].value != None and "custom fabrication" in str(row[headerIndexes[0]+5].value).lower()):
                        #print("Custom Fab")
                        manuf = "custom fabrication"
                    #If neither, skip
                    else:
                        continue
                    #Find specs which match for given item
                    for index in [i for i,e in enumerate(specRefs[1]) if e == manuf]:
                        if index in [j for j,s in enumerate(specRefs[0]) if s == row[headerIndexes[0]+4].value.lower()]: #== specRefs[0].index(row[headerIndexes[0]+4].value.lower()):
                    
                            #COPY AND PASTE FROM ASSOCIATED DOC
                            if(specRefs[3][index] == ""):
                                continue
                            copySpecs("V:\\Specs\\Specs Script\\Template Specs_Word Files\\" + specRefs[3][index] + ".docx", doc.add_paragraph('', style = 'Spec_Header'),highlight,cur)
                            break                           
                    #print("Maybe Specs found for item# " + str(row[0].value))
    try:
        doc.save(outputFilepath+"\\Specs.docx")
    except:
        msgLabel.config(text="Error: Cannot save Specs.docx while file is open")
        return
    msgLabel.config(text="Successfully Created Specs Document")

#GUI

#Select Input File
def getFilepath(inputLabel):
    global inputFilepath
    inputFilepath = filedialog.askopenfilename(filetypes = (("Microsoft Excel Worksheet", "*.xlsx"),))
    inputLabel.config(text= "The input file is: " + inputFilepath)

#Select Output Location
def getOutputFolder(outputLabel):
    global outputFilepath 
    outputFilepath = filedialog.askdirectory()
    outputLabel.config(text= "The output folder is: " + outputFilepath)

#Select Excel File
def getExcelFile(xlLabel):
    global excelFilepath 
    excelFilepath = filedialog.askopenfilename(filetypes = (("Microsoft Excel Worksheet", "*.xlsx"),))
    xlLabel.config(text= "The ref sheet location is: " + excelFilepath)

def Search(tv, text, cur, var):
    fields = ["","","",""]
    match var.get():
        case "Model":
            fields[2] = text
        case "Description":
            fields[0] = text
        case "Manufacturer":
            fields[1] = text
        case _:
            print("How did you do that?????")
    if fields[0] or fields[1] or fields[2] or fields[3]:
        found = db.FindEntry(fields, cur)
        tv.delete(*tv.get_children())
        for item in found:
            tv.insert("", tk.END, text = item[0], values = (item[1], item[2], item[0]+"_"+item[1]+"_"+item[2]+".docx"))
    else:
        tv.delete(*tv.get_children())
        entries = cur.execute("SELECT * FROM item").fetchall()
        for entry in entries:
            tv.insert("", tk.END, text = entry[0], values = (entry[1], entry[2], entry[0]+"_"+entry[1]+"_"+entry[2]+".docx"))
    
        print("Text entry is empty")
#Temporary Placeholder Function
def Nothing():
    pass

def DBWindow():
    for widget in root.winfo_children():
        widget.destroy()
    root.geometry("1000x600")
    rightFrame = tk.Frame(root)
    rightFrame.pack(side="right")
    searchFrame= tk.Frame(rightFrame)
    searchFrame.pack(side="top")
    
    DBview = tk.Frame(rightFrame)
    DBview.pack(side = "bottom", padx = 20)
    treeview = ttk.Treeview(DBview, selectmode = 'browse', columns=("Manufacturer", "Model", "Word Doc"))
    treeview.heading("#0", text="Description")
    treeview.heading("Manufacturer", text="Manufacturer")
    treeview.heading("Model", text="Model")
    treeview.heading("Word Doc", text="Word Doc")
    con = sqlite3.connect("SpecDB.db")
    cur = con.cursor()
    entries = cur.execute("SELECT * FROM item").fetchall()
    for entry in entries:
        treeview.insert("", tk.END, text = entry[0], values = (entry[1], entry[2], entry[0]+"_"+entry[1]+"_"+entry[2]+".docx"))
    

    scrollBar = ttk.Scrollbar(DBview, orient="vertical", command = treeview.yview)
    scrollBar.pack(side='right', fill = 'y')
    treeview.pack(side = 'right', fill = 'y')
    treeview.configure(yscrollcommand = scrollBar.set)

    buttonFrame = tk.Frame(root)
    buttonFrame.pack(side="left")
    
    add = tk.Button(buttonFrame, text= "Add Entry", command=Nothing)
    add.pack(side="top", pady=10, padx=30, expand=True)
    modify = tk.Button(buttonFrame, text = "Modify Entry", command=Nothing)
    modify.pack(side="top", pady=10, padx=30, expand=True)
    delete = tk.Button(buttonFrame, text = "Delete Entry", command=Nothing)
    delete.pack(side="top", pady=10, padx=30, expand=True)
    update = tk.Button(buttonFrame, text = "Update Entries", command=Nothing)
    update.pack(side="top", pady=10, padx=30, expand=True)

    search = tk.Label(searchFrame, text= "Search for ")
    search.pack(side="left")
    T = tk.Text(searchFrame, height=1, width=15)
    T.pack(side="left")
    search2 = tk.Label(searchFrame, text=" in ")
    search2.pack(side="left")
    options = ["Description", "Manufacturer", "Model"]
    var = tk.StringVar(root)
    var.set("Model")
    drop = tk.OptionMenu(searchFrame, var, *options)
    drop.pack(side="left")
    submit = tk.Button(searchFrame, text="Search", command=lambda:Search(treeview, T.get("1.0", 'end-1c'), cur, var))
    submit.pack(side="left", padx = 20)
    
#Create window for choosing functionality
def selectionWindow():
    for widget in root.winfo_children():
        widget.destroy()

    frame = tk.Frame(root)
    frame.pack()
    
    findSpecButton = tk.Button(frame, text="Find Specs", command=DBWindow, height=2, width=15, bg= '#afafaf')
    findSpecButton.pack(padx=80,pady=150,side=tk.LEFT)

    writeSpecButton = tk.Button(frame, text="Write Specs", command=wsWindow, height=2, width=15, bg= '#afafaf')
    writeSpecButton.pack(padx=80,pady=100,side=tk.RIGHT)

#Create Window for Writing Specs Function
def wsWindow():
    for widget in root.winfo_children():
        widget.destroy()
        
    backButton = tk.Button(root, text="<-", command=selectionWindow, bg = '#dadada')
    backButton.grid(row=0, column=0, sticky = 'W', ipadx=15,ipady=5)

    padLabel = tk.Label(root, text="")
    padLabel.grid(row=1,column=1, padx=80, pady=30)

    messageLabel = tk.Label(root, text="")
    messageLabel.grid(row=7, column = 2, columnspan = 3)

    infoLabel = tk.Label(root, text="Write Specs Document")
    infoLabel.grid(row=1, column=3, sticky= 'N')

    inputLabel = tk.Label(root, wraplength = 250, text="The input file is: " + inputFilepath)
    inputLabel.grid(row=3, column=2, columnspan =3)

    inputButton = tk.Button(root, text="Select Input File", command=lambda:getFilepath(inputLabel))
    inputButton.grid(row=2, column=2, padx=10, pady=30)

    outputLabel = tk.Label(root, wraplength = 250, text="The output location is: " + outputFilepath)
    outputLabel.grid(row=4, column=2, columnspan =3)

    outputButton = tk.Button(root, text="Select Output Folder", command=lambda:getOutputFolder(outputLabel))
    outputButton.grid(row=2, column=4, padx=10, pady=10)

    xlLabel = tk.Label(root, wraplength = 250, text="The ref sheet location is: " + excelFilepath)
    xlLabel.grid(row=5, column=2, columnspan =3)
    
    xlButton = tk.Button(root, text="Select Ref Sheet\n(Optional)", command=lambda:getExcelFile(xlLabel))
    xlButton.grid(row=2, column=3, padx=10, pady=10)

    submitButton = tk.Button(root, text="Create Specs", command=lambda:writeSpecs(messageLabel))
    submitButton.grid(row=6, column = 3, pady = 20)

#Create Window for Finding Specs Function
def fsWindow():
    for widget in root.winfo_children():
        widget.destroy()

    
        
    backButton = tk.Button(root, text="<-", command=selectionWindow, bg = '#dadada')
    backButton.grid(row=0, column=0, sticky = 'W', ipadx=15,ipady=5)

    messageLabel = tk.Label(root, text="")
    messageLabel.grid(row=6, column = 2, columnspan = 3)

    padLabel = tk.Label(root, text="")
    padLabel.grid(row=1,column=1, padx=80, pady=30)

    infoLabel = tk.Label(root, text="Find Existing Specs")
    infoLabel.grid(row=1, column=3, sticky= 'N')

    inputButton = tk.Button(root, text="Select Input File", command=lambda:getFilepath(inputLabel))
    inputButton.grid(row=2, column=2, padx=10, pady=30)

    outputButton = tk.Button(root, text="Select Output Folder", command=lambda:getOutputFolder(outputLabel))
    outputButton.grid(row=2, column=4, padx=10, pady=10)

    inputLabel = tk.Label(root, wraplength = 250, text="The input file is: " + inputFilepath)
    inputLabel.grid(row=3, column=2, columnspan =3)

    outputLabel = tk.Label(root, wraplength = 250, text="The output location is: " + outputFilepath)
    outputLabel.grid(row=4, column=2, columnspan =3)

    submitButton = tk.Button(root, text="Find Specs", command=lambda:findSpecs(messageLabel))
    submitButton.grid(row=5, column = 3, pady = 20)


selectionWindow()

'''
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

'''
root.mainloop()
