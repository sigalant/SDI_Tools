import openpyxl as opx
import tkinter as tk
from tkinter import filedialog

inputFilepath = ""
outputFilepath = ""

def getFilepath():
    global inputFilepath
    inputFilepath = filedialog.askopenfilename(filetypes = (("Microsoft Excel Worksheet", "*.xlsx"),))

def getOutputFolder():
    global outputFilepath 
    inputFilepath = filedialog.askdirectory()
    
def formatFile():
    
    if inputFilepath == "":
        print("No input file selected")
    if outputFilepath == "":
        print("No output folder selected")

    #Open Excel File TODO:Generalize for multiple budgets from folder, or create a GUI with file selection
    wb = opx.load_workbook(inputFilepath)
    sheet = wb.active

    #Search for headers
    filterList = [0,0,0,0,0,0,0] #To hold column# for [ItemNo, Qty, Category, Sell, Remarks, Specs, Model] 

    for cell in sheet[1]:
        if type(cell.value) is not str:
            continue
        if cell.value.lower() == "ItemNo".lower():
            filterList[0] = cell.column
        if cell.value.lower() == "Qty".lower():
            filterList[1] = cell.column
        if cell.value.lower() == "Category".lower():
            filterList[2] = cell.column
        if cell.value.lower() == "Sell".lower():
            filterList[3] = cell.column
        if cell.value.lower() == "Remarks".lower():
            filterList[4] = cell.column
        if cell.value.lower() == "Spec".lower():
            filterList[5] = cell.column
        if cell.value.lower() == "Model".lower():
            filterList[6] = cell.column
    print(filterList)

    #Check for missing columns
    try:
        assert (0 not in filterList), "Column may be missing"
    except Exception as e:
        print(e)

    #Fill 2D Array (Maybe linkedlist/Hash) with information

    data = []
    #data = [["Item", "Qty", "Description", "Model", "Unit Cost", "Total", "Remarks"]]

    for i in range(2, len(sheet["A"])+1):
        rowData = []
    
        for item in filterList:
            try:
                assert (item != 0)
                rowData.append(sheet[i][item-1].value)
            except Exception:
                rowData.append("Column Missing")
        data.append(rowData)
        print(rowData)



    #Create new Excel File

    newFile = "./AQ_Budgets/FormattedBudget.xlsx"

    wbNew = opx.Workbook()


    #TODO:Fill information
    
    sheetNew = wbNew.active

    sheetNew.sheet_view.showGridLines= False

    sheetNew.row_dimensions[1].height = 30
    sheetNew.row_dimensions[2].height = 30.75
    sheetNew.row_dimensions[3].height = 23.25
    sheetNew.row_dimensions[4].height = 27
    sheetNew.row_dimensions[5].height = 30
    sheetNew.row_dimensions[6].height = 15.75
    sheetNew.row_dimensions[7].height = 18.75

    sheetNew.column_dimensions['A'].width = 10
    sheetNew.column_dimensions['B'].width = 13.3
    sheetNew.column_dimensions['C'].width = 30.1
    sheetNew.column_dimensions['D'].width = 28.9
    sheetNew.column_dimensions['E'].width = 15.9
    sheetNew.column_dimensions['F'].width = 19
    sheetNew.column_dimensions['G'].width = 3.7
    sheetNew.column_dimensions['H'].width = 64.9


    headerBorder = opx.styles.borders.Border(top=opx.styles.borders.Side(style='thick', color='80002060'), bottom=opx.styles.borders.Side(style='thick'))

    sheetNew['A2'] = "Title"
    sheetNew['A2'].font = opx.styles.Font(size=24, bold=True)

    sheetNew['A3'] = "Date"
    sheetNew['A3'].font = opx.styles.Font(size=18, bold=True) 

    sheetNew['A5'] = "Item"
    sheetNew['A5'].border = headerBorder
    sheetNew['A5'].alignment = opx.styles.Alignment(horizontal = 'left', vertical = 'center')

    sheetNew['B5'] = "Qty"
    sheetNew['B5'].border = headerBorder
    sheetNew['B5'].alignment = opx.styles.Alignment(horizontal = 'center', vertical = 'center')

    sheetNew['C5'] = "Description"
    sheetNew['C5'].border = headerBorder
    sheetNew['C5'].alignment = opx.styles.Alignment(horizontal = 'left', vertical = 'center')

    sheetNew['D5'] = "Model"
    sheetNew['D5'].border = headerBorder
    sheetNew['D5'].alignment = opx.styles.Alignment(horizontal = 'left', vertical = 'center')

    sheetNew['E5'] = "Unit Cost"
    sheetNew['E5'].border = headerBorder
    sheetNew['E5'].alignment = opx.styles.Alignment(horizontal = 'right', vertical = 'center')

    sheetNew['F5'] = "Total"
    sheetNew['F5'].border = headerBorder
    sheetNew['F5'].alignment = opx.styles.Alignment(horizontal = 'right', vertical = 'center')

    sheetNew['H5'] = "Remarks"
    sheetNew['H5'].border = headerBorder
    sheetNew['H5'].alignment = opx.styles.Alignment(horizontal = 'left', vertical = 'center')

    rowNum = 6

    for i in range(len(data)):
        if data[i][0] == None:
            rowNum = rowNum+1
            c = "A"+str(rowNum)
            sheetNew[c] = data[i][5]
            sheetNew[c].font = opx.styles.Font(size = 14, color = 'FFFFFF', bold = True)
            sheetNew[c].fill = opx.styles.PatternFill(fgColor="002060", fill_type="solid")
            sheetNew["B"+str(rowNum)].fill = opx.styles.PatternFill(fgColor = "002060", fill_type="solid")
            sheetNew["C"+str(rowNum)].fill = opx.styles.PatternFill(fgColor = "002060", fill_type="solid")
            sheetNew["D"+str(rowNum)].fill = opx.styles.PatternFill(fgColor = "002060", fill_type="solid")
            sheetNew["E"+str(rowNum)].fill = opx.styles.PatternFill(fgColor = "002060", fill_type="solid")
            sheetNew["F"+str(rowNum)].fill = opx.styles.PatternFill(fgColor = "002060", fill_type="solid")
            sheetNew["H"+str(rowNum)].fill = opx.styles.PatternFill(fgColor = "002060", fill_type="solid")
            rowNum = rowNum+1
        else:
            sheetNew[("A"+str(rowNum))] = data[i][0]
            sheetNew[("B"+str(rowNum))] = data[i][1]
            sheetNew[("C"+str(rowNum))] = data[i][2]
            sheetNew[("D"+str(rowNum))] = data[i][5]
            sheetNew[("E"+str(rowNum))] = data[i][3]
            sheetNew[("E"+str(rowNum))].number_format = "$#,##0.00"
            try:
                sheetNew[("F"+str(rowNum))] = float(data[i][3])*float(data[i][1])
            except Exception:
                sheetNew[("F"+str(rowNum))] = data[i][3]
            sheetNew["F"+str(rowNum)].number_format = "$#,##0.00"
            sheetNew[("H"+str(rowNum))] = data[i][4]
            sheetNew[("H"+str(rowNum))].font = opx.styles.Font(color = '595959')
        rowNum = rowNum+1

    for i in range(6):
        sheetNew[rowNum][i].border = opx.styles.borders.Border(bottom=opx.styles.borders.Side(style='thick'))

    sheetNew[("E" + str(rowNum+1))] = "Equipment SubTotal"
    sheetNew['E'+str(rowNum+1)].alignment = opx.styles.Alignment(horizontal = 'right', vertical = 'center')

    sheetNew[("E" + str(rowNum+2))] = "Delivery, Installation, Set-in Place"
    sheetNew['E'+str(rowNum+2)].alignment = opx.styles.Alignment(horizontal = 'right', vertical = 'center')

    sheetNew[("E" + str(rowNum+3))] = "Total"
    sheetNew['E'+str(rowNum+3)].alignment = opx.styles.Alignment(horizontal = 'right', vertical = 'center')
    sheetNew['E'+str(rowNum+3)].font = opx.styles.Font(bold=True)

    sheetNew[("F" + str(rowNum+1))] = "=SUM(F7:F" + str(rowNum) + ")"
    sheetNew['F'+str(rowNum+1)].number_format = '$#,##0.00'
    sheetNew[("F" + str(rowNum+2))] = "=F" + str(rowNum+1) + "*.18"
    sheetNew['F'+str(rowNum+2)].number_format = '$#,##0.00'
    sheetNew[("F" + str(rowNum+3))] = "=SUM(F" + str(rowNum+1) + ":F" + str(rowNum+2) + ")"
    sheetNew['F'+str(rowNum+3)].number_format = '$#,##0.00'
    sheetNew['F'+str(rowNum+3)].font = opx.styles.Font(bold=True)

    for i in range(3,6):
        sheetNew[rowNum+2][i].border = opx.styles.borders.Border(bottom=opx.styles.borders.Side(style='thick'))

    sheetNew.merge_cells(start_row=rowNum+6, start_column=1, end_row=rowNum+6, end_column=3)
    sheetNew['A'+str(rowNum+6)] = "QUALIFICATIONS"
    sheetNew['A'+str(rowNum+6)].font = opx.styles.Font(bold=True, color='FFFFFF')
    sheetNew['A'+str(rowNum+6)].fill = opx.styles.PatternFill(fgColor = '002060', fill_type='solid')

    sheetNew['A'+str(rowNum+7)] = "1. Price does not include project contingency."
    sheetNew['A'+str(rowNum+7)].font = opx.styles.Font(color ='595959')

    sheetNew['A'+str(rowNum+8)] = "2. Price does not include sales tax or use taxes."
    sheetNew['A'+str(rowNum+8)].font = opx.styles.Font(color ='595959')

    sheetNew['A'+str(rowNum+9)] = "3. Price is good through _____________"
    sheetNew['A'+str(rowNum+9)].font = opx.styles.Font(color ='595959')

    wbNew.save(newFile)




root = tk.Tk()

root.geometry("400x400")

frame = tk.Frame(root)
frame.pack()

fileFrame = tk.Frame(root)
fileFrame.pack()

bottomFrame = tk.Frame(root)
bottomFrame.pack(side=tk.BOTTOM)

errorFrame = tk.Frame(root)
errorFrame.pack(side=tk.BOTTOM)

in_file = tk.Button(frame, text="input file", command=getFilepath)
in_file.pack(padx=10, pady=15, side=tk.LEFT)

out_folder = tk.Button(frame, text="output file", command=getOutputFolder)
out_folder.pack(padx=10, pady=15, side=tk.LEFT)

format_button = tk.Button(bottomFrame, text="format file", command=formatFile)
format_button.pack(padx=10, pady=15, side=tk.BOTTOM)

in_text = tk.Label(root, text="The input file is: " + inputFilepath)
in_text.pack(side=tk.LEFT)

out_text = tk.Label(root, text="The output folder is: " + outputFilepath)
out_text.pack(side=tk.RIGHT)

root.mainloop()
