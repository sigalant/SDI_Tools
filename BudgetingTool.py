from openpyxl import load_workbook

#Open Excel File TODO:Generalize for multiple budgets from folder, or create a GUI with file selection
wb = load_workbook("./AQ_Budgets/AQ Budget.xlsx")
sheet = wb.active

#Search for headers
filterList = [0,0,0,0,0,0] #To hold column# for [ItemNo, Qty, Category, Sell, Remarks, Specs] 

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
print(filterList)

#Check for missing columns
try:
    assert (0 not in filterList), "Column may be missing"
except Exception as e:
    print(e)


#Fill 2D Array (Maybe linkedlist/Hash) with information

data = [["Item", "Qty", "Description", "Model", "Unit Cost", "Total", "Remarks"]]

for i in range(2, len(sheet["A"])+1):
    rowData = []
    
    for item in filterList:
        try:
            assert (item != 0)
            rowData.append(sheet[i][item-1].value)
        except Exception:
            rowData.append("Column Missing")
    data.append(rowData)

        

#TODO:Create new Excel File

#TODO:Fill information


