import openpyxl as opx
from tkinter import filedialog

def FindHeaders(sheet):
    if sheet.max_column < 2:
        sheet.reset_dimensions()
        sheet.calculate_dimension(force=True)
    temp = [i for i in range(len(sheet[1]))]
    indexDict = {}
    for row in sheet.rows:
        found = []
        for index in temp:
            val = row[index].value
            if val != None:
                while str(val).lower() in indexDict:
                    val = str(val) + "_"
                indexDict[str(val).lower()] = index
                found.append(index)
        for field in found:
            temp.remove(field)
        if not temp:
            break
    return indexDict

def CopySheet(sheet):
    indexDict = FindHeaders(sheet)
    info = [[] for i in range(len(indexDict))]
    for row in sheet.rows:
        for head in indexDict:
            info[indexDict[head]].append(row[indexDict[head]].value)

    return indexDict, info

if __name__=="__main__":
        CopySheet(filedialog.askopenfilename(filetypes=(("Microsoft Excel Sheet","*.xlsx"),)))

