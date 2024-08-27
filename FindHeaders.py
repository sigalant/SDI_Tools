import openpyxl as opx
from tkinter import filedialog


def FindHeaders(path):
    wb = opx.load_workbook(path, read_only=True)
    sheet = wb.active
    #print(len(sheet[1]))
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
            print(indexDict)
            break
    wb.close()
    return indexDict

def CopySheet(path):
    wb = opx.load_workbook(path, read_only=True)
    sheet = wb.active
    indexDict = FindHeaders(path)
    info = [[] for i in range(len(indexDict))]
    for row in sheet.rows:
        for head in indexDict:
            #print(str(indexDict[head])+":"+str(len(indexDict)))
            info[indexDict[head]].append(row[indexDict[head]].value)
    '''
    #printing to console
    for r in range(len(info[0])):
        row=[]
        for col in info:
            row.append(str(col[r]))
        print(", ".join(row))
    '''
    return indexDict, info

if __name__=="__main__":
        CopySheet(filedialog.askopenfilename(filetypes=(("Microsoft Excel Sheet","*.xlsx"),)))

