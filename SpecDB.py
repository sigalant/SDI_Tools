from tinydb import TinyDB, Query, where
import docx as d
from os import listdir
from os.path import isfile, join

db = TinyDB('./SpecDB.json')
itemTable = db.table('Items')
specTable = db.table('Specs')

#Take Docx file and convert into String
def parseSpecs(file):
    doc = d.Document(file)
    docText = ''
    fullText = []
    for para in doc.paragraphs:
        p_runs = []
        addRuns = False
        for run in para.runs:
            if run.font.color.rgb != d.shared.RGBColor(0x00, 0x00, 0x00) and run.font.color.rgb != None:
                if fullText:
                    docText = docText + ('\n'.join(fullText) + '\n')
                    fullText = []
                if p_runs:
                    docText= docText + (''.join(p_runs))
                    p_runs = []
                addRuns = True
                docText = docText + ("~" + run.text + "~" + str(run.font.color.rgb) + '~')
            else:
                p_runs.append(run.text)
        if addRuns:
            docText = docText + (''.join(p_runs))
            addRuns = False
        else:
            fullText.append(para.text)
        if not fullText:
            docText = docText + ('\n')
        p_runs = []
    docText = docText + ('\n'.join(fullText))
    print(docText)
    return docText

#Add new entry into DB
def addEntry(info):
    global itemTable
    global specTable
    #Ignore if entry already exists
    if(itemTable.search(Query().fragment({'Word_Doc':info[3], 'Model_No.':info[2]}))):
        print("Entry for " + info[0] + ":" + info[2] + " already exists.")
    else:
        #Add Entry
        itemTable.insert({'Description': info[0], 'Manufacturer': info[1], 'Model_No.': info[2], 'Word_Doc': info[3]})
        specText = parseSpecs(info[3])
        specTable.insert({'Word_Doc': info[3], 'Spec_Text':specText})

#Add entries from folder
def addEntries(folderPath):
    onlyfiles = [f for f in listdir(folderPath) if isfile(join(folderPath, f))]
    for f in listdir(folderPath):
        splitFile = f.split('_')
        if isfile(join(folderPath, f)) and len(splitFile) == 3 and '$' not in splitFile[0]:
            addEntry([splitFile[0], splitFile[1], splitFile[2].split('.')[0], join(folderPath, f)])

#TODO: take entry and docx file, then change docx path and doc text in DB (or delete and recreate)
def ModifyEntry():
    pass

#TODO: take entry and remove from DB
def DeleteEntry():
    pass

#TODO: Take argument for some field and return any entries that match 
def FindEntry():
    pass

if __name__ == '__main__':
    addEntries()
