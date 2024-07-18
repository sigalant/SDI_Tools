import sqlite3
import docx as d
from os import listdir
from os.path import isfile, join

#desc manu model doc
#doc text

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
def addEntry(info, cur):
    
    #Ignore if entry already exists
    query = ("SELECT doc FROM item WHERE doc='" +str(info[3]) +"'")
    print(query)
    res = cur.execute(query)
    if(res.fetchone()): 
        print("Entry for " + info[0] + ":" + info[2] + " already exists.")
    else:
        #Add Entry
        cur.execute("INSERT INTO item VALUES ('" + str(info[0]) + "','" + str(info[1]) + "','" + str(info[2]) + "','" + str(info[3])+ "')")

        specText = parseSpecs(info[3])
        specText = specText.replace("'", "''").replace('"','""')
        
        query2 = ("INSERT INTO spec VALUES (\'" + str(info[3]) + "\',\'" + str(specText) + "\')")
        cur.execute(query2)

#Add entries from folder
def addEntries(folderPath, cur):
    onlyfiles = [f for f in listdir(folderPath) if isfile(join(folderPath, f))]
    for f in listdir(folderPath):
        f = f.replace("'", "''").replace('"','""')
        splitFile = f.split('_')
        if isfile(join(folderPath, f)) and len(splitFile) == 3 and '$' not in splitFile[0]:
            addEntry([splitFile[0], splitFile[1], splitFile[2].split('.docx')[0], join(folderPath, f)],cur)

#TODO: take entry and docx file, then change docx path and doc text in DB (or delete and recreate)
def ModifyEntry(entry, changes):
    pass

#Take entry and remove from DB
#change to test all fields?
def DeleteEntry(entry, cur):
    cur.execute("DELETE FROM item WHERE doc='" + entry + "'")
    cur.execute("DELETE FROM spec WHERE doc='" + entry + "'")

#Take argument for some field and return any entries that match 
def FindEntry(fields, cur):
    #[desc, manu, model, doc]
    conditionals = []
    if fields[0]:
        conditionals.append("desc='" + fields[0] + "'")
    if fields[1]:
        conditionals.append("manu='" + fields[1] + "'")
    if fields[2]:
        conditionals.append("model='" + fields[2] + "'")
    if fields[3]:
        conditionals.append("doc='" + fields[3] + "'")
    cond = " AND ".join(conditionals)
    return cur.execute("SELECT desc, manu, model, doc FROM item WHERE " + cond).fetchall()
    

#Re-read Specs '.docx' file and write contents to spec table
def UpdateSpecs(doc, cur):
    newText = parseSpecs(doc)
    cur.execute("UPDATE spec SET text='" + newText.replace("'","''").replace('"','""') + "' WHERE doc = '" + doc + "'")

    
if __name__ == '__main__':
    con = sqlite3.connect("SpecDB.db")
    cur = con.cursor()
    addEntries("V:\\Specs\\Specs Script\\Template Specs_Word Files", cur)
    con.commit()
    con.close()
