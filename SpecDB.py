import sqlite3
import docx as d
from os import listdir
from os.path import isfile, join
import os
#desc manu model doc
#doc text modifiedtime

#Take Docx file and convert into String
def parseSpecs(file):
    doc=None
    try:
        doc = d.Document(file.strip())
    except:
        print("File not found at: " + file)
        return -1
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
    
    return docText

#Add new entry into DB
def addEntry(info, cur):
    for i in range(len(info)):
        info[i] = info[i].replace("'","''").replace('"','""')
    #Ignore if entry already exists
    query = ("SELECT doc FROM item WHERE doc='" +str(info[3]) +"'")
   
    res = cur.execute(query)
    if(res.fetchone()): 
        
        return False
    else:
        
        #Add Entry
        q=("INSERT INTO item VALUES ('" + str(info[0]) + "','" + str(info[1]) + "','" + str(info[2]) + "','" + str(info[3])+ "')")
        
        cur.execute(q)
        specText = parseSpecs(info[3].replace("''","'").replace('""','"'))
        specText = specText.replace("'", "''").replace('"','""')
        
        query2 = ("INSERT INTO spec VALUES (\'" + str(info[3]) + "\',\'" + str(specText) + "\',\'"+str(os.path.getmtime(info[3].replace('""','"').replace("''","'")))+"\')")
        cur.execute(query2)
        return True
       

#Add entries from folder
def addEntries(folderPath, cur):
    onlyfiles = [f for f in listdir(folderPath) if isfile(join(folderPath, f))]
    for f in listdir(folderPath):
        f = f.replace("'", "''").replace('"','""')
        splitFile = f.split('_')
        if isfile(join(folderPath, f)) and len(splitFile) == 3 and '$' not in splitFile[0]:
            addEntry([splitFile[0].strip(), splitFile[1].strip(), splitFile[2].split('.docx')[0].strip(), join(folderPath, f)],cur)

#Take an entry and a list of changes to make in the DB
def ModifyEntry(entry, changes, cur):
    
    changeList = []
    
    if changes[0]:
        changeList.append("desc='"+changes[0]+"'")
    if changes[1]:
        changeList.append("manu='"+changes[1]+"'")
    if changes[2]:
        changeList.append("model='"+changes[2]+"'")
    if changes[3]:
        changeList.append("doc='"+changes[3]+"'")
        cur.execute("UPDATE spec SET doc='"+str(changes[3])+"', text='"+str(parseSpecs(changes[3]).replace("'","''").replace('"','""')) +"', modTime='"+str(os.path.getmtime(changes[3]))+"' WHERE doc='"+entry[3]+"'")
    if not changeList:
        
        return
    query = "UPDATE item SET " + ", ".join(changeList) + " WHERE desc='"+str(entry[0])+"'AND manu='"+str(entry[1])+"'AND model='"+str(entry[2])+"'AND doc='"+entry[3]+"'"
    
    return cur.execute(query).rowcount
    

#Take entry and remove from DB
#change to test all fields?
def DeleteEntry(entry, cur):
    res = [0,0]
    res[0] = cur.execute("DELETE FROM item WHERE doc='" + entry + "'").rowcount
    res[1] = cur.execute("DELETE FROM spec WHERE doc='" + entry + "'").rowcount
    return res

#Take argument for some field and return any entries that match 
def FindEntry(fields, cur):
    #[desc, manu, model, doc]
    conditionals = []
    if fields[0]:
        conditionals.append("desc LIKE '%" + fields[0] + "%'")
    if fields[1]:
        conditionals.append("manu LIKE '%" + fields[1] + "%'")
    if fields[2]:
        conditionals.append("model LIKE '%" + fields[2] + "%'")
    if fields[3]:
        conditionals.append("doc LIKE '%" + fields[3] + "%'")
    cond = " AND ".join(conditionals)
    return cur.execute(("SELECT desc, manu, model, doc FROM item WHERE " + cond)).fetchall()
    

#Re-read Specs '.docx' file and write contents to spec table
def UpdateSpecs(cur):
    missing = []
    for item in cur.execute("SELECT * FROM spec").fetchall():
        try:
            if str(item[2]) == str(os.path.getmtime(item[0])):
                
                continue
        except:
            missing.append(item[0])
            continue
        newText = parseSpecs(item[0])
        if newText == -1:
            missing.append(item[0])
            continue
        
        cur.execute("UPDATE spec SET text='" + newText.replace("'","''").replace('"','""') + "', modTime='" + str(os.path.getmtime(item[0])) + "' WHERE doc = '" + item[0] + "'") 
    return missing

    
if __name__ == '__main__':
    con = sqlite3.connect("SpecDB.db")
    cur = con.cursor()
    addEntries("V:\\Specs\\Specs Script\\Template Specs_Word Files", cur)
    con.commit()
    con.close()
