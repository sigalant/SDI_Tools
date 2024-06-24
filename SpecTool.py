import docx as d
import openpyxl as opx
import num2words as n2m

doc = d.Document()

doc_styles = doc.styles
header_style = doc_styles.add_style('Spec_Header', d.enum.style.WD_STYLE_TYPE.PARAGRAPH)
header_font = header_style.font
header_font.size = d.shared.Pt(10)
header_font.name = 'Univers LT Std 55'




wb = opx.load_workbook("./_ALL INCLUSIVE.xlsx", read_only=True)
sheet = wb.active

for row in sheet.rows:
    if row[0].value == None or row[0].value == 'NO' or row[0].value == 'EQUIPMENT':
        continue
    elif row[1].value == None:
        p = doc.add_paragraph('', style = 'Spec_Header')
        p.alignment = 1
        p.add_run(row[0].value + "\n").bold = True
    else:
        p = doc.add_paragraph('', style = 'Spec_Header')
        p.alignment = 0

        
        tab_stops = p.paragraph_format.tab_stops
        tab_stop = tab_stops.add_tab_stop(d.shared.Inches(1.31), d.enum.text.WD_TAB_ALIGNMENT.LEFT)

        tab_stop = tab_stops.add_tab_stop(d.shared.Inches(1.69))

        run = ""

        run = run + ("ITEM #" + str(row[0].value) + ":")
        run = run + ("\t" + str(row[4].value))
        
        if "SPARE NUMBER" in row[4].value:
            p.add_run(run)
            continue
        run = run + ("\nQuantity:\t")
        run = run + (n2m.num2words(row[1].value).capitalize() + " (" + str(row[1].value) + ")")
        
        if row[5].value != None and ("by vendor" in row[5].value.lower() or "by os&e" in row[5].value.lower() or "by general contractor" in row[5].value.lower()):
            run = run+ ("\nPertinent Data:\t")
            if(row[5].value == None):
                run = run + add_run("---")
            else:
                run = run + (str(row[5].value))
            p.add_run(run)
            continue

        run = run + ("\nManufacturer:\t")
        if(row[2].value == None):
            run = run + ("---")
        else:
            run = run + (str(row[2].value))

        run = run + ("\nModel No.:\t")
        if(row[3].value == None):
            run = run + ("---")
        else: 
            run = run + (str(row[3].value))

        run = run + ("\nPertinent Data:\t")
        if(row[5].value == None):
            run = run + ("---")
        else:
            run = run + (str(row[5].value))

        run = run + ("\nUtilities Req'd:\t")
        is_empty = True
        if(row[25].value != None and row[25].value[0] == "-" and row[20].value == None):
            run = run + (str(row[24].value + " drain recessed " + str(row[25].value)[1:]))
            is_empty = False
        if(row[49].value != None):
            if not is_empty:
                run = run + " ;"
            if "(" in str(row[49].value):
                run = run + str(row[49].value)[:3]+ " " + str(row[50].value) + "/" + str(row[51].value)+", " + str(row[49].value)[3:] 
            else:
                run = run +(str(row[50].value) + "/" + str(row[51].value)+ ", "+ str(row[49].value))
            if (row[52].value != None):
                run = run + " (" + str(row[52].value) + ")"
            is_empty = False
        cfm = [] 

        if row[9].value != None:
            values = row[9].value.split()
            for value in values:
                cfm.append(value[:4] + " CFM Exhaust")

        if row[12].value != None:
            values = row[12].value.split()
            for value in values:
                cfm.append(value[:4] + " CFM Supply")
        
        if cfm:
            if not is_empty:
                run = run + "; "
            run = run + ", ".join(cfm)
            is_empty = False

        tempList = []
        
        if row[20].value != None:
            tempList.append(str(row[20].value) + " CW")
        if row[21].value != None:
            tempList.append(str(row[21].value) + " HW")
        if row[23].value != None:
            tempList.append(str(row[23].value) + " IW")
        if row[24].value != None and tempList:
            tempList.append(str(row[24].value) + " DW")
            

        if not is_empty and tempList:
            run = run + "; " 
            
        run = run + (", ".join(tempList))

        if tempList:
            is_empty = False

        if row[27].value != None:
            if not is_empty:
                run = run + "; "
            run = run + row[27].value + " Gas @ " + str(row[28].value) + " BTU; " + str(row[29].value) + " WC"
            is_empty = False

        if row[31].value != None:
            if not is_empty:
                run = run + "; "
            else:
                is_empty = False
            run = run + row[31].value + " Chilled Water Supply, " + row[32].value + " Chilled Water Return"


        if is_empty:
            run = run +("---")
        p.add_run(run)


doc.save("temp.docx")
