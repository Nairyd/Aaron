print ("Aaron\n\nExodus 2,16 'Aaron wird für dich zum Volk sprechen. Es ist so, als ob du durch ihn sprichst. Und er wird deine Botschaften weitergeben, so wie ein Prophet meine.'\n\n")

# Die Idee: Anhand der Daten aus Infoinput.xlsx wird ein Dokument zusammengestellt.
# Als Grundlage für das Dokument stehen mehrere docx Dokumente zur Verfügung (Bricks)
#
#


import os
import datetime
import random

import openpyxl                 # to work with excel
from openpyxl import load_workbook
from docx import Document       # to work with Document
from pathlib import Path


path = "."
document = Document()

# Vars to check ...
lname = "NACHNAME"
sname = "VORNAME"
sort = "STERBEORT"
gdate = "GEBURTSDATUM"
sdate = "STERBEDATUM"
lalter = "LEBENSALTER"
bvers = "BIBELVERS"
tmotiv = "TRAUERMOTIV"
# Erstelle eine placeholder Liste, von allen placeholdern, die in den Bricks vorhanden sind um später lplaceh durch lvars zu ersetzen
lplaceh = [lname,sname,sort,gdate,sdate,lalter,bvers,tmotiv]


def getTheVars():       #export variables from Excel?           könnte ich wohl auch als funktion machen, so nach dem motto ... if string value in spalte ?? == var dann geh eine spalte weiter und hol dir den Wert ....
    print("This are all the Infos we got:")
    global path
    wb = load_workbook(path+"\\Infoinput.xlsx")          # am besten baue ich hier direkt ein, dass er auf die excel tabelle im gleichen Ordner schaut ....
    #name= #!/usr/bin/python
    ws = wb.active
    global lname
    lname = (ws["C2"].value)
    global sname
    sname = (ws["D2"].value)
    print("Name:            ",sname," ",lname)
    global sort
    sort = (ws["C7"].value)
    print("Sterbeort:       ",sort)
    global gdate
    gdate = datetime.datetime.strftime((ws["C4"].value), "%d/%b/%Y")        #convert date into string ... looks better -.-
    print("Geburtsdatum:    ",gdate)
    global sdate
    sdate = datetime.datetime.strftime((ws["C5"].value), "%d/%b/%Y")        #convert date into string ... looks better -.-
    print("Sterbedatum:     ",sdate)
    global lalter
    lalter = str(ws["C5"].value.year - ws["C4"].value.year - ((ws["C5"].value.month, ws["C5"].value.day) < (ws["C4"].value.month, ws["C4"].value.day)))
    print("Lebensalter:     ",lalter)
    global tmotiv
    tmotiv = (ws["C17"].value)
    print("Trauermotiv:     ",tmotiv)
    bvers = choseRightBrick("\\Bibelverse\\",tmotiv)
    print("Bibelvers:\n",bvers,"\n")
    global lvars        # creating a list of all the vars
    lvars = [lname,sname,sort,gdate,sdate,lalter,bvers,tmotiv]            #put new vars here ... also put them in the lplaceh list ....

def paraMove(output_doc_name, paragraph):
    output_para = output_doc_name.add_paragraph()
    for run in paragraph.runs:
        output_run = output_para.add_run(run.text)
        output_run.bold = run.bold                                      # Run's bold data
        output_run.italic = run.italic                                  # Run's italic data
        output_run.underline = run.underline                            # Run's underline data
        output_run.font.color.rgb = run.font.color.rgb                  # Run's color data
        output_run.style.name = run.style.name                          # Run's font data
        output_para.paragraph_format.alignment = paragraph.paragraph_format.alignment


def brickMove(doc):
    print("Brick Moved")
    global document
    input_doc = Document(path+doc)
    for para in input_doc.paragraphs:
        paraMove(document, para)


def fillVars():
    print("Filling the Vars")
    lpos=-1
    for var in lplaceh:
        lpos += 1                   #check list position
        for paragraph in document.paragraphs:
            if var in paragraph.text:
                paragraph.text = paragraph.text.replace(str(var),lvars[lpos])


def choseRightBrick(brickpath,parameter):              # letztlich muss ich irgendwo noch ne Funktion einbauen, die Selektieren kann
#    print("Chosing right Brick in ",brickpath," with the parameter: ",parameter)
    bricklist =[]
    brickpath = Path("."+"\\Bricks\\"+brickpath)
    for item in brickpath.iterdir():            # abfrage ob er ne excel tabelle durchgehen soll, oder eben verschiedene docx dateien
        if item.suffix == ".xlsx":
            lpos = 4                            #startwer ab dem es in der Excell tabelle einträge gibt
            excel = load_workbook(item)         #Workbook
            sheet = excel.active                #Aktive Tabellenliste/Seite
            for cell in sheet["B5":"B7"]:       #durchsuche die Spalte .. länge muss noch manuell angepasst werden
                lpos +=1                        #trakt die position in der Liste ... in diesem Fall die Zeile
                if str(cell[0].value) == parameter:
                    bricklist.append(sheet["C"+str(lpos)].value)
    brick = random.choice(bricklist)
    return brick

#    if item in blabla == itemtype (Excel ... dann ...)

#    elif item in blala == itemtype (docx ...)



def buildEverything():
    buildTheIntro()
    #buildAnsprache()
    #buildOuttro()
    #buildAmGrab()
    fillVars()
    document.save(lname+'.docx')
    print("Build Everything")


def buildTheIntro():
    print("\nBuilding the intro now ...")
    global path

    #Headline
    header = document.sections[0].header
    header.paragraphs[0].text = "Trauerfeier "+sname+" "+lname

    #Header
    document.add_heading("Trauerfeier von "+sname+" "+lname, 0)

    #Votum
    document.add_heading("Votum", 1)
    brickMove("\\Bricks\\Votum\\Votum.docx")            # ich muss die datei weglassen denk ich und die Datei mit den Random sachen auswählen????

    #Begrüßung
    document.add_heading("Begrüßung", 1)
    brickMove("\\Bricks\\Begrüßung\\Begrüßung1.docx")

    #Mögliches Lied
    document.add_heading("Lied:", 1)

    #Eingangsgebet
    document.add_heading("Eingangsgebet", 1)

    #Psalm
    document.add_heading("Psalm", 1)


    print("Intro finished")



try:

    getTheVars()

    buildEverything()



finally:
    input("\nhats geklappt?")




# ------ ToDo:  --------------------------------------------------------------------------------------------
#   x  Platzhalter ersetzen im Abschlusstext ... hoffe das klappt
#   o  Möglichkeit zur Random auswahl von Bricks, bzw. zur geordneten Auswahl
#   o  Formatierung überarbeiten ...
#
#   o  Weitere Texte erstellen ... (nicht so wichtig)
#
#
#
