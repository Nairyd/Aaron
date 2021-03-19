print ("Aaron\n\nExodus 2,16 'Aaron wird für dich zum Volk sprechen. Es ist so, als ob du durch ihn sprichst. Und er wird deine Botschaften weitergeben, so wie ein Prophet meine.'\n\n")

# Die Idee: Anhand der Daten aus Infoinput.xlsx wird ein Dokument zusammengestellt.
# Als Grundlage für das Dokument stehen mehrere docx Dokumente zur Verfügung (Bricks)
#
#


import os
import datetime

import openpyxl                 # to work with excel
from openpyxl import load_workbook
from docx import Document       # to work with Document

path = "."
document = Document()

# Vars to check ...
lname = "NACHNAME"
sname = "VORNAME"
sort = "STERBEORT"
gdate = "GEBURTSDATUM"
sdate = "STERBEDATUM"
lalter = "LEBENSALTER"
# Erstelle eine placeholder Liste, von allen placeholdern, die in den Bricks vorhanden sind um später lplaceh durch lvars zu ersetzen
lplaceh = [lname,sname,sort,gdate,sdate,lalter]


def getTheVars():       #export variables from Excel?
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

    global lvars        # creating a list of all the vars
    lvars = [lname,sname,sort,gdate,sdate,lalter]            #put new vars here ... also put them in the lplaceh list ....

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


def choseRightBrick():              # letztlich muss ich irgendwo noch ne Funktion einbauen, die Selektieren kann
    print("Chosing right Brick ....")



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
