print ("Aaron\n\nExodus 2,16 'Aaron wird für dich zum Volk sprechen. Es ist so, als ob du durch ihn sprichst. Und er wird deine Botschaften weitergeben, so wie ein Prophet meine.'\n\n")

#Die Idee: Anhand der Daten aus Infoinput.xlsx wird ein Dokument zusammengestellt.
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
sdate = "STERBEDATUM"


def getTheVars():       #export variables from Excel?
    global path
    wb = load_workbook(path+"\\Infoinput.xlsx")          # am besten baue ich hier direkt ein, dass er auf die excel tabelle im gleichen Ordner schaut ....
    #name= #!/usr/bin/python
    ws = wb.active
    global lname
    lname = (ws['C2'].value)
    global sname
    sname = (ws['D2'].value)
    print("Name:            ",sname," ",lname)
    global sort
    sort = (ws['C7'].value)
    print("Sterbeort:       ",sort)
    global sdate
    sdate = datetime.datetime.strftime((ws['C5'].value), '%d/%b/%Y')        #convert date into string ... looks better -.-
    print("Sterbedatum:     ",sdate)


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













def choseRightBrick():              # letztlich muss ich irgendwo noch ne Funktion einbauen, die Selektieren kann
    print("Chosing right Brick ....")


def buildEverything():
    buildTheIntro()
    #buildAnsprache()
    #buildOuttro()
    #buildAmGrab()

    document.save(lname+'.docx')
    print("Build Everything")


def buildTheIntro():
    print("Building the intro now ...")
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
    input("hats geklappt?")





#Das hier brauch in unebdingt!!!! Damit kann ich quasi die Formatierung der Brick Dateien übernehmen ... das ist awesome
#
#
#def get_para_data(output_doc_name, paragraph):
#    # Write the run to the new file and then set its font, bold, alignment, color etc. data.
#    output_para = output_doc_name.add_paragraph()
#    for run in paragraph.runs:
#        output_run = output_para.add_run(run.text)
#        # Run's bold data
#        output_run.bold = run.bold
#        # Run's italic data
#        output_run.italic = run.italic
#        # Run's underline data
#        output_run.underline = run.underline
#        # Run's color data
#        output_run.font.color.rgb = run.font.color.rgb
#        # Run's font data
#        output_run.style.name = run.style.name
    # Paragraph's alignment data
#    output_para.paragraph_format.alignment = paragraph.paragraph_format.alignment

#input_doc = Document('InputDoc.docx')
#output_doc = Document()
#
# Call the function
#get_para_data(output_doc, input_doc.paragraphs[3])
#
# Save the new file
#output_doc.save('OutputDoc.docx')
#
#for para in input_doc.paragraphs:
#    get_para_data(output_doc, para)
#
#output_doc.save('OutputDoc.docx')
