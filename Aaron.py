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
#from docx.text.parargaph import Paragraph      klappt irgendwie nicht ...



path = "."

# Vars to check ...
lname = "MUSTERMANN"
sname = "MAX"
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


def buildEverything():
    buildTheIntro()
    #buildNextPart ....





def buildTheIntro():
    input("Building the intro now ...")
    global path
    document = Document()

    #Headline
    header = document.sections[0].header
    header.paragraphs[0].text = "Trauerfeier "+sname+" "+lname

    #Header
    document.add_heading("Trauerfeier von "+sname+" "+lname, 0)

    #Votum
    document.add_heading("Votum", 1)
    p = document.add_paragraph("Im Namen des Vaters ....")
    input("zwischentest..")
    brick = Document(path+"\\Bricks\\Votum\\Votum.docx")

    paragraphs = []

    for para in brick.paragraphs:
        p = para.text
        paragraphs.append(p)
#    output = Document()
    for item in paragraphs:
        document.add_paragraph(item)
#    document.save('OutputDoc.docx')



    p = document.add_paragraph("And this is text222222222222222222222 ")

    p.add_run("some bold text").bold = True
    p.add_run("and italic text.").italic = True
    p.add_run("weitererText").bold = True


    document.add_heading('This is the title2', 0)
    p = document.add_paragraph('This is the second paragraph')
    p.add_run('some bold text').bold = True
    
    document.save(lname+'.docx')
    input("Intro finished")



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
