print ("Aaron\n\nExodus 2,16 'Aaron wird für dich zum Volk sprechen. Es ist so, als ob du durch ihn sprichst. Und er wird deine Botschaften weitergeben, so wie ein Prophet meine.'\n\n")

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
    global path
    document = Document()
    document.add_heading('This is the title', 0)
    p = document.add_paragraph('And this is text222222222222222222222 ')
    p.add_run('some bold text').bold = True
    p.add_run('and italic text.').italic = True
    document.save(lname+'.docx')


    buildTheIntro()
    #buildNextPart ....



def buildTheIntro():
    input("Building the intro now ...")





try:
    getTheVars()

    buildEverything()



finally:
    input("hats geklappt?")
