print ("Aaron\n\nExodus 2,16 'Aaron wird für dich zum Volk sprechen. Es ist so, als ob du durch ihn sprichst. Und er wird deine Botschaften weitergeben, so wie ein Prophet meine.'\n\n")

import os
import datetime
import openpyxl
from openpyxl import load_workbook

lname = "MUSTERMANN"
sname = "MAX"
sort = "STERBEORT"
sdate = "STERBEDATUM"

def getTheVars():       #export variables from Excel?
    wb = load_workbook("C:/Users/Dyrian/Desktop/Aaron/Infoinput.xlsx")
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


def buildTheIntro():
    print(sort)
    input("Building the intro now ...")





try:
    getTheVars()

    buildTheIntro()



finally:
    input("hats geklappt?")
