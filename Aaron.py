print ("Aaron\n\nExodus 2,16 'Aaron wird für dich zum Volk sprechen. Es ist so, als ob du durch ihn sprichst. Und er wird deine Botschaften weitergeben, so wie ein Prophet meine.'\n\n")

import os

import openpyxl
from openpyxl import load_workbook

def getTheVars():       #export variables from Excel?
    wb = load_workbook("C:/Users/Dyrian/Desktop/Aaron/Infoinput.xlsx")
    name= #read aus wb aus der entsprechenden Zeile/Spalte
    print(wb)
    input()


try:
    getTheVars()



finally:
    input("hats geklappt?")
