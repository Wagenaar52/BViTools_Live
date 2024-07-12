import clr
import math
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import forms
from Autodesk.Revit.DB import ViewSheet

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

FPath = 'C:\\Users\\Wagner.Human\\Desktop\\book3.xlsx'
#FPath =  "C:\\Users\\Wagner.Human\\DC\\ACCDocs\\BVi - Western Cape\\33808 - C1159- R300 EXTENSION NORTH FROM THE N1 TO N7\\Project Files\\00_COMMON\\00_MIDP\\33808-BVI-00-9000-DDP-MID-00000.xlsx" 

#get hould of excel file using ironpython
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
excel = Excel.ApplicationClass()
excel.Visible = False
workbook = excel.Workbooks.Open(FPath)
xl = workbook.ActiveSheet

# get all titleblocks in FEC
FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType().ToElements()
TitleBlock = []
for titleblock in FEC:
    if titleblock.LookupParameter("Family Name").AsString() == "BVi_WCG_TITLEBLOCK_A0":
        TitleBlock.append(titleblock)  # Append the title block to the list
        break

t = Transaction(doc, 'Create Sheets')

t.Start()

for row in range(2, 8):
    d01 = str(xl.Cells(row,12).Value2)
    d02 = str(xl.Cells(row,6).Value2)
    d03 = str(xl.Cells(row,15).Value2)
    d04 = str(xl.Cells(row,16).Value2)

    dwg_str = d01 + d02 + '-01-'  + d04
#SheetList = []
#ON_form = forms.SelectFromList.show(SheetList, button_name='Select Parameters to turn on', multiselect=True)

    sht = ViewSheet.Create(doc, TitleBlock[0].Id)
    sht.Name = dwg_str
    sht.SheetNumber = str(xl.Cells(row,14).Value2)

    paramCol = 21
    for param in sht.Parameters:
        print(str(param.Definition.Name) + " : " + str(param.AsString()))
        xl.Cells(row, 20).Value2 = sht.Name
        xl.Cells(1, paramCol).Value2 = param.Definition.Name 
        xl.Cells(row, paramCol).Value2 = param.AsString()
        paramCol += 1


    print("*"*50)

t.Commit()
workbook.Save()


excel.Quit()



