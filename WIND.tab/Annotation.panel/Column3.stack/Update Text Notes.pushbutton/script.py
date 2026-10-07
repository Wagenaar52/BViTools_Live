import clr
import math
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import forms

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document


FPath = forms.pick_file(file_ext='xlsx', multi_file=False, unc_paths=False)


#get hould of excel file using ironpython
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
excel = Excel.ApplicationClass()
excel.Visible = False
workbook = excel.Workbooks.Open(FPath)
xl = workbook.Worksheets['Annotation']

t = Transaction(doc)
t.Start("Apply parameter values")


for row in range(3, 1000):
    if xl.Cells(row, 4).Value2 is None:
        break
    else:
        textId = xl.Cells(row, 4).Value2
        textValue = xl.Cells(row, 5).Value2
        textNote = doc.GetElement(ElementId(int(textId)))
        textNote.Text = textValue

t.Commit()

print("condeRan ")
# Close Excel application object
workbook.Close(False)
excel.Quit()