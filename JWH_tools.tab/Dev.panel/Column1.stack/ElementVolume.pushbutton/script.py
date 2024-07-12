import clr

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
selection = uidoc.Selection
view = doc.ActiveView


selected_elementsID = selection.GetElementIds()
selected_elements = [doc.GetElement(eid) for eid in selected_elementsID]

total_Volume = 0
panelNumber = 0
count = 0   

for element in selected_elements:
    total_Volume += element.LookupParameter("Volume").AsDouble()*0.3048**3
    panelNumber += 1
    print(total_Volume)
    print(panelNumber)



print(total_Volume)