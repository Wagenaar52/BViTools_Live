from Autodesk.Revit.DB import Transaction, Structure, FilteredElementCollector, ElementId, ElementParameterFilter, ParameterValueProvider, FilterStringRule, FilterStringEquals
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, Line, XYZ
from pyrevit import forms, output
import math
from Autodesk.Revit.DB import ElementId

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
selection = uidoc.Selection
view = doc.ActiveView


# apply filter to elements in active view of category OST_DetailComponents and family SOFiSTiK_Detail_Station

chainages = []

FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DetailComponents).WhereElementIsNotElementType().ToElements()

for element in FEC:
    if element.Name == "On Axis":
        # print(element.Name)
        # print(element.LookupParameter("Station_km").AsValueString())
        # print(element.LookupParameter("Station_m").AsValueString())

        km = int(element.LookupParameter("Station_km").AsValueString())
        m = round(float(element.LookupParameter("Station_m").AsValueString()),0)
        chainage = (km*1000 + m)/1000
        chainages.append(chainage)
        

# print('#'*50)
# print(chainages)
# print(len(chainages))

# sort list of chainages from smallest to largest

chainages.sort()

# print(chainages)

Grids = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Grids).WhereElementIsNotElementType().ToElements()

gridNames = []

for element in Grids:
    gridNames.append(element.Name)
    chainVal = int(element.Name[2:])
    print(chainVal)
    print(35050+(chainVal*10))
    t = Transaction(doc, "Update Grids")
    t.Start()
    element.Name = str(35050+(chainVal*10))
    t.Commit()
    print("klaar")
#     print(element.Curve.Origin)
    
    # if element name contains "_" delete element 
    # if 'A-0 (0)' in str(element.Name):
    #     print(element.Name)
    #     tx = Transaction(doc, "Delete Grids")
    #     tx.Start()
    #     doc.Delete(element.Id)
    #     tx.Commit()
   

    # create text note
    # TextNoteType = FilteredElementCollector(doc).OfClass(TextNoteType).ToElements()
    # textNote = TextNote.Create(doc, view.Id, element.Curve.Origin, element.Name, TextNoteType)

# print(gridNames)
# print(len(gridNames))

# gridNames.sort()

# print(gridNames)
# print(50*'#')
# for grid in gridNames:
#     chainVal = int(grid[2:])
    

