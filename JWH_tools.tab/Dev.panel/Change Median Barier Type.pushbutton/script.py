from Autodesk.Revit.DB import Transaction, Structure, FilteredElementCollector, ElementId, ElementParameterFilter, ParameterValueProvider, FilterStringRule, FilterStringEquals
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, Line, XYZ
from pyrevit import forms, output
import math
from Autodesk.Revit.DB import ElementId

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
selection = uidoc.Selection
view = doc.ActiveView


# get all Median Barrier Retaining Walls in project

mbrw = []

FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()

for element in FEC:
    if element.Name == "LHS":
        mbrw.append(element)
        # get location curve of wall panel
        print(element.LookupParameter("Type").AsValueString())
        

# print('#'*50)
# print(mbrw)

for element in mbrw:
    # get location curve of wall panel
    print(element.LookupParameter("Type").AsValueString())
    print(element.GetOrderedParameters().AsValueString())
 
    paramList = []

    for param in element.GetOrderedParameters():
        paramList.append(param)
    print(paramList)


# t = Transaction(doc, "Change visibility parameter")

