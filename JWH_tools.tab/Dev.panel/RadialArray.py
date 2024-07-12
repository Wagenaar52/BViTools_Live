from Autodesk.Revit.DB import Transaction, FilteredElementCollector, BuiltInCategory, Line, XYZ, RadialArray
# from Autodesk.Revit.DB import
from pyrevit import forms
import math

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
selection = uidoc.Selection
view = doc.ActiveView

selected_elementsID = selection.GetElementIds()
selected_elements = [doc.GetElement(eid) for eid in selected_elementsID]


t = Transaction(doc, 'Rotate Rebar')

t.Start()

for element in selected_elements:
    if element.Category.Name == 'Structural Rebar':
        print("1")
        RadialArray(doc, view, element.Id, 20, Line.CreateBound(XYZ(0, 0, 0), XYZ(0, 0, 1)), 90, XYZ(0, 0, 1))
        print("1")
        #Create(Document, View, ElementId, int, Line, double, ArrayAnchorMember)
    else:
        pass

t.Commit()










