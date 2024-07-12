from Autodesk.Revit.DB import Transaction, Structure, FilteredElementCollector, ElementId, ElementParameterFilter, ParameterValueProvider, FilterStringRule, FilterStringEquals
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, Line, XYZ
from pyrevit import forms, output
import math
from Autodesk.Revit.DB import ElementId
from Autodesk.Revit.DB import IndependentTag
from pyrevit import revit, DB
from pyrevit import forms, output

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
selection = uidoc.Selection
view = doc.ActiveView


# get all wall panels in view
selected_elementsID = selection.GetElementIds()
selected_elements = [doc.GetElement(eid) for eid in selected_elementsID]
walls = []

# tag all elements in selected elements with a tag on the location of the element   


FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()

for element in selected_elements:
    # move element tag
    t = Transaction(doc, "move element tag")
    t.Start()    
    element.TagHeadPosition = XYZ(element.TagHeadPosition.X,element.TagHeadPosition.Y,300)
    t.Commit()

