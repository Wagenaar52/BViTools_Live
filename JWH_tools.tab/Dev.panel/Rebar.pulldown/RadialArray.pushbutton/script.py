from Autodesk.Revit.DB import Transaction, FilteredElementCollector, BuiltInCategory, Line, XYZ, RadialArray, ElementId, ArrayAnchorMember
from System.Collections.Generic import List
# from Autodesk.Revit.DB import
from pyrevit import forms
import math

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
selection = uidoc.Selection
view = doc.ActiveView

selected_elementsID = selection.GetElementIds()
selected_elements = [doc.GetElement(eid) for eid in selected_elementsID]

RotAngle = int(forms.ask_for_string(prompt="Enter Rotation Angle",default="360"))*math.pi/180

t = Transaction(doc, 'Rotate Rebar')

t.Start()

for element in selected_elements:
    if element.Category.Name == 'Structural Rebar':
        RadialArray.Create(doc, view, element.Id, 4, Line.CreateBound(XYZ(0, 0, 0), XYZ(0, 0, 1)), RotAngle, ArrayAnchorMember.Last)
    else:
        pass

t.Commit()




















