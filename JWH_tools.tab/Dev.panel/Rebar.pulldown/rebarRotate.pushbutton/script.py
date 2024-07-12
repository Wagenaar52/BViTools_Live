from Autodesk.Revit.DB import Transaction, Structure, FilteredElementCollector
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, Line, XYZ
from pyrevit import forms
import math

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
selection = uidoc.Selection
view = doc.ActiveView

selected_elementsID = selection.GetElementIds()
selected_elements = [doc.GetElement(eid) for eid in selected_elementsID]


userDeg = forms.ask_for_string(
    default='90',
    prompt='Enter Rotation Angle',
    title='Rotate Rebar',
    width=200,
    height=200
    )
deg = float(userDeg)*-math.pi/180

t = Transaction(doc, 'Rotate Rebar')

t.Start()

for element in selected_elements:
    if element.Category.Name == 'Structural Rebar':
        element.Location.Rotate(Line.CreateBound(XYZ(0, 0, 0), XYZ(0, 0, 1)), deg)
    else:
        pass

t.Commit()










