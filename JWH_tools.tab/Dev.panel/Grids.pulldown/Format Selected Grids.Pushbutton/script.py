import clr
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
import math

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
selection = uidoc.Selection
view = doc.ActiveView

selected_elementsID = selection.GetElementIds()
selected_elements = [doc.GetElement(eid) for eid in selected_elementsID]


FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Grids).WhereElementIsNotElementType().ToElements()

# print(FEC)
TextNoteType = FilteredElementCollector(doc).OfClass(TextNoteType).FirstElementId()

deg = (90)*-math.pi/180

count = 0

view_direction = view.ViewDirection

#FEC = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_Grids).WhereElementIsNotElementType().ToElements()

for element in selected_elements:
    t = Transaction(doc, "Update Grids")
    t.Start()
    Text = "CH %s  +  %s" %(str(element.Name)[:2] , str(element.Name)[2:])
    tx_point = XYZ(element.Curve.Origin.X+1.5,element.Curve.Origin.Y,element.GetExtents().MaximumPoint.Z)#-1.6)
    a = TextNote.Create(doc,view.Id,tx_point,Text,TextNoteType)   
    # rotate text note
    rotationPoint = tx_point
    line = Line.CreateBound(rotationPoint, XYZ(rotationPoint.X+view_direction.X, rotationPoint.Y+view_direction.Y, rotationPoint.Z))
    ElementTransformUtils.RotateElement(doc, a.Id, line, deg) # 1.5708 radians = 90 degrees
    count += 1
    t.Commit()

print("%s Grids Updated in %s " %(count, view.Name))

print("DONE!")