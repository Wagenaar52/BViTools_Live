import clr
from Autodesk.Revit.DB import Transaction, ElementTransformUtils
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
selection = uidoc.Selection
view = doc.ActiveView


selected_elementsID = selection.GetElementIds()
selected_elements = [doc.GetElement(eid) for eid in selected_elementsID]




OriginZList = []
IncreasingList = []

for elem in selected_elements:
    if elem.Category.Name == "Dimensions":

        print(elem.Origin)
        OriginZList.append(elem.Origin.Z)
        print(elem.LeaderEndPosition)
        print('loop ran')

spacing = (max(OriginZList) - min(OriginZList))/len(OriginZList)



t = Transaction(doc)
t.Start("Set leader end")

print(spacing)


for elem in selected_elements:
    if elem.Category.Name == "Dimensions":
        mve = XYZ(elem.Location.X, elem.Location.Y, elem.Location.Z) + XYZ(0,0,spacing)
        ElementTransformUtils.MoveElement(doc, elem.Id, mve)
        print('loop ran')

t.Commit()

