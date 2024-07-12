import clr
import math
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
selection = uidoc.Selection
view = doc.ActiveView


selected_elementsID = selection.GetElementIds()
selected_elements = [doc.GetElement(eid) for eid in selected_elementsID]

spotElevation_elements = []
count = 0
TextNoteType = FilteredElementCollector(doc).OfClass(TextNoteType).FirstElementId()  

for element in selected_elements:
    t = Transaction(doc, "Update Spot Elevations")
    t.Start()
    Text = "%s" %(str(round((element.Location.Point.Z *0.3048),3) ))
    tx_point = XYZ(element.Location.Point.X,element.Location.Point.Y,element.Location.Point.Z+2)
    a = TextNote.Create(doc,view.Id,tx_point,Text,TextNoteType)   
    
    # create detail line under spot elevation text
    Pt1 = XYZ(element.Location.Point.X,element.Location.Point.Y,element.Location.Point.Z+((0.2/0.3048)))
    Pt2 = XYZ(element.Location.Point.X+((1/0.3048)),element.Location.Point.Y,element.Location.Point.Z+((0.2/0.3048)))
    curve1 = Line.CreateBound(Pt1,Pt2)
    doc.Create.NewDetailCurve(view, curve1)
    
    count += 1
    
    t.Commit()

print("%s Spot Elevations Updated" %count)
    

