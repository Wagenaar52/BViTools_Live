from Autodesk.Revit.DB.Structure import * 
from Autodesk.Revit.DB.Structure import RebarShape
import math, clr
from Autodesk.Revit.DB import Transaction, Structure, FilteredElementCollector, RadialArray, ArrayAnchorMember
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, Line, XYZ
from Autodesk.Revit.DB import FailureSeverity, FailureProcessingResult,IFailuresPreprocessor
from pyrevit import forms
from Autodesk.Revit.DB import *
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
view = doc.ActiveView


def pol_car(radius, angle_degrees):
    x = radius * math.cos(math.radians(angle_degrees))
    y = radius * math.sin(math.radians(angle_degrees))
    return x, y

Lstyle = doc.GetElement(ElementId(1019189))

t = Transaction(doc, "Draw line")
t.Start()

radlist = [3950	,
4450	,
4950	,
5450	,
5950	,
6500	,
7100	,
7700	,
8300	,
8900	,
9500	,
10075	,
10591.5	,
11083	
]

for rad in radlist:
    radius = rad/304.8
    new_arc = Arc.Create(XYZ(radius,0,0)  , XYZ(-radius,0,0) ,  XYZ(0,radius,0))#,  XYZ(pol_car(radius,45)[0],pol_car(radius,45)[1],0)  )
    #detail line from arc
    model_line = doc.Create.NewDetailCurve(view, new_arc)
    model_line.LineStyle = Lstyle
    det_line = doc.Create.NewDetailCurve(view, Line.CreateBound(XYZ(radius,0,0),XYZ(radius,0.1,0)))


t.Commit()
############################################################################################################################################################################
############################################################################################################################################################################
############################################################################################################################################################################

# FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()

# t = Transaction(doc, "Update r custom")
# t.Start()

# for rebar in FEC:
#     if rebar.LookupParameter("Rebar r Custom").AsDouble() == 0.0:
#         if rebar.LookupParameter("r").AsDouble() != 0.0:
#             rebar.LookupParameter("Rebar r Custom").Set(rebar.LookupParameter("r").AsDouble())
#             print(rebar.LookupParameter("Schedule Mark").AsString())
#             print(rebar.LookupParameter("Rebar r Custom").AsDouble())
#             print(rebar.LookupParameter("r").AsDouble())
#         elif rebar.LookupParameter("r").AsDouble() == 0.0:
#             rebar.LookupParameter("Rebar r Custom").Set(0.0)
#             print(rebar.LookupParameter("Schedule Mark").AsString())
#             print(rebar.LookupParameter("Rebar r Custom").AsDouble())
#             print(rebar.LookupParameter("r").AsDouble())

# print("############ done ############")
# t.Commit()

############################################################################################################################################################################
############################################################################################################################################################################
############################################################################################################################################################################
############################################################################################################################################################################

# alphaList = ['a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z']
# FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()

# t = Transaction(doc, "Update r custom")
# t.Start()

# for rebar in FEC:
#     if "GR100" in rebar.LookupParameter("Schedule Mark").AsString():
#             currentLetter = str(rebar.LookupParameter("Schedule Mark").AsString()[-1])
#             print(currentLetter)
#             currentLetterIndex = alphaList.index(currentLetter)
#             print(currentLetterIndex)
#             rebar.LookupParameter("Schedule Mark").Set("GR100" + str(alphaList[currentLetterIndex-1]))
#             print(alphaList[currentLetterIndex+1])

# print("############ done ############")
# t.Commit()


# FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
# scheduleMarkList = []
# for elem in FEC:
#     if elem.LookupParameter("Schedule Mark").AsString() not in scheduleMarkList:
#         scheduleMarkList.append(elem.LookupParameter("Schedule Mark").AsString())


            # Create a polyline from the points
            # points = [rebar_p1, rebar_p2, rebar_p3, rebar_p4, rebar_p5, rebar_p6]
            # polyline = PolyLine.Create(points)
            # #draw polyline from curveList   

            # curve = PolyLine.Create(curveList)
            
            # geomPlane = Plane.CreateByThreePoints(rebar_p1, rebar_p2, rebar_p3)
            # # Create a sketch plane in current document
            # sketch = SketchPlane.Create(doc, geomPlane)
            #create model lines form curve

            # model_line = doc.Create.NewModelCurve(curve1, sketch)
            # model_line = doc.Create.NewModelCurve(curve2, sketch)
            # model_line = doc.Create.NewModelCurve(curve3, sketch)
            # model_line = doc.Create.NewModelCurve(curve4, sketch)
            # model_line = doc.Create.NewModelCurve(curve5, sketch)
            # model_line = doc.Create.NewModelCurve(curve6, sketch)
        
            #for curve in curve_array:
            #for curve in curveList:



