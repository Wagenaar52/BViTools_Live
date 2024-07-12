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

#FPath =  "C:\Users\Wagner.Human\Desktop\Wolf_RebarData_RevD_V162r5.xlsx" 
FPath = forms.pick_file(file_ext='xlsx', multi_file=False, unc_paths=False)

excel = Excel.ApplicationClass()
excel.Visible = False
workbook = excel.Workbooks.Open(FPath)
xl = workbook.Worksheets['A']

radlist = []

# get all inputs form excel for BC in a dictionary
barDict = {}
for i in range(1,100):
    i += 1
    if "BC" in str(xl.Cells(i, 1).Value2).replace(" ","") :
        bar_mark = str(xl.Cells(i,1).Value2)
        startRadius = float(xl.Cells(i,2).Value2)/304.8
        endRadius = float(xl.Cells(i,3).Value2)/304.8
        bar_size = "Y" + str(xl.Cells(i,4).Value2)[1:3]
        bar_dia = int(xl.Cells(i,4).Value2[1:3])/304.8
        spacing = int(xl.Cells(i,7).Value2)/304.8
        
        bar_parameters = {
            "bar_mark": bar_mark,
            "startRadius": startRadius,
            "endRadius": endRadius,
            "bar_size": bar_size,
            "bar_dia": bar_dia,
            "spacing": spacing
        }
        
        barDict[bar_mark] = bar_parameters
        if bar_mark == "BC100":
            radlist.append(startRadius)
        else:
            radlist.append(endRadius)

print(radlist)

LapConst = 45


def pol_car(radius, angle_degrees):
    x = radius * math.cos(math.radians(angle_degrees))
    y = radius * math.sin(math.radians(angle_degrees))
    return x, y

Lstyle = doc.GetElement(ElementId(1019189))

t = Transaction(doc, "Update r")
t.Start()


# for radius in radlist:
#     new_arc = Arc.Create(XYZ(radius,0,0)  , XYZ(0,radius,0) ,  XYZ(pol_car(radius,45)[0],pol_car(radius,45)[1],0))
#     #detail line from arc
#     model_line = doc.Create.NewDetailCurve(view, new_arc)
#     model_line.LineStyle = Lstyle
#     # create text for radius

#create bottom concentric bars between barmark star and end radius

bcList1 = []    ; bcList2 = []  ;bcList3 = []   ;bcList4 = []   ;bcList5 = []   ;bcList6 = []   ;bcList7 = []   ;bcList8 = []   ;bcList9 = []

bcLists =[bcList1, bcList2, bcList3, bcList4, bcList5, bcList6, bcList7, bcList8, bcList9]

FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
for i in range(1,10):
    for elem in FEC:
        if ("BC"+str(i)) in elem.LookupParameter("Schedule Mark").AsString():
            bcLists[i-1].append(elem)
    


for bar_mark, bar_parameters in barDict.items():
    startRadius = bar_parameters["startRadius"]
    endRadius = bar_parameters["endRadius"]
    bar_size = bar_parameters["bar_size"]
    bar_dia = bar_parameters["bar_dia"]
    spacing = bar_parameters["spacing"]
    
    # Create a text note
    # Define the text for the note
    text = str(bar_size)+"-"+str(bar_mark)+"-"+str(spacing*304.8)[:-2]

    # Get the default text note type
    text_note_type_id = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType)

    # Create the text note
    text_note = TextNote.Create(doc, view.Id, XYZ(endRadius,0,0), text, text_note_type_id)
    text_note.TextNoteType = doc.GetElement(ElementId(1018389))
    text_note.Coord = XYZ(text_note.Coord.X, text_note.Coord.Y-1, text_note.Coord.Z)
    #rotate text note
    # p_1 = text_note.Coord
    # p_2 = XYZ(text_note.Coord.X, text_note.Coord.Y, text_note.Coord.Z+1)
    # text_note.Location.Rotate(Line.CreateBound(p_1,p_2), math.radians(90))


t.Commit()





def pol_car(radius, angle_degrees):
    x = radius * math.cos(math.radians(angle_degrees))
    y = radius * math.sin(math.radians(angle_degrees))
    return x, y

Lstyle = doc.GetElement(ElementId(1019189))

t = Transaction(doc, "Update r")
t.Start()



for rad in radlist:
    new_arc = Arc.Create(XYZ(rad,0,0)  ,  XYZ(0,rad,0) ,   XYZ(pol_car(rad,45)[0],pol_car(rad,45)[1],0)  )
    #detail line from arc
    model_line = doc.Create.NewDetailCurve(view, new_arc)
    model_line.LineStyle = Lstyle
    model_line = doc.Create.NewDetailCurve(view, Line.CreateBound(XYZ(rad,0,0), XYZ(rad,0.1,0)))

t.Commit()



























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



