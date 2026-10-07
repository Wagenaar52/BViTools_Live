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
view = doc.GetElement(ElementId(446567))

def pol_car(radius, angle_degrees):
    x = radius * math.cos(math.radians(angle_degrees))
    y = radius * math.sin(math.radians(angle_degrees))
    return x, y

#FPath =  "C:\Users\Wagner.Human\Desktop\Wolf_RebarData_RevD_V162r5.xlsx" 
FPath = forms.pick_file(file_ext='xlsx', multi_file=False, unc_paths=False)

excel = Excel.ApplicationClass()
excel.Visible = False
workbook = excel.Workbooks.Open(FPath)
xl = workbook.Worksheets['A']

radlist = []

# get all inputs form excel for BC in a dictionary
barDict = {}
for i in range(1,200):
    i += 1
    if "TC" in str(xl.Cells(i, 1).Value2).replace(" ","") :
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

        radlist.append(startRadius)

largestEndRadius = 0
for i in range(1,200):
    if "TC" in str(xl.Cells(i, 1).Value2).replace(" ","") :
        bar_mark = str(xl.Cells(i,1).Value2)
        startRadius = float(xl.Cells(i,2).Value2)/304.8
        endRadius = float(xl.Cells(i,3).Value2)/304.8
        if endRadius > largestEndRadius:
            largestEndRadius = endRadius

radlist.append(largestEndRadius)

FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
for elem in FEC:
    if elem.Name == "1PA_WTF_SteelTower":
        Foundation = elem
        break
rFoundation =  Foundation.LookupParameter("rBase").AsDouble()

barMarkList = []
FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
for elem in FEC:
    if elem.LookupParameter("Mark").AsString() == "TOP CONCENTRIC":
        barMarkList.append(elem.LookupParameter("Schedule Mark").AsString())
# Remove duplicate 
barMarkList = list(dict.fromkeys(barMarkList))


t = Transaction(doc, "Top Concentric Annotation ")
t.Start()

#hiden lines at border of view

def polar_to_cartesian(radius, angle):
    x = radius * math.cos(angle)
    y = radius * math.sin(angle)
    return XYZ(x, y, 0)

Lstyle = doc.GetElement(ElementId(1019189))

horModel_line = doc.Create.NewDetailCurve(view, Line.CreateBound(XYZ(0,0,0), XYZ(rFoundation,0,0)))
horModel_line.LineStyle = Lstyle
verModel_line = doc.Create.NewDetailCurve(view, Line.CreateBound(XYZ(0,0,0), XYZ(0,rFoundation,0)))
verModel_line.LineStyle = Lstyle
model_line = doc.Create.NewDetailCurve(view, Line.CreateBound(XYZ(0,0,0), polar_to_cartesian(rFoundation, math.radians(45))))
model_line.LineStyle = Lstyle

#Dimensions 
oldCircle = horModel_line
radlist.sort()

for rad in radlist:
    new_circle = Arc.Create(Plane.CreateByNormalAndOrigin(XYZ(0,0,1), XYZ(0,0,0)), rad, 0, 2*math.pi)
    #detail line from circle
    model_line = doc.Create.NewDetailCurve(view, new_circle)
    model_line.LineStyle = Lstyle
    
    refArray = ReferenceArray()
    refArray.Append(oldCircle.GeometryCurve.Reference)
    refArray.Append(model_line.GeometryCurve.Reference)
    dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(-1,0,0), XYZ(-1,rFoundation,0)), refArray, doc.GetElement(ElementId(1018370)))
    text = None
    for bm in barDict.keys():
        if str(barDict[bm]["endRadius"]) == str(rad):
            bar_size = barDict[bm]["bar_size"]
            bar_mark = bm

            tempList = []
            for mark in barMarkList:
                if str(bar_mark)[:3] in mark.replace(" ","") :
                    tempList.append(int(str(mark)[3:]))
            if len(tempList) > 1:
                bar_mark = str(bar_mark)[:2] +"("+str(bar_mark)[2:4] + str(min(tempList)) + "-" +str(bar_mark)[2:4] + str(max(tempList)) + ")"

            spacing = barDict[bm]["spacing"]
            text = str(bar_size)+"-"+str(bar_mark)+"-"+str(spacing*304.8)[:-2]
    if text is not None:
        dim.Below = text
    else:
        dim.Below = "PLINTH"
   
    oldCircle = model_line


#rBase dimension
new_circle = Arc.Create(Plane.CreateByNormalAndOrigin(XYZ(0,0,1), XYZ(0,0,0)), rFoundation, 0, 2*math.pi)
model_line = doc.Create.NewDetailCurve(view, new_circle)
model_line.LineStyle = Lstyle
refArray = ReferenceArray()
refArray.Append(horModel_line.GeometryCurve.Reference)
refArray.Append(model_line.GeometryCurve.Reference)
dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(-2,0,0), XYZ(-2,rFoundation,0)), refArray, doc.GetElement(ElementId(1018370)))


t.Commit()

