from Autodesk.Revit.DB.Structure import * 
from Autodesk.Revit.DB.Structure import RebarShape
import math, clr
from System.Collections.Generic import List
from Autodesk.Revit.DB import Transaction, Structure, FilteredElementCollector, RadialArray, ArrayAnchorMember
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, Line, XYZ
from Autodesk.Revit.DB import FailureSeverity, FailureProcessingResult,IFailuresPreprocessor
from pyrevit import forms
from Autodesk.Revit.DB import *
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
view = doc.GetElement(ElementId(453671))


def polar_to_cartesian(radius, angle):
    x = radius * math.cos(angle)
    y = radius * math.sin(angle)
    return XYZ(x, y, 0)


def get_rebar_tag_reference(rebar_element):
    if rebar_element is None or not rebar_element.IsValidObject:
        raise Exception("Invalid rebar element passed to get_rebar_tag_reference")

    subelements = rebar_element.GetSubelements()
    if subelements and len(subelements) > 0:
        return subelements[0].GetReference()

    return Reference(rebar_element)


FPath = forms.pick_file(file_ext='xlsx', multi_file=False, unc_paths=False)

excel = Excel.ApplicationClass()
excel.Visible = False
workbook = excel.Workbooks.Open(FPath)
xl = workbook.Worksheets['A']

radlist = []

# get all inputs form excel for BC in a dictionary
Br_barDict = {}
Tr_barDict = {}
St_barDict = {}
for i in range(1,200):
    i += 1
    if "BR" in str(xl.Cells(i, 1).Value2).replace(" ","") :
        bar_mark = str(xl.Cells(i,1).Value2)
        startRadius = float(xl.Cells(i,2).Value2)/304.8
        endRadius = float(xl.Cells(i,3).Value2)/304.8
        bar_size = "Y" + str(xl.Cells(i,4).Value2)[1:3]
        bar_dia = int(xl.Cells(i,4).Value2[1:3])/304.8
        count = int(xl.Cells(i,8).Value2)
        spacing = int(xl.Cells(i,7).Value2)/304.8
        level = int(xl.Cells(i,10).Value2)

        bar_parameters = {
            "bar_mark": bar_mark,
            "startRadius": startRadius,
            "endRadius": endRadius,
            "bar_size": bar_size,
            "bar_dia": bar_dia,
            "spacing": spacing,
            "level": level,
            "count": count
        }
        Br_barDict[bar_mark] = bar_parameters

    elif "TR" in str(xl.Cells(i, 1).Value2).replace(" ","") :
        bar_mark = str(xl.Cells(i,1).Value2)
        startRadius = float(xl.Cells(i,2).Value2)/304.8
        endRadius = float(xl.Cells(i,3).Value2)/304.8
        bar_size = "Y" + str(xl.Cells(i,4).Value2)[1:3]
        bar_dia = int(xl.Cells(i,4).Value2[1:3])/304.8
        count = int(xl.Cells(i,8).Value2)
        spacing = int(xl.Cells(i,7).Value2)/304.8
        level = int(xl.Cells(i,10).Value2)

        bar_parameters = {
            "bar_mark": bar_mark,
            "startRadius": startRadius,
            "endRadius": endRadius,
            "bar_size": bar_size,
            "bar_dia": bar_dia,
            "spacing": spacing,
            "level": level,
            "count": count
        }
        Tr_barDict[bar_mark] = bar_parameters

    elif "ST" in str(xl.Cells(i, 1).Value2).replace(" ","") :
        bar_mark = str(xl.Cells(i,1).Value2)
        Radius = float(xl.Cells(i,2).Value2)/304.8
        count = float(xl.Cells(i,3).Value2)/304.8
        bar_size = "Y" + str(xl.Cells(i,4).Value2)[1:3]
        bar_dia = int(xl.Cells(i,4).Value2[1:3])/304.8

        bar_parameters = {
            "bar_mark": bar_mark,
            "Radius": Radius,
            "count": count,
            "bar_size": bar_size,
            "bar_dia": bar_dia,
        }
        St_barDict[bar_mark] = bar_parameters

#botom radial
barlistB1 = []
for bar in Br_barDict.keys():
    if Br_barDict[bar]["level"] == 1:
        barlistB1.append(Br_barDict[bar]["bar_mark"])

barlistB2 = []
for bar in Br_barDict.keys():
    if Br_barDict[bar]["level"] == 2:
        barlistB2.append(Br_barDict[bar]["bar_mark"])

#top radial
barlistT1 = []
for bar in Tr_barDict.keys():
    if Tr_barDict[bar]["level"] == 1:
        barlistT1.append(Tr_barDict[bar]["bar_mark"])

barlistT2 = []
for bar in Tr_barDict.keys():
    if Tr_barDict[bar]["level"] == 2:
        barlistT2.append(Tr_barDict[bar]["bar_mark"])

#stools
barlistST = []
for bar in St_barDict.keys():
    barlistST.append(St_barDict[bar]["bar_mark"])

FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
for elem in FEC:
    if elem.Name == "1PA_WTF_SteelTower":
        Foundation = elem
        break
rFoundation =  Foundation.LookupParameter("rBase").AsDouble()

bottomRadialList = []
topRadialList = []
bottomRadials = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()

for bar in bottomRadials:
    if "BR" in bar.LookupParameter("Schedule Mark").AsString():
        bottomRadialList.append(bar)
    elif "TR" in bar.LookupParameter("Schedule Mark").AsString():
        topRadialList.append(bar)


#region##### Bottom 1 Radial Annotation ################################################################################################
t = Transaction(doc, "Radial Annotation DWG1")
t.Start()

#hiden lines at border of view
Lstyle = doc.GetElement(ElementId(1019189))
horModel_line = doc.Create.NewDetailCurve(view, Line.CreateBound(XYZ(0,0,0), XYZ(rFoundation,0,0)))
horModel_line.LineStyle = Lstyle
verModel_line = doc.Create.NewDetailCurve(view, Line.CreateBound(XYZ(0,0,0), XYZ(0,rFoundation,0)))
verModel_line.LineStyle = Lstyle
model_line = doc.Create.NewDetailCurve(view, Line.CreateBound(XYZ(0,0,0), polar_to_cartesian(rFoundation, math.radians(22.5))))
model_line.LineStyle = Lstyle

new_circle = Arc.Create(Plane.CreateByNormalAndOrigin(XYZ(0,0,1), XYZ(0,0,0)), rFoundation, 0, 2*math.pi)
line = doc.Create.NewDetailCurve(view, new_circle)
line.LineStyle = Lstyle

verOffset = -1
#Dimensions 
for bar in barlistB1:
    endRadius = Br_barDict[bar]["endRadius"]
    startRadius = Br_barDict[bar]["startRadius"]
    count = Br_barDict[bar]["count"]
    bar_size = Br_barDict[bar]["bar_size"]
    bar_mark = Br_barDict[bar]["bar_mark"]
    new_circle = Arc.Create(Plane.CreateByNormalAndOrigin(XYZ(0,0,1), XYZ(0,0,0)), startRadius, 0, 2*math.pi)
    line1 = doc.Create.NewDetailCurve(view, new_circle)
    line1.LineStyle = Lstyle
    new_circle = Arc.Create(Plane.CreateByNormalAndOrigin(XYZ(0,0,1), XYZ(0,0,0)), endRadius, 0, 2*math.pi)
    if endRadius == rFoundation:
        line2 = line
    else:
        line2 = doc.Create.NewDetailCurve(view, new_circle)
        line2.LineStyle = Lstyle
    refArray = ReferenceArray()
    refArray.Append(line1.GeometryCurve.Reference)
    refArray.Append(line2.GeometryCurve.Reference)
    dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(0,verOffset,0), XYZ(rFoundation,verOffset,0)), refArray, doc.GetElement(ElementId(1018370)))
    text = str(count)+"x"+str(bar_size)+"-"+str(bar_mark)+"-"+str(round(360.0/count,3)) + u"\u00b0"
    dim.Below = text

    refArray = ReferenceArray()
    refArray.Append(verModel_line.GeometryCurve.Reference)
    refArray.Append(line1.GeometryCurve.Reference)
    dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(0,verOffset,0), XYZ(rFoundation,verOffset,0)), refArray, doc.GetElement(ElementId(1018370)))

    refArray = ReferenceArray()
    refArray.Append(line2.GeometryCurve.Reference)
    refArray.Append(line.GeometryCurve.Reference)
    dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(0,0,0), XYZ(0,rFoundation,0)), refArray, doc.GetElement(ElementId(1018370)))
    
    verOffset -= 1

#rBase dimension
refArray = ReferenceArray()
refArray.Append(horModel_line.GeometryCurve.Reference)
refArray.Append(line.GeometryCurve.Reference)
dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(-2,0,0), XYZ(-2,rFoundation,0)), refArray, doc.GetElement(ElementId(1018370)))

for bar in bottomRadialList:
    if bar.LookupParameter("Schedule Mark").AsString() not in barlistB1 :
        view.HideElements(List[ElementId]([bar.Id]))

#endregion##########################################################################################################################


#region##### Bottom 2 Radial Annotation ################################################################################################
view = doc.GetElement(ElementId(5915879))

#hiden lines at border of view
Lstyle = doc.GetElement(ElementId(1019189))
model_line = doc.Create.NewDetailCurve(view, Line.CreateBound(XYZ(0,0,0), polar_to_cartesian(rFoundation, math.radians(22.5))))
model_line.LineStyle = Lstyle
model_line = doc.Create.NewDetailCurve(view, Line.CreateBound(XYZ(0,0,0), polar_to_cartesian(rFoundation, math.radians(45))))
model_line.LineStyle = Lstyle
new_circle = Arc.Create(Plane.CreateByNormalAndOrigin(XYZ(0,0,1), XYZ(0,0,0)), rFoundation, 0, 2*math.pi)
line = doc.Create.NewDetailCurve(view, new_circle)
line.LineStyle = Lstyle

#Dimensions 
for bar in barlistB2:
    endRadius = Br_barDict[bar]["endRadius"]
    startRadius = Br_barDict[bar]["startRadius"]
    count = Br_barDict[bar]["count"]
    bar_size = Br_barDict[bar]["bar_size"]
    bar_mark = Br_barDict[bar]["bar_mark"]
    new_circle = Arc.Create(Plane.CreateByNormalAndOrigin(XYZ(0,0,1), XYZ(0,0,0)), startRadius, 0, 2*math.pi)
    line1 = doc.Create.NewDetailCurve(view, new_circle)
    line1.LineStyle = Lstyle
    new_circle = Arc.Create(Plane.CreateByNormalAndOrigin(XYZ(0,0,1), XYZ(0,0,0)), endRadius, 0, 2*math.pi)
    if endRadius == rFoundation:
        line2 = line
    else:
        line2 = doc.Create.NewDetailCurve(view, new_circle)
        line2.LineStyle = Lstyle
    
    text = str(count)+"x"+str(bar_size)+"-"+str(bar_mark)+"-"+str(round(360.0/count,3)) + u"\u00b0"
    #create textnote placed at the endradius of the bar(line2) with the text variable above as text and oriented at 33.75 degrees
    textNote = TextNote.Create(doc, view.Id, XYZ(endRadius-2.5,0,0), text, ElementId(1018389))
    ElementTransformUtils.RotateElement(doc, textNote.Id, Line.CreateBound(XYZ(0,0,0), XYZ(0,0,1)), math.radians(33.75))

for bar in bottomRadialList:
    if bar.LookupParameter("Schedule Mark").AsString() not in barlistB2 :
        view.HideElements(List[ElementId]([bar.Id]))

#endregion##########################################################################################################################


#region##### Top 1 Radial Annotation ################################################################################################
view = doc.GetElement(ElementId(2511791))


#hiden lines at border of view
Lstyle = doc.GetElement(ElementId(1019189))
horModel_line = doc.Create.NewDetailCurve(view, Line.CreateBound(XYZ(0,0,0), XYZ(rFoundation,0,0)))
horModel_line.LineStyle = Lstyle
verModel_line = doc.Create.NewDetailCurve(view, Line.CreateBound(XYZ(0,0,0), XYZ(0,rFoundation,0)))
verModel_line.LineStyle = Lstyle
model_line = doc.Create.NewDetailCurve(view, Line.CreateBound(XYZ(0,0,0), polar_to_cartesian(rFoundation, math.radians(67.5))))
model_line.LineStyle = Lstyle

new_circle = Arc.Create(Plane.CreateByNormalAndOrigin(XYZ(0,0,1), XYZ(0,0,0)), rFoundation, 0, 2*math.pi)
line = doc.Create.NewDetailCurve(view, new_circle)
line.LineStyle = Lstyle

verOffset = -5
#Dimensions 
for bar in barlistT1:
    endRadius = Tr_barDict[bar]["endRadius"]
    startRadius = Tr_barDict[bar]["startRadius"]
    count = Tr_barDict[bar]["count"]
    bar_size = Tr_barDict[bar]["bar_size"]
    bar_mark = Tr_barDict[bar]["bar_mark"]
    new_circle = Arc.Create(Plane.CreateByNormalAndOrigin(XYZ(0,0,1), XYZ(0,0,0)), startRadius, 0, 2*math.pi)
    line1 = doc.Create.NewDetailCurve(view, new_circle)
    line1.LineStyle = Lstyle
    new_circle = Arc.Create(Plane.CreateByNormalAndOrigin(XYZ(0,0,1), XYZ(0,0,0)), endRadius, 0, 2*math.pi)
    if endRadius == rFoundation:
        line2 = line
    else:
        line2 = doc.Create.NewDetailCurve(view, new_circle)
        line2.LineStyle = Lstyle

    refArray = ReferenceArray()
    refArray.Append(line1.GeometryCurve.Reference)
    refArray.Append(line2.GeometryCurve.Reference)
    dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(verOffset,0,0), XYZ(verOffset,rFoundation,0)), refArray, doc.GetElement(ElementId(1018370)))
    text = str(count)+"x"+str(bar_size)+"-"+str(bar_mark)+"-"+str(round(360.0/count,3)) + u"\u00b0"
    dim.Below = text

    refArray = ReferenceArray()
    refArray.Append(horModel_line.GeometryCurve.Reference)
    refArray.Append(line1.GeometryCurve.Reference)
    dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(verOffset,0,0), XYZ(verOffset,rFoundation,0)), refArray, doc.GetElement(ElementId(1018370)))
    if endRadius < rFoundation:
        refArray = ReferenceArray()
        refArray.Append(line2.GeometryCurve.Reference)
        refArray.Append(line.GeometryCurve.Reference)
        dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(verOffset,0,0), XYZ(verOffset,rFoundation,0)), refArray, doc.GetElement(ElementId(1018370)))
        
    verOffset -= 0.7
# STOOLS ##############################################################################################################################################################

for bar in barlistST:
    Radius = St_barDict[bar]["Radius"]
    bar_dia = St_barDict[bar]["bar_dia"]
    count = St_barDict[bar]["count"]
    bar_size = St_barDict[bar]["bar_size"]
    bar_mark = St_barDict[bar]["bar_mark"]

    new_circle = Arc.Create(Plane.CreateByNormalAndOrigin(XYZ(0,0,1), XYZ(0,0,0)), Radius, 0, 2*math.pi)
    stLine = doc.Create.NewDetailCurve(view, new_circle)
    stLine.LineStyle = Lstyle

    tempList = []
    FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
    for stool in FEC:
        if stool.LookupParameter("Schedule Mark").AsString() == bar_mark:
            tempList.append(stool)


    positive_x_rebars = [
        x for x in tempList
        if x.get_BoundingBox(view) and x.get_BoundingBox(view).Max.X > 0
    ]
    if not positive_x_rebars:
        raise Exception("No rebar with positive X found for Schedule Mark: {}".format(bar_mark))

    rebar = max(positive_x_rebars, key=lambda x: x.get_BoundingBox(view).Max.Y)

    rebar.get_BoundingBox(view).Max.X

    tag = IndependentTag.Create(
            doc,
            ElementId(3488323),          # TagTypeId
            view.Id,                     # ViewId
            get_rebar_tag_reference(rebar),
            True,                        # addLeader
            TagOrientation.Horizontal,
            XYZ(-2,Radius,0)             # tag head location (XYZ)
    )

    tag.TagHeadPosition = XYZ(-2,Radius,0)
    # tag.SetLeaderEnd(XYZ(rebar.get_BoundingBox(view).Max.X,Radius,0))
    tag.LeaderEndCondition = LeaderEndCondition.Free
    if rebar.get_BoundingBox(view).Min.X > 0:
        Xvalue = rebar.get_BoundingBox(view).Min.X + bar_dia*1.1
    else:
        Xvalue = 0

    tag.SetLeaderEnd(tag.GetTaggedReferences()[0], XYZ(Xvalue,Radius,0))


##############################################################################################################################################################

#rBase dimension
refArray = ReferenceArray()
refArray.Append(horModel_line.GeometryCurve.Reference)
refArray.Append(line.GeometryCurve.Reference)
dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(verOffset,0,0), XYZ(verOffset,rFoundation,0)), refArray, doc.GetElement(ElementId(1018370)))

for bar in topRadialList:
    if bar.LookupParameter("Schedule Mark").AsString() not in barlistT1 :
        view.HideElements(List[ElementId]([bar.Id]))

#endregion##########################################################################################################################


#region##### Top 2 Radial Annotation ################################################################################################
view = doc.GetElement(ElementId(2511780))

#hiden lines at border of view
Lstyle = doc.GetElement(ElementId(1019189))
model_line = doc.Create.NewDetailCurve(view, Line.CreateBound(XYZ(0,0,0), polar_to_cartesian(rFoundation, math.radians(67.5))))
model_line.LineStyle = Lstyle
model_line = doc.Create.NewDetailCurve(view, Line.CreateBound(XYZ(0,0,0), polar_to_cartesian(rFoundation, math.radians(45))))
model_line.LineStyle = Lstyle
new_circle = Arc.Create(Plane.CreateByNormalAndOrigin(XYZ(0,0,1), XYZ(0,0,0)), rFoundation, 0, 2*math.pi)
line = doc.Create.NewDetailCurve(view, new_circle)
line.LineStyle = Lstyle

#Dimensions 
for bar in barlistT2:
    endRadius = Tr_barDict[bar]["endRadius"]
    startRadius = Tr_barDict[bar]["startRadius"]
    count = Tr_barDict[bar]["count"]
    bar_size = Tr_barDict[bar]["bar_size"]
    bar_mark = Tr_barDict[bar]["bar_mark"]
    new_circle = Arc.Create(Plane.CreateByNormalAndOrigin(XYZ(0,0,1), XYZ(0,0,0)), startRadius, 0, 2*math.pi)
    line1 = doc.Create.NewDetailCurve(view, new_circle)
    line1.LineStyle = Lstyle
    new_circle = Arc.Create(Plane.CreateByNormalAndOrigin(XYZ(0,0,1), XYZ(0,0,0)), endRadius, 0, 2*math.pi)
    if endRadius == rFoundation:
        line2 = line
    else:
        line2 = doc.Create.NewDetailCurve(view, new_circle)
        line2.LineStyle = Lstyle

    text = str(count)+"x"+str(bar_size)+"-"+str(bar_mark)+"-"+str(round(360.0/count,3)) + u"\u00b0"
    #create textnote placed at the endradius of the bar(line2) with the text variable above as text and oriented at 33.75 degrees
    textNote = TextNote.Create(doc, view.Id, XYZ(endRadius-2.5,0,0), text, ElementId(1018389))
    ElementTransformUtils.RotateElement(doc, textNote.Id, Line.CreateBound(XYZ(0,0,0), XYZ(0,0,1)), math.radians(56.25))

for bar in topRadialList:
    if bar.LookupParameter("Schedule Mark").AsString() not in barlistT2 :
        view.HideElements(List[ElementId]([bar.Id]))

#endregion##########################################################################################################################


t.Commit()

workbook.Close(False)
excel.Quit()