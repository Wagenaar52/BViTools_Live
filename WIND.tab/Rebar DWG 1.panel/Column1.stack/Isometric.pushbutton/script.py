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
view = doc.GetElement(ElementId(2503522))


def polar_to_cartesian(radius, angle):
    x = radius * math.cos(angle)
    y = radius * math.sin(angle)
    return XYZ(x, y, 0)

def slab_height_at_radius(r, rFoundation, rPlinth, hBase, hPlinth, hCone):
    if r <= rFoundation and r > rPlinth:
        return hBase + (hCone - hBase) * ((rFoundation - r) / (rFoundation - rPlinth))
    elif r <= rPlinth:
        return hPlinth

def get_rebar_endpoints(rebar):
    if not isinstance(rebar, Rebar):
        raise TypeError("Expected a Rebar element, got: {}".format(type(rebar)))

    curves = list(rebar.GetCenterlineCurves(
        False, False, False,
        MultiplanarOption.IncludeOnlyPlanarCurves,
        0
    ))

    if not curves:
        return None, None

    start_pt = curves[0].GetEndPoint(0)
    end_pt   = curves[-1].GetEndPoint(1)

    return start_pt, end_pt

def get_rebar_plane_intersection(rebar, plane_normal, plane_origin=XYZ(0, 0, 0)):
    curves = list(rebar.GetCenterlineCurves(
        False, False, False,
        MultiplanarOption.IncludeOnlyPlanarCurves,
        0
    ))

    if not curves:
        return None

    intersections = []

    for curve in curves:
        pts = list(curve.Tessellate())
        if len(pts) < 2:
            continue

        for i in range(len(pts) - 1):
            p1 = pts[i]
            p2 = pts[i + 1]

            d1 = plane_normal.DotProduct(p1 - plane_origin)
            d2 = plane_normal.DotProduct(p2 - plane_origin)

            if abs(d1) < 1e-9:
                intersections.append(p1)
                continue

            if abs(d2) < 1e-9:
                intersections.append(p2)
                continue

            if d1 * d2 < 0:
                t = d1 / (d1 - d2)
                x = p1.X + t * (p2.X - p1.X)
                y = p1.Y + t * (p2.Y - p1.Y)
                z = p1.Z + t * (p2.Z - p1.Z)
                intersections.append(XYZ(x, y, z))

    if not intersections:
        return None

    positive_y = [pt for pt in intersections if pt.Y > 0]
    if positive_y:
        return max(positive_y, key=lambda p: p.Y)

    return intersections[0]

def get_rebar_tag_reference(rebar_element):
    if rebar_element is None or not rebar_element.IsValidObject:
        raise Exception("Invalid rebar element passed to get_rebar_tag_reference")

    subelements = rebar_element.GetSubelements()
    if subelements and len(subelements) > 0:
        return subelements[0].GetReference()

    return Reference(rebar_element)

# FPath = r"C:\Users\Wagner.Human\Downloads\GoldwindTemplate2023\80000.00 - WTFNAME_IMPORT_FILE.xlsx"
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
Bc_barDict = {}
Tc_barDict = {}

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

    elif "BC" in str(xl.Cells(i, 1).Value2).replace(" ","") :
        bar_mark = str(xl.Cells(i,1).Value2)
        startRadius = float(xl.Cells(i,2).Value2)/304.8
        count = float(xl.Cells(i,3).Value2)/304.8
        bar_size = "Y" + str(xl.Cells(i,4).Value2)[1:3]
        bar_dia = int(xl.Cells(i,4).Value2[1:3])/304.8

        bar_parameters = {
            "bar_mark": bar_mark,
            "startRadius": startRadius,
            "count": count,
            "bar_size": bar_size,
            "bar_dia": bar_dia,
        }
        Bc_barDict[bar_mark] = bar_parameters

    elif "TC" in str(xl.Cells(i, 1).Value2).replace(" ","") :
        bar_mark = str(xl.Cells(i,1).Value2)
        startRadius = float(xl.Cells(i,2).Value2)/304.8
        count = float(xl.Cells(i,3).Value2)/304.8
        bar_size = "Y" + str(xl.Cells(i,4).Value2)[1:3]
        bar_dia = int(xl.Cells(i,4).Value2[1:3])/304.8

        bar_parameters = {
            "bar_mark": bar_mark,
            "startRadius": startRadius,
            "count": count,
            "bar_size": bar_size,
            "bar_dia": bar_dia,
        }
        Tc_barDict[bar_mark] = bar_parameters
#botom Radials
barlistBotRad = []
for bar in Br_barDict.keys():
        barlistBotRad.append(Br_barDict[bar]["bar_mark"])

# bottom Concentric
barlistBotCon = []
for bar in Bc_barDict.keys():
        barlistBotCon.append(Bc_barDict[bar]["bar_mark"])

#top Radials
barlistTopRad = []
for bar in Tr_barDict.keys():
        barlistTopRad.append(Tr_barDict[bar]["bar_mark"])

# top Concentric
barlistTopCon = []
for bar in Tc_barDict.keys():
        barlistTopCon.append(Tc_barDict[bar]["bar_mark"])
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
rPlinth = Foundation.LookupParameter("rPlinth").AsDouble()
hCone = Foundation.LookupParameter("hCone").AsDouble()
hPlinth = Foundation.LookupParameter("hPlinth").AsDouble()
hBase = Foundation.LookupParameter("hBase").AsDouble()

bottomRadialList = []
topRadialList = []
bottomRadials = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()

for bar in bottomRadials:
    if "BR" in bar.LookupParameter("Schedule Mark").AsString():
        bottomRadialList.append(bar)
    elif "TR" in bar.LookupParameter("Schedule Mark").AsString():
        topRadialList.append(bar)


#region############ STOOLS ##############################################################################################################################################################

t = Transaction(doc, "Tag Rebar in ISO views")
t.Start()

stool_tags = []

for bar in barlistST:
    Radius = St_barDict[bar]["Radius"]
    bar_dia = St_barDict[bar]["bar_dia"]
    count = St_barDict[bar]["count"]
    bar_size = St_barDict[bar]["bar_size"]
    bar_mark = St_barDict[bar]["bar_mark"]

    tempList = []
    FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
    for stool in FEC:
        if stool.LookupParameter("Schedule Mark").AsString() == bar_mark:
            tempList.append(stool)

    # Select one rebar per bar mark: candidate must be in positive Y, then take largest X.
    positive_y_rebars = []
    for rb in tempList:
        bb = rb.get_BoundingBox(view)
        if bb and bb.Max.Y > 0:
            positive_y_rebars.append(rb)

    if not positive_y_rebars:
        continue

    rebar = max(positive_y_rebars, key=lambda x: x.get_BoundingBox(view).Max.X)
    bb = rebar.get_BoundingBox(view)

    tag = IndependentTag.Create(
            doc,
            ElementId(3488323),
            view.Id,
            get_rebar_tag_reference(rebar),
            True,
            TagOrientation.Horizontal,
            XYZ(bb.Max.X, -2000.0 / 304.8, 0)
    )

    tag.TagHeadPosition = XYZ(bb.Max.X, -2000.0 / 304.8, 0)
    tag.LeaderEndCondition = LeaderEndCondition.Free
    if bb.Min.X > 0:
        Xvalue = bb.Min.X + bar_dia * 1.1
    else:
        Xvalue = 0

    tag_ref = tag.GetTaggedReferences()[0]
    tag.SetLeaderEnd(tag_ref, XYZ(Xvalue, Radius, 0))

    stool_tags.append((tag, tag_ref, rebar, bb))


# Align tag heads on y = -2m and keep each head to the right of its tagged rebar.
if len(stool_tags) > 0:
    y_line = -2000.0 / 304.8
    right_offset = 300.0 / 304.8
    min_spacing = 150.0 / 304.8

    stool_tags_sorted = sorted(stool_tags, key=lambda x: x[3].Max.X)
    previous_x = None
    for tg, tg_ref, _, bb in stool_tags_sorted:
        target_x = bb.Max.X + right_offset
        if previous_x is not None:
            target_x = max(target_x, previous_x + min_spacing)

        z_mid = (bb.Min.Z + bb.Max.Z) / 2.0
        head_pt = XYZ(target_x, y_line, z_mid)
        tg.TagHeadPosition = head_pt

        leader_end = XYZ(bb.Max.X, bb.Max.Y, z_mid)
        tg.LeaderEndCondition = LeaderEndCondition.Free
        tg.SetLeaderEnd(tg_ref, leader_end)

        previous_x = target_x

#endregion

view = doc.GetElement(ElementId(446535))

#region############ TOP RADIAL##############################################################################################################################################################



topRad_tags = []
Radius = rPlinth 
for bar in barlistTopRad:
    Radius += 1
    bar_dia = Tr_barDict[bar]["bar_dia"]
    count = Tr_barDict[bar]["count"]
    bar_size = Tr_barDict[bar]["bar_size"]
    bar_mark = Tr_barDict[bar]["bar_mark"]

    tempList = []
    FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
    for bar in FEC:
        if bar.LookupParameter("Schedule Mark").AsString() == bar_mark:
            tempList.append(bar)

    # Select one rebar per bar mark: start point must be in +Y and +X, then take minimum start Y.
    positive_y_rebarsStart = []
    positive_y_rebarsEnd = []
    for rb in tempList:
        startPoint, endPoint = get_rebar_endpoints(rb)
        if startPoint and startPoint.Y > 0 and startPoint.X > 0:
            positive_y_rebarsStart.append(rb)
        if endPoint and endPoint.Y > 0:
            positive_y_rebarsEnd.append(rb)

    if not positive_y_rebarsStart:
        continue

    rebar = min(positive_y_rebarsStart, key=lambda x: get_rebar_endpoints(x)[0].Y)


    tagStart = IndependentTag.Create(
            doc,
            ElementId(3488323),
            view.Id,
            get_rebar_tag_reference(rebar),
            True,
            TagOrientation.Horizontal,
            get_rebar_endpoints(rebar)[0]
    )
    
    tagEnd = IndependentTag.Create(
            doc,
            ElementId(3488323),
            view.Id,
            get_rebar_tag_reference(rebar),
            True,
            TagOrientation.Horizontal,
            get_rebar_endpoints(rebar)[1]
    )

    tagStart.TagHeadPosition = XYZ(get_rebar_endpoints(rebar)[0].X, -1000.0 / 304.8, 0)
    tagEnd.TagHeadPosition = XYZ(get_rebar_endpoints(rebar)[1].X, -1000.0 / 304.8, 0)
    tagStart.LeaderEndCondition = LeaderEndCondition.Free
    tagEnd.LeaderEndCondition = LeaderEndCondition.Free
    
    tag_ref = tagStart.GetTaggedReferences()[0]
    tagStart.SetLeaderEnd(tag_ref, get_rebar_endpoints(rebar)[0])
    tagEnd.SetLeaderEnd(tag_ref, get_rebar_endpoints(rebar)[1])

    # topRad_tags.append((tagStart, tag_ref, rebar, bb))

#endregion

#region############ TOP CONCENTRIC ##############################################################################################################################################################

topConc_tags = []
Radius = rPlinth 
FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
tempDict = {}
for bar in FEC:
    if "TC" in bar.LookupParameter("Schedule Mark").AsString():
        tempDict[bar] = {
            "schedule_mark": bar.LookupParameter("Schedule Mark").AsString(),
            "intersection": get_rebar_plane_intersection(bar, XYZ(1,0,0))
        }


# now remove any entries where the intersection is None or where the intersection point is not in the positive Y space
tempDict = {bar: data for bar, data in tempDict.items() if data["intersection"] and data["intersection"].Y > 0}
#remove all duplicates based on the schedule mark, keeping the one with the lowest Y value for the intersection point
uniqueMarks = {}
for bar, data in tempDict.items():
    mark = data["schedule_mark"]
    if mark not in uniqueMarks or data["intersection"].Y < uniqueMarks[mark]["intersection"].Y:
        uniqueMarks[mark] = {
            "bar": bar,
            "schedule_mark": mark,
            "intersection": data["intersection"],
        }
tempDict = uniqueMarks

for mark, data in tempDict.items():
    rebar = data["bar"]
    intersection = data["intersection"]
    
    tag = IndependentTag.Create(
            doc,
            ElementId(3488323),
            view.Id,
            get_rebar_tag_reference(rebar),
            True,
            TagOrientation.Horizontal,
            intersection
    )

    Zval = slab_height_at_radius(rebar.LookupParameter("r").AsDouble(), rFoundation, rPlinth, hBase, hPlinth, hCone)

    tag.TagHeadPosition = XYZ( -2000.0 / 304.8, intersection.Y, Zval)
    tag.LeaderEndCondition = LeaderEndCondition.Free

    tag_ref = tag.GetTaggedReferences()[0]
    tag.SetLeaderEnd( tag_ref, XYZ(0, rebar.LookupParameter("r").AsDouble(), Zval))

#endregion

view = doc.GetElement(ElementId(446525))

#region############ BOTTOM RADIAL##############################################################################################################################################################



bottomRad_tags = []
Radius = rPlinth 
for bar in barlistBotRad:
    Radius += 1
    bar_dia = Br_barDict[bar]["bar_dia"]
    count = Br_barDict[bar]["count"]
    bar_size = Br_barDict[bar]["bar_size"]
    bar_mark = Br_barDict[bar]["bar_mark"]

    tempList = []
    FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
    for bar in FEC:
        if bar.LookupParameter("Schedule Mark").AsString() == bar_mark:
            tempList.append(bar)

    # Select one rebar per bar mark: start point must be in +Y and +X, then take minimum start Y.
    positive_y_rebarsStart = []
    positive_y_rebarsEnd = []
    for rb in tempList:
        startPoint, endPoint = get_rebar_endpoints(rb)
        if startPoint and startPoint.Y > 0 and startPoint.X > 0:
            positive_y_rebarsStart.append(rb)
        if endPoint and endPoint.Y > 0:
            positive_y_rebarsEnd.append(rb)

    if not positive_y_rebarsStart:
        continue

    rebar = min(positive_y_rebarsStart, key=lambda x: get_rebar_endpoints(x)[0].Y)


    tagStart = IndependentTag.Create(
            doc,
            ElementId(3488323),
            view.Id,
            get_rebar_tag_reference(rebar),
            True,
            TagOrientation.Horizontal,
            get_rebar_endpoints(rebar)[0]
    )
    
    tagEnd = IndependentTag.Create(
            doc,
            ElementId(3488323),
            view.Id,
            get_rebar_tag_reference(rebar),
            True,
            TagOrientation.Horizontal,
            get_rebar_endpoints(rebar)[1]
    )

    tagStart.TagHeadPosition = XYZ(get_rebar_endpoints(rebar)[0].X, -2000.0 / 304.8, 0)
    tagEnd.TagHeadPosition = XYZ(get_rebar_endpoints(rebar)[1].X, -2000.0 / 304.8, 0)
    tagStart.LeaderEndCondition = LeaderEndCondition.Free
    tagEnd.LeaderEndCondition = LeaderEndCondition.Free
    
    tag_ref = tagStart.GetTaggedReferences()[0]
    tagStart.SetLeaderEnd(tag_ref, get_rebar_endpoints(rebar)[0])
    tagEnd.SetLeaderEnd(tag_ref, get_rebar_endpoints(rebar)[1])

#endregion

#region############ BOTTOM CONCENTRIC ##############################################################################################################################################################

bottomConc_tags = []
Radius = rPlinth 
FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
tempDict = {}
for bar in FEC:
    if "BC" in bar.LookupParameter("Schedule Mark").AsString():
        tempDict[bar] = {
            "schedule_mark": bar.LookupParameter("Schedule Mark").AsString(),
            "intersection": get_rebar_plane_intersection(bar, XYZ(1,0,0))
        }


# now remove any entries where the intersection is None or where the intersection point is not in the positive Y space
tempDict = {bar: data for bar, data in tempDict.items() if data["intersection"] and data["intersection"].Y > 0}
#remove all duplicates based on the schedule mark, keeping the one with the lowest Y value for the intersection point
uniqueMarks = {}
for bar, data in tempDict.items():
    mark = data["schedule_mark"]
    if mark not in uniqueMarks or data["intersection"].Y < uniqueMarks[mark]["intersection"].Y:
        uniqueMarks[mark] = {
            "bar": bar,
            "schedule_mark": mark,
            "intersection": data["intersection"],
        }
tempDict = uniqueMarks

for mark, data in tempDict.items():
    rebar = data["bar"]
    intersection = data["intersection"]
    
    tag = IndependentTag.Create(
            doc,
            ElementId(3488323),
            view.Id,
            get_rebar_tag_reference(rebar),
            True,
            TagOrientation.Horizontal,
            intersection
    )

    Zval = slab_height_at_radius(rebar.LookupParameter("r").AsDouble(), rFoundation, rPlinth, hBase, hPlinth, hCone)

    tag.TagHeadPosition = XYZ( -2000.0 / 304.8, intersection.Y, intersection.Z)
    tag.LeaderEndCondition = LeaderEndCondition.Free

    tag_ref = tag.GetTaggedReferences()[0]
    tag.SetLeaderEnd( tag_ref, XYZ(0, rebar.LookupParameter("r").AsDouble(), intersection.Z))

#endregion

t.Commit()

workbook.Close(False)
excel.Quit()
