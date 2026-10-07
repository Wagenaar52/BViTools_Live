# -*- coding: utf-8 -*-
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

def get_rebar_tag_reference(rebar_element):
    if rebar_element is None or not rebar_element.IsValidObject:
        raise Exception("Invalid rebar element passed to get_rebar_tag_reference")

    subelements = rebar_element.GetSubelements()
    if subelements and len(subelements) > 0:
        return subelements[0].GetReference()

    return Reference(rebar_element)

def get_rebar_references(rebar, view):
    opt = Options()
    opt.ComputeReferences = True
    opt.IncludeNonVisibleObjects = True
    opt.View = view

    ref_array = ReferenceArray()

    for geom_obj in rebar.get_Geometry(opt):
        ref = geom_obj.Reference
        if ref is not None:
            ref_array.Append(ref)
        if ref_array.Size == 2:
            break  # two references is enough for a linear dimension

    if ref_array.Size < 2:
        return None

    return ref_array

FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()

gridList = []
for bar in FEC:
    if "GR" in bar.LookupParameter("Schedule Mark").AsString():
        gridList.append(bar)

t = Transaction(doc, "Tag Rebar in Grid views")
t.Start()

view = doc.GetElement(ElementId(453468))
#region############ GRID 1 ##############################################################################################################################################################



FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
hor_rebar = []
ver_rebar = []

for bar in FEC:
        if "GR100" in bar.LookupParameter("Schedule Mark").AsString():
            startPt = get_rebar_endpoints(bar)[0]
            endPt = get_rebar_endpoints(bar)[1]
            if startPt is None or endPt is None:
                continue
            if abs(startPt.X - endPt.X) < 1e-6 and startPt.X < 0:
                ver_rebar.append(bar)
            elif abs(startPt.Y - endPt.Y) < 1e-6 and startPt.X < 0:
                hor_rebar.append(bar)

#in ver_rebar list, remove all duplicate bars that have the same Schedule Mark(keep the one with the highest max Y value based on the start and end points)
unique_ver_rebar = {}
for bar in ver_rebar:
    mark = bar.LookupParameter("Schedule Mark").AsString()
    startPt = get_rebar_endpoints(bar)[0]
    endPt = get_rebar_endpoints(bar)[1]
    if startPt is None or endPt is None:
        continue
    maxY = max(startPt.Y, endPt.Y)
    if mark not in unique_ver_rebar or maxY > unique_ver_rebar[mark][1]:
        unique_ver_rebar[mark] = (bar, maxY)

positive_y_rebars = [bar_info[0] for bar_info in unique_ver_rebar.values() if bar_info[1] > 0]

for bar in positive_y_rebars:
    tag = IndependentTag.Create(
            doc,
            ElementId(9338727),
            view.Id,
            get_rebar_tag_reference(bar),
            True,
            TagOrientation.Horizontal,
            get_rebar_endpoints(bar)[0])

    tag.TagHeadPosition = XYZ(get_rebar_endpoints(bar)[0].X -0.5, max(get_rebar_endpoints(bar)[0].Y,get_rebar_endpoints(bar)[1].Y)-0.05, 0)
    tag.LeaderEndCondition = LeaderEndCondition.Free

    tag_ref = tag.GetTaggedReferences()[0]
    tag.SetLeaderEnd(tag_ref, XYZ(min(get_rebar_endpoints(bar)[0].X ,get_rebar_endpoints(bar)[1].X ), max(get_rebar_endpoints(bar)[0].Y,get_rebar_endpoints(bar)[1].Y)-0.05, 0))

#in hor_rebar list, remove all duplicate bars that have the same Schedule Mark(keep the one with the highest max Y value based on the start and end points)
unique_hor_rebar = {}
for bar in hor_rebar:
    mark = bar.LookupParameter("Schedule Mark").AsString()
    startPt = get_rebar_endpoints(bar)[0]
    endPt = get_rebar_endpoints(bar)[1]
    if startPt is None or endPt is None:
        continue
    maxY = max(startPt.Y, endPt.Y)
    if mark not in unique_hor_rebar or maxY < unique_hor_rebar[mark][1]:
        unique_hor_rebar[mark] = (bar, maxY)


negative_y_rebars = [bar_info[0] for bar_info in unique_hor_rebar.values() if bar_info[1] < 0]

for bar in negative_y_rebars:
    tag = IndependentTag.Create(
            doc,
            ElementId(9338727),
            view.Id,
            get_rebar_tag_reference(bar),
            True,
            TagOrientation.Horizontal,
            get_rebar_endpoints(bar)[0])

    tag.TagHeadPosition = XYZ(get_rebar_endpoints(bar)[0].X -0.5 , max(get_rebar_endpoints(bar)[0].Y,get_rebar_endpoints(bar)[1].Y), 0)
    tag.LeaderEndCondition = LeaderEndCondition.Free

    tag_ref = tag.GetTaggedReferences()[0]
    tag.SetLeaderEnd(tag_ref, XYZ(min(get_rebar_endpoints(bar)[0].X ,get_rebar_endpoints(bar)[1].X ), max(get_rebar_endpoints(bar)[0].Y,get_rebar_endpoints(bar)[1].Y), 0))


#get the comment parameter of the first bar in hor_rebar list and get the letter that is furthest down the alphabet, store in variable maxLetter
maxLetter = max(bar.LookupParameter("Schedule Mark").AsString()[-1:] for bar in hor_rebar)

maxBarLength = max(bar.LookupParameter("Bar Length").AsDouble() for bar in hor_rebar)
offset = maxBarLength/2 + 2

#vertical dimension

line1 = Line.CreateBound(XYZ(-offset+1,maxBarLength/2,0), XYZ(-offset+0.9,maxBarLength/2,0))
line1 = doc.Create.NewDetailCurve(view, line1)
line2 = Line.CreateBound(XYZ(-offset+1,-maxBarLength/2,0), XYZ(-offset+0.9,-maxBarLength/2,0))
line2 = doc.Create.NewDetailCurve(view, line2)

refArray = ReferenceArray()
refArray.Append(line1.GeometryCurve.Reference)
refArray.Append(line2.GeometryCurve.Reference)
dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(-offset,offset,0), XYZ(-offset,-offset,0)), refArray, doc.GetElement(ElementId(1018370)))
text = "Y{}-""GR100(a-{})-{}".format(bar.LookupParameter("Bar Diameter").AsValueString()[0:3], maxLetter, bar.LookupParameter("Rebar Spacing").AsValueString())
dim.ValueOverride = text

#horizontal dimension

line1 = Line.CreateBound(XYZ(maxBarLength/2,offset-1,0), XYZ(maxBarLength/2,offset+0.1,0))
line1 = doc.Create.NewDetailCurve(view, line1)
line2 = Line.CreateBound(XYZ(-maxBarLength/2,offset-1,0), XYZ(-maxBarLength/2,offset+0.1,0))
line2 = doc.Create.NewDetailCurve(view, line2)

refArray = ReferenceArray()
refArray.Append(line1.GeometryCurve.Reference)
refArray.Append(line2.GeometryCurve.Reference)
dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(offset,offset,0), XYZ(-offset,offset,0)), refArray, doc.GetElement(ElementId(1018370)))
text = "Y{}-""GR100(a-{})-{}".format(bar.LookupParameter("Bar Diameter").AsValueString()[0:3], maxLetter, bar.LookupParameter("Rebar Spacing").AsValueString())
dim.ValueOverride = text
#endregion

view = doc.GetElement(ElementId(453478))
#region############ GRID 2 ##############################################################################################################################################################



FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
hor_rebar = []
ver_rebar = []

for bar in FEC:
        if "GR200" in bar.LookupParameter("Schedule Mark").AsString():
            startPt = get_rebar_endpoints(bar)[0]
            endPt = get_rebar_endpoints(bar)[1]
            if startPt is None or endPt is None:
                continue
            if abs(startPt.X - endPt.X) < 1e-6 and startPt.X < 0:
                ver_rebar.append(bar)
            elif abs(startPt.Y - endPt.Y) < 1e-6 and startPt.X < 0:
                hor_rebar.append(bar)

#in ver_rebar list, remove all duplicate bars that have the same Schedule Mark(keep the one with the highest max Y value based on the start and end points)
unique_ver_rebar = {}
for bar in ver_rebar:
    mark = bar.LookupParameter("Schedule Mark").AsString()
    startPt = get_rebar_endpoints(bar)[0]
    endPt = get_rebar_endpoints(bar)[1]
    if startPt is None or endPt is None:
        continue
    maxY = max(startPt.Y, endPt.Y)
    if mark not in unique_ver_rebar or maxY > unique_ver_rebar[mark][1]:
        unique_ver_rebar[mark] = (bar, maxY)

positive_y_rebars = [bar_info[0] for bar_info in unique_ver_rebar.values() if bar_info[1] > 0]

for bar in positive_y_rebars:
    tag = IndependentTag.Create(
            doc,
            ElementId(9338727),
            view.Id,
            get_rebar_tag_reference(bar),
            True,
            TagOrientation.Horizontal,
            get_rebar_endpoints(bar)[0])

    tag.TagHeadPosition = XYZ(get_rebar_endpoints(bar)[0].X -0.5, max(get_rebar_endpoints(bar)[0].Y,get_rebar_endpoints(bar)[1].Y)-0.05, 0)
    tag.LeaderEndCondition = LeaderEndCondition.Free

    tag_ref = tag.GetTaggedReferences()[0]
    tag.SetLeaderEnd(tag_ref, XYZ(min(get_rebar_endpoints(bar)[0].X ,get_rebar_endpoints(bar)[1].X ), max(get_rebar_endpoints(bar)[0].Y,get_rebar_endpoints(bar)[1].Y)-0.05, 0))

#in hor_rebar list, remove all duplicate bars that have the same Schedule Mark(keep the one with the highest max Y value based on the start and end points)
unique_hor_rebar = {}
for bar in hor_rebar:
    mark = bar.LookupParameter("Schedule Mark").AsString()
    startPt = get_rebar_endpoints(bar)[0]
    endPt = get_rebar_endpoints(bar)[1]
    if startPt is None or endPt is None:
        continue
    maxY = max(startPt.Y, endPt.Y)
    if mark not in unique_hor_rebar or maxY < unique_hor_rebar[mark][1]:
        unique_hor_rebar[mark] = (bar, maxY)


negative_y_rebars = [bar_info[0] for bar_info in unique_hor_rebar.values() if bar_info[1] < 0]

for bar in negative_y_rebars:
    tag = IndependentTag.Create(
            doc,
            ElementId(9338727),
            view.Id,
            get_rebar_tag_reference(bar),
            True,
            TagOrientation.Horizontal,
            get_rebar_endpoints(bar)[0])

    tag.TagHeadPosition = XYZ(get_rebar_endpoints(bar)[0].X -0.5 , max(get_rebar_endpoints(bar)[0].Y,get_rebar_endpoints(bar)[1].Y), 0)
    tag.LeaderEndCondition = LeaderEndCondition.Free

    tag_ref = tag.GetTaggedReferences()[0]
    tag.SetLeaderEnd(tag_ref, XYZ(min(get_rebar_endpoints(bar)[0].X ,get_rebar_endpoints(bar)[1].X ), max(get_rebar_endpoints(bar)[0].Y,get_rebar_endpoints(bar)[1].Y), 0))


#get the comment parameter of the first bar in hor_rebar list and get the letter that is furthest down the alphabet, store in variable maxLetter
maxLetter = max(bar.LookupParameter("Schedule Mark").AsString()[-1:] for bar in hor_rebar)

maxBarLength = max(bar.LookupParameter("Bar Length").AsDouble() for bar in hor_rebar)
offset = maxBarLength/2 + 2

#vertical dimension

line1 = Line.CreateBound(XYZ(-offset+1,maxBarLength/2,0), XYZ(-offset+0.9,maxBarLength/2,0))
line1 = doc.Create.NewDetailCurve(view, line1)
line2 = Line.CreateBound(XYZ(-offset+1,-maxBarLength/2,0), XYZ(-offset+0.9,-maxBarLength/2,0))
line2 = doc.Create.NewDetailCurve(view, line2)

refArray = ReferenceArray()
refArray.Append(line1.GeometryCurve.Reference)
refArray.Append(line2.GeometryCurve.Reference)
dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(-offset,offset,0), XYZ(-offset,-offset,0)), refArray, doc.GetElement(ElementId(1018370)))
text = "Y{}-""GR200(a-{})-{}".format(bar.LookupParameter("Bar Diameter").AsValueString()[0:3], maxLetter, bar.LookupParameter("Rebar Spacing").AsValueString())
dim.ValueOverride = text

#horizontal dimension

line1 = Line.CreateBound(XYZ(maxBarLength/2,offset-1,0), XYZ(maxBarLength/2,offset+0.1,0))
line1 = doc.Create.NewDetailCurve(view, line1)
line2 = Line.CreateBound(XYZ(-maxBarLength/2,offset-1,0), XYZ(-maxBarLength/2,offset+0.1,0))
line2 = doc.Create.NewDetailCurve(view, line2)

refArray = ReferenceArray()
refArray.Append(line1.GeometryCurve.Reference)
refArray.Append(line2.GeometryCurve.Reference)
dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(offset,offset,0), XYZ(-offset,offset,0)), refArray, doc.GetElement(ElementId(1018370)))
text = "Y{}-""GR200(a-{})-{}".format(bar.LookupParameter("Bar Diameter").AsValueString()[0:3], maxLetter, bar.LookupParameter("Rebar Spacing").AsValueString())
dim.ValueOverride = text
#endregion


view = doc.GetElement(ElementId(9137627))
#region############ GRID 3 ##############################################################################################################################################################



FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
hor_rebar = []
ver_rebar = []

for bar in FEC:
        if "GR300" in bar.LookupParameter("Schedule Mark").AsString():
            startPt = get_rebar_endpoints(bar)[0]
            endPt = get_rebar_endpoints(bar)[1]
            if startPt is None or endPt is None:
                continue
            if abs(startPt.X - endPt.X) < 1e-6 and startPt.X < 0:
                ver_rebar.append(bar)
            elif abs(startPt.Y - endPt.Y) < 1e-6 and startPt.X < 0:
                hor_rebar.append(bar)

#in ver_rebar list, remove all duplicate bars that have the same Schedule Mark(keep the one with the highest max Y value based on the start and end points)
unique_ver_rebar = {}
for bar in ver_rebar:
    mark = bar.LookupParameter("Schedule Mark").AsString()
    startPt = get_rebar_endpoints(bar)[0]
    endPt = get_rebar_endpoints(bar)[1]
    if startPt is None or endPt is None:
        continue
    maxY = max(startPt.Y, endPt.Y)
    if mark not in unique_ver_rebar or maxY > unique_ver_rebar[mark][1]:
        unique_ver_rebar[mark] = (bar, maxY)

positive_y_rebars = [bar_info[0] for bar_info in unique_ver_rebar.values() if bar_info[1] > 0]

for bar in positive_y_rebars:
    tag = IndependentTag.Create(
            doc,
            ElementId(9338727),
            view.Id,
            get_rebar_tag_reference(bar),
            True,
            TagOrientation.Horizontal,
            get_rebar_endpoints(bar)[0])

    tag.TagHeadPosition = XYZ(get_rebar_endpoints(bar)[0].X -0.5, max(get_rebar_endpoints(bar)[0].Y,get_rebar_endpoints(bar)[1].Y)-0.05, 0)
    tag.LeaderEndCondition = LeaderEndCondition.Free

    tag_ref = tag.GetTaggedReferences()[0]
    tag.SetLeaderEnd(tag_ref, XYZ(min(get_rebar_endpoints(bar)[0].X ,get_rebar_endpoints(bar)[1].X ), max(get_rebar_endpoints(bar)[0].Y,get_rebar_endpoints(bar)[1].Y)-0.05, 0))

#in hor_rebar list, remove all duplicate bars that have the same Schedule Mark(keep the one with the highest max Y value based on the start and end points)
unique_hor_rebar = {}
for bar in hor_rebar:
    mark = bar.LookupParameter("Schedule Mark").AsString()
    startPt = get_rebar_endpoints(bar)[0]
    endPt = get_rebar_endpoints(bar)[1]
    if startPt is None or endPt is None:
        continue
    maxY = max(startPt.Y, endPt.Y)
    if mark not in unique_hor_rebar or maxY < unique_hor_rebar[mark][1]:
        unique_hor_rebar[mark] = (bar, maxY)


negative_y_rebars = [bar_info[0] for bar_info in unique_hor_rebar.values() if bar_info[1] < 0]

for bar in negative_y_rebars:
    tag = IndependentTag.Create(
            doc,
            ElementId(9338727),
            view.Id,
            get_rebar_tag_reference(bar),
            True,
            TagOrientation.Horizontal,
            get_rebar_endpoints(bar)[0])

    tag.TagHeadPosition = XYZ(get_rebar_endpoints(bar)[0].X -0.5 , max(get_rebar_endpoints(bar)[0].Y,get_rebar_endpoints(bar)[1].Y), 0)
    tag.LeaderEndCondition = LeaderEndCondition.Free

    tag_ref = tag.GetTaggedReferences()[0]
    tag.SetLeaderEnd(tag_ref, XYZ(min(get_rebar_endpoints(bar)[0].X ,get_rebar_endpoints(bar)[1].X ), max(get_rebar_endpoints(bar)[0].Y,get_rebar_endpoints(bar)[1].Y), 0))


#get the comment parameter of the first bar in hor_rebar list and get the letter that is furthest down the alphabet, store in variable maxLetter
maxLetter = max(bar.LookupParameter("Schedule Mark").AsString()[-1:] for bar in hor_rebar)

maxBarLength = max(bar.LookupParameter("Bar Length").AsDouble() for bar in hor_rebar)
offset = maxBarLength/2 + 2

#vertical dimension

line1 = Line.CreateBound(XYZ(-offset+1,maxBarLength/2,0), XYZ(-offset+0.9,maxBarLength/2,0))
line1 = doc.Create.NewDetailCurve(view, line1)
line2 = Line.CreateBound(XYZ(-offset+1,-maxBarLength/2,0), XYZ(-offset+0.9,-maxBarLength/2,0))
line2 = doc.Create.NewDetailCurve(view, line2)

refArray = ReferenceArray()
refArray.Append(line1.GeometryCurve.Reference)
refArray.Append(line2.GeometryCurve.Reference)
dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(-offset,offset,0), XYZ(-offset,-offset,0)), refArray, doc.GetElement(ElementId(1018370)))
text = "Y{}-""GR300(a-{})-{}".format(bar.LookupParameter("Bar Diameter").AsValueString()[0:3], maxLetter, bar.LookupParameter("Rebar Spacing").AsValueString())
dim.ValueOverride = text

#horizontal dimension

line1 = Line.CreateBound(XYZ(maxBarLength/2,offset-1,0), XYZ(maxBarLength/2,offset+0.1,0))
line1 = doc.Create.NewDetailCurve(view, line1)
line2 = Line.CreateBound(XYZ(-maxBarLength/2,offset-1,0), XYZ(-maxBarLength/2,offset+0.1,0))
line2 = doc.Create.NewDetailCurve(view, line2)

refArray = ReferenceArray()
refArray.Append(line1.GeometryCurve.Reference)
refArray.Append(line2.GeometryCurve.Reference)
dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(offset,offset,0), XYZ(-offset,offset,0)), refArray, doc.GetElement(ElementId(1018370)))
text = "Y{}-""GR300(a-{})-{}".format(bar.LookupParameter("Bar Diameter").AsValueString()[0:3], maxLetter, bar.LookupParameter("Rebar Spacing").AsValueString())
dim.ValueOverride = text
#endregion


view = doc.GetElement(ElementId(4593885))
#region############ GRID 4 ##############################################################################################################################################################



FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
hor_rebar = []
ver_rebar = []

for bar in FEC:
        if "GR400" in bar.LookupParameter("Schedule Mark").AsString():
            startPt = get_rebar_endpoints(bar)[0]
            endPt = get_rebar_endpoints(bar)[1]
            if startPt is None or endPt is None:
                continue
            if abs(startPt.X - endPt.X) < 1e-6 and startPt.X < 0:
                ver_rebar.append(bar)
            elif abs(startPt.Y - endPt.Y) < 1e-6 and startPt.X < 0:
                hor_rebar.append(bar)

#in ver_rebar list, remove all duplicate bars that have the same Schedule Mark(keep the one with the highest max Y value based on the start and end points)
unique_ver_rebar = {}
for bar in ver_rebar:
    mark = bar.LookupParameter("Schedule Mark").AsString()
    startPt = get_rebar_endpoints(bar)[0]
    endPt = get_rebar_endpoints(bar)[1]
    if startPt is None or endPt is None:
        continue
    maxY = max(startPt.Y, endPt.Y)
    if mark not in unique_ver_rebar or maxY > unique_ver_rebar[mark][1]:
        unique_ver_rebar[mark] = (bar, maxY)

positive_y_rebars = [bar_info[0] for bar_info in unique_ver_rebar.values() if bar_info[1] > 0]

for bar in positive_y_rebars:
    tag = IndependentTag.Create(
            doc,
            ElementId(9338727),
            view.Id,
            get_rebar_tag_reference(bar),
            True,
            TagOrientation.Horizontal,
            get_rebar_endpoints(bar)[0])

    tag.TagHeadPosition = XYZ(get_rebar_endpoints(bar)[0].X -0.5, max(get_rebar_endpoints(bar)[0].Y,get_rebar_endpoints(bar)[1].Y)-0.05, 0)
    tag.LeaderEndCondition = LeaderEndCondition.Free

    tag_ref = tag.GetTaggedReferences()[0]
    tag.SetLeaderEnd(tag_ref, XYZ(min(get_rebar_endpoints(bar)[0].X ,get_rebar_endpoints(bar)[1].X ), max(get_rebar_endpoints(bar)[0].Y,get_rebar_endpoints(bar)[1].Y)-0.05, 0))

#in hor_rebar list, remove all duplicate bars that have the same Schedule Mark(keep the one with the highest max Y value based on the start and end points)
unique_hor_rebar = {}
for bar in hor_rebar:
    mark = bar.LookupParameter("Schedule Mark").AsString()
    startPt = get_rebar_endpoints(bar)[0]
    endPt = get_rebar_endpoints(bar)[1]
    if startPt is None or endPt is None:
        continue
    maxY = max(startPt.Y, endPt.Y)
    if mark not in unique_hor_rebar or maxY < unique_hor_rebar[mark][1]:
        unique_hor_rebar[mark] = (bar, maxY)


negative_y_rebars = [bar_info[0] for bar_info in unique_hor_rebar.values() if bar_info[1] < 0]

for bar in negative_y_rebars:
    tag = IndependentTag.Create(
            doc,
            ElementId(9338727),
            view.Id,
            get_rebar_tag_reference(bar),
            True,
            TagOrientation.Horizontal,
            get_rebar_endpoints(bar)[0])

    tag.TagHeadPosition = XYZ(get_rebar_endpoints(bar)[0].X -0.5 , max(get_rebar_endpoints(bar)[0].Y,get_rebar_endpoints(bar)[1].Y), 0)
    tag.LeaderEndCondition = LeaderEndCondition.Free

    tag_ref = tag.GetTaggedReferences()[0]
    tag.SetLeaderEnd(tag_ref, XYZ(min(get_rebar_endpoints(bar)[0].X ,get_rebar_endpoints(bar)[1].X ), max(get_rebar_endpoints(bar)[0].Y,get_rebar_endpoints(bar)[1].Y), 0))


#get the comment parameter of the first bar in hor_rebar list and get the letter that is furthest down the alphabet, store in variable maxLetter
maxLetter = max(bar.LookupParameter("Schedule Mark").AsString()[-1:] for bar in hor_rebar)

maxBarLength = max(bar.LookupParameter("Bar Length").AsDouble() for bar in hor_rebar)
offset = maxBarLength/2 + 2

#vertical dimention

line1 = Line.CreateBound(XYZ(-offset+1,maxBarLength/2,0), XYZ(-offset+0.9,maxBarLength/2,0))
line1 = doc.Create.NewDetailCurve(view, line1)
line2 = Line.CreateBound(XYZ(-offset+1,-maxBarLength/2,0), XYZ(-offset+0.9,-maxBarLength/2,0))
line2 = doc.Create.NewDetailCurve(view, line2)

refArray = ReferenceArray()
refArray.Append(line1.GeometryCurve.Reference)
refArray.Append(line2.GeometryCurve.Reference)
dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(-offset,offset,0), XYZ(-offset,-offset,0)), refArray, doc.GetElement(ElementId(1018370)))
text = "Y{}-""GR400(a-{})-{}".format(bar.LookupParameter("Bar Diameter").AsValueString()[0:3], maxLetter, bar.LookupParameter("Rebar Spacing").AsValueString())
dim.ValueOverride = text

#horizontal dimention

line1 = Line.CreateBound(XYZ(maxBarLength/2,offset-1,0), XYZ(maxBarLength/2,offset+0.1,0))
line1 = doc.Create.NewDetailCurve(view, line1)
line2 = Line.CreateBound(XYZ(-maxBarLength/2,offset-1,0), XYZ(-maxBarLength/2,offset+0.1,0))
line2 = doc.Create.NewDetailCurve(view, line2)

refArray = ReferenceArray()
refArray.Append(line1.GeometryCurve.Reference)
refArray.Append(line2.GeometryCurve.Reference)
dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(offset,offset,0), XYZ(-offset,offset,0)), refArray, doc.GetElement(ElementId(1018370)))
text = "Y{}-""GR400(a-{})-{}".format(bar.LookupParameter("Bar Diameter").AsValueString()[0:3], maxLetter, bar.LookupParameter("Rebar Spacing").AsValueString())
dim.ValueOverride = text
#endregion


t.Commit()

