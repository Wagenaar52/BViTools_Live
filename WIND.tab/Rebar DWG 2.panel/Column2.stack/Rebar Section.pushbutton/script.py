# -*- coding: utf-8 -*-
import clr
import math
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from collections import OrderedDict
from pyrevit import forms
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel


uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
view = doc.GetElement(ElementId(4355490))

text_note_type_id = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType)

pick_file_fn = getattr(forms, 'pick_file', None)
if callable(pick_file_fn):
    FPath = pick_file_fn(file_ext='xlsx', multi_file=False, unc_paths=False)
elif isinstance(pick_file_fn, str):
    FPath = pick_file_fn
else:
    FPath = None

if not FPath:
    raise SystemExit

def get_safe_tag_head_position(rebar_element, plane_normal, plane_origin=XYZ(0, 0, 0)):
    intersection_pt = get_rebar_plane_intersection(rebar_element, plane_normal, plane_origin)
    if intersection_pt is not None:
        return intersection_pt
    return get_rebar_midpoint(rebar_element)

def get_rebar_plane_intersection(rebar, plane_normal, plane_origin=XYZ(0, 0, 0)):
    curves = list(rebar.GetCenterlineCurves(
        False, False, False,
        Structure.MultiplanarOption.IncludeOnlyPlanarCurves,
        0
    ))

    if not curves:
        return None

    intersections = []

    for curve in curves:
        # Sample curve points by normalized parameter for denser arc handling.
        # Fall back to Tessellate if Evaluate sampling cannot be used.
        pts = []
        try:
            step_ft = 100.0 / 304.8
            sample_count = max(12, int(curve.Length / step_ft) + 1)
            for idx in range(sample_count + 1):
                t_norm = float(idx) / float(sample_count)
                pts.append(curve.Evaluate(t_norm, True))
        except Exception:
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

    negative_x = [pt for pt in intersections if pt.X < 0]
    if negative_x:
        return max(negative_x, key=lambda p: p.X)

    return intersections[0]

def get_rebar_midpoint(rebar_element):
    if rebar_element is None or not rebar_element.IsValidObject:
        raise Exception("Invalid rebar element passed to get_rebar_midpoint")

    bbox = rebar_element.get_BoundingBox(None)
    if bbox is not None:
        return XYZ(
            (bbox.Min.X + bbox.Max.X) / 2.0,
            (bbox.Min.Y + bbox.Max.Y) / 2.0,
            (bbox.Min.Z + bbox.Max.Z) / 2.0
        )

    raise Exception("Could not determine midpoint for rebar element")

def get_rebar_intersection_on_y0_most_negative_x(rebar):
    curves = list(rebar.GetCenterlineCurves(
        False, False, False,
        Structure.MultiplanarOption.IncludeOnlyPlanarCurves,
        0
    ))

    if not curves:
        return None

    intersections = []
    tol = 1e-9

    for curve in curves:
        pts = []
        try:
            step_ft = 100.0 / 304.8
            sample_count = max(12, int(curve.Length / step_ft) + 1)
            for idx in range(sample_count + 1):
                t_norm = float(idx) / float(sample_count)
                pts.append(curve.Evaluate(t_norm, True))
        except Exception:
            pts = list(curve.Tessellate())

        if len(pts) < 2:
            continue

        for i in range(len(pts) - 1):
            p1 = pts[i]
            p2 = pts[i + 1]

            y1 = p1.Y
            y2 = p2.Y

            if abs(y1) < tol and abs(y2) < tol:
                intersections.append(p1 if p1.X <= p2.X else p2)
                continue

            if abs(y1) < tol:
                intersections.append(p1)
                continue

            if abs(y2) < tol:
                intersections.append(p2)
                continue

            if y1 * y2 < 0:
                t = y1 / (y1 - y2)
                x = p1.X + t * (p2.X - p1.X)
                z = p1.Z + t * (p2.Z - p1.Z)
                intersections.append(XYZ(x, 0.0, z))

    if not intersections:
        return None

    return min(intersections, key=lambda p: p.X)

def get_rebar_tag_reference(rebar_element):
    if rebar_element is None or not rebar_element.IsValidObject:
        raise Exception("Invalid rebar element passed to get_rebar_tag_reference")

    subelements = rebar_element.GetSubelements()
    if subelements and len(subelements) > 0:
        return subelements[0].GetReference()

    return Reference(rebar_element)

def place_cl_annotation(location):
    cl_anno_symbol = doc.GetElement(ElementId(1124927))
    if cl_anno_symbol is None:
        return None

    if not cl_anno_symbol.IsActive:
        cl_anno_symbol.Activate()
        doc.Regenerate()

    cl_anno = doc.Create.NewFamilyInstance(location + XYZ(-0.143784846, 0, 0.129055994), cl_anno_symbol, view)
    element_name_param = cl_anno.LookupParameter("Element Name")
    if element_name_param is not None:
        element_name_param.Set("FOUNDATION")

    return cl_anno

def Tag_pc_bars_in_view(tag_bars, view):
    # Build the requested PC mark set from the input list.
    requested_marks = set()
    for bar in tag_bars:
        mark_param = bar.LookupParameter("Schedule Mark")
        if mark_param is None:
            continue
        bar_mark = mark_param.AsString()
        if bar_mark and "PC" in bar_mark:
            requested_marks.add(bar_mark)

    if not requested_marks:
        return tag_bars

    # Search all project rebars to find, for each mark, the bar that intersects Y=0
    # at the most negative X location.
    all_rebars = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
    bars_by_mark = {}
    for bar in all_rebars:
        mark_param = bar.LookupParameter("Schedule Mark")
        if mark_param is None:
            continue
        bar_mark = mark_param.AsString()
        if bar_mark in requested_marks:
            if bar_mark not in bars_by_mark:
                bars_by_mark[bar_mark] = []
            bars_by_mark[bar_mark].append(bar)

    selected_pc_bars = {}
    for bar_mark in requested_marks:
        bars = bars_by_mark.get(bar_mark, [])
        selected_bar = None
        selected_pt = None

        for bar in bars:
            intersection_pt = get_rebar_intersection_on_y0_most_negative_x(bar)
            if intersection_pt is None:
                continue
            if selected_pt is None or intersection_pt.X < selected_pt.X:
                selected_bar = bar
                selected_pt = intersection_pt

        if selected_bar is None:
            # Fallback to input list for this mark if no Y=0 intersection is found.
            for in_bar in tag_bars:
                in_mark_param = in_bar.LookupParameter("Schedule Mark")
                if in_mark_param is not None and in_mark_param.AsString() == bar_mark:
                    selected_bar = in_bar
                    selected_pt = get_rebar_midpoint(in_bar)
                    break

        if selected_bar is not None:
            if selected_pt is None:
                selected_pt = get_rebar_midpoint(selected_bar)
            selected_pc_bars[bar_mark] = (selected_bar, selected_pt)

    # Tag one selected bar per mark at the selected Y=0 intersection point.
    for key in selected_pc_bars.keys():
        bar, tag_point = selected_pc_bars[key]
        bar_mark = bar.LookupParameter("Schedule Mark").AsString()
        if bar_mark[-2:] == "00" or bar_mark[-2:] == "01":
            tagType = ElementId(3488323)
            tag_point = tag_point +XYZ(0,10,0)
            leader_extension = 1.0
        else:
            tagType = ElementId(9387342)# 3488323) 
            leader_extension = 0.0

        if len(bar_mark) >= 2 and bar_mark[-2] == "0":
            tag_head_position = tag_point + XYZ(0.127959661, 0, -(0.3 + leader_extension))
        else:
            tag_head_position = tag_point + XYZ(0.127959661, 0, 0.3 + leader_extension) 

        tag = IndependentTag.Create(
                doc,
                tagType,
                view.Id,
                get_rebar_tag_reference(bar),
                True,
                TagOrientation.Horizontal,
                tag_point
        )
        tag.TagHeadPosition = tag_head_position
        tag.LeaderEndCondition = LeaderEndCondition.Free

        tag_ref = tag.GetTaggedReferences()[0]
        tag.SetLeaderElbow(tag_ref, tag_head_position)
        tag.SetLeaderEnd(tag_ref, tag_point)

    return tag_bars

def Tag_ra_bars_in_view(tag_bars, view):
    # Build the requested RA mark set from the input list.
    requested_marks = set()
    for bar in tag_bars:
        mark_param = bar.LookupParameter("Schedule Mark")
        if mark_param is None:
            continue
        bar_mark = mark_param.AsString()
        if bar_mark and "RA" in bar_mark:
            requested_marks.add(bar_mark)

    if not requested_marks:
        return tag_bars

    # Search all project rebars to find, for each mark, the bar that intersects Y=0
    # at the most negative X location.
    all_rebars = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
    bars_by_mark = {}
    for bar in all_rebars:
        mark_param = bar.LookupParameter("Schedule Mark")
        if mark_param is None:
            continue
        bar_mark = mark_param.AsString()
        if bar_mark in requested_marks:
            if bar_mark not in bars_by_mark:
                bars_by_mark[bar_mark] = []
            bars_by_mark[bar_mark].append(bar)

    selected_ra_bars = {}
    for bar_mark in requested_marks:
        bars = bars_by_mark.get(bar_mark, [])
        selected_bar = None
        selected_pt = None

        for bar in bars:
            intersection_pt = get_rebar_intersection_on_y0_most_negative_x(bar)
            if intersection_pt is None:
                continue
            if selected_pt is None or intersection_pt.X < selected_pt.X:
                selected_bar = bar
                selected_pt = intersection_pt

        if selected_bar is None:
            # Fallback to input list for this mark if no Y=0 intersection is found.
            for in_bar in tag_bars:
                in_mark_param = in_bar.LookupParameter("Schedule Mark")
                if in_mark_param is not None and in_mark_param.AsString() == bar_mark:
                    selected_bar = in_bar
                    selected_pt = get_rebar_midpoint(in_bar)
                    break

        if selected_bar is not None:
            if selected_pt is None:
                selected_pt = get_rebar_midpoint(selected_bar)
            selected_ra_bars[bar_mark] = (selected_bar, selected_pt)

    # Tag one selected bar per mark at the selected Y=0 intersection point.
    for key in selected_ra_bars.keys():
        bar, tag_point = selected_ra_bars[key]
        bar_mark = bar.LookupParameter("Schedule Mark").AsString()
        if bar_mark[-2:] == "00" or bar_mark[-2:] == "01":
            tagType = ElementId(3488323)
            leader_extension = 1
        else:
            tagType = ElementId(3488323) 
            leader_extension = 1

        if len(bar_mark) >= 2 and bar_mark[-2] == "0":
            tag_head_position = tag_point + XYZ(leader_extension, 0, 0)
        else:
            tag_head_position = tag_point + XYZ(leader_extension, 0, 0) 

        tag = IndependentTag.Create(
                doc,
                tagType,
                view.Id,
                get_rebar_tag_reference(bar),
                True,
                TagOrientation.Horizontal,
                tag_point
        )
        tag.TagHeadPosition = tag_head_position
        tag.LeaderEndCondition = LeaderEndCondition.Free

        tag_ref = tag.GetTaggedReferences()[0]
        tag.SetLeaderElbow(tag_ref, tag_head_position)
        tag.SetLeaderEnd(tag_ref, tag_point)

    return tag_bars

def get_rebar_endpoints(rebar):
    if not isinstance(rebar, Structure.Rebar):
        raise TypeError("Expected a Rebar element, got: {}".format(type(rebar)))

    curves = list(rebar.GetCenterlineCurves(
        False, False, False,
        Structure.MultiplanarOption.IncludeOnlyPlanarCurves,
        0
    ))

    if not curves:
        return None, None

    start_pt = curves[0].GetEndPoint(0)
    end_pt   = curves[-1].GetEndPoint(1)

    return start_pt, end_pt

def get_face_reference_by_geometry_id(element, face_id):
    options = Options()
    options.ComputeReferences = True
    options.IncludeNonVisibleObjects = True

    def find_in_geometry(geometry_element):
        if geometry_element is None:
            return None

        for geometry_object in geometry_element:
            solid = geometry_object if isinstance(geometry_object, Solid) else None
            if solid is not None and solid.Faces.Size > 0:
                for face in solid.Faces:
                    if face.Id == face_id and face.Reference is not None:
                        return face.Reference

            if isinstance(geometry_object, GeometryInstance):
                reference = find_in_geometry(geometry_object.GetInstanceGeometry())
                if reference is not None:
                    return reference

        return None

    return find_in_geometry(element.get_Geometry(options))

def triangle_detail(hCone, hBase,rPlinth,rBase,spacing,x,y):
    # draw a triagle and dimention the skew side and under side 
    slope = (hCone-hBase)/(rBase-rPlinth)
    pt1 = XYZ(x, 0, y)
    pt2 = XYZ(x+float(spacing), 0, y)
    pt3 = XYZ(x+float(spacing), 0, y+(float(spacing)*slope))

    line1 = doc.Create.NewDetailCurve(view, Line.CreateBound(pt1, pt2))
    line2 = doc.Create.NewDetailCurve(view, Line.CreateBound(pt2, pt3))
    line3 = doc.Create.NewDetailCurve(view, Line.CreateBound(pt3, pt1))

    refArray = ReferenceArray()
    base_curve = line1.GeometryCurve
    refArray.Append(base_curve.GetEndPointReference(0))
    refArray.Append(base_curve.GetEndPointReference(1))
    dim = doc.Create.NewDimension(view, Line.CreateBound(pt1 + XYZ(0,0,-0.5), pt2 + XYZ(0,0,-0.5)), refArray, doc.GetElement(ElementId(1018370)))
    
    refArray = ReferenceArray()
    skew_curve = line3.GeometryCurve
    refArray.Append(skew_curve.GetEndPointReference(0))
    refArray.Append(skew_curve.GetEndPointReference(1))
    perVector = XYZ.BasisY.CrossProduct(skew_curve.Direction).Normalize()
    dim = doc.Create.NewDimension(view, Line.CreateBound(pt1+0.5*perVector, pt3+ 0.5*perVector), refArray, doc.GetElement(ElementId(1018370)))
    dim.TextPosition = dim.TextPosition + perVector*(-0.4)


excel = Excel.ApplicationClass()
excel.Visible = False
workbook = excel.Workbooks.Open(FPath)
xl = workbook.Worksheets['A']

# get all inputs form excel in a dictionary
BarDict_TR = {}
for i in range(1,200):
    i += 1
    if "TR" in str(xl.Cells(i, 1).Value2).replace(" ","") :
        bar_mark = str(xl.Cells(i,1).Value2)
        rStart = float(xl.Cells(i,2).Value2)/304.8
        rEnd = float(xl.Cells(i,3).Value2)/304.8
        Size = str(xl.Cells(i,4).Value2)
        count = int(xl.Cells(i,8).Value2)
        print(bar_mark)
        BarDict_TR[bar_mark] = [rStart, rEnd, count, Size]
        #print(BarDict_TR[bar_mark])


BarDict_BR = {}
for i in range(1,200):
    i += 1
    if "BR" in str(xl.Cells(i, 1).Value2).replace(" ","") :
        bar_mark = str(xl.Cells(i,1).Value2)
        rStart = float(xl.Cells(i,2).Value2)/304.8
        rEnd = float(xl.Cells(i,3).Value2)/304.8
        Size = str(xl.Cells(i,4).Value2)
        count = int(xl.Cells(i,8).Value2)
        print(bar_mark)
        BarDict_BR[bar_mark] = [rStart, rEnd, count, Size]


BarDict_TC = {}
for i in range(1,200):
    i += 1
    if "TC" in str(xl.Cells(i, 1).Value2).replace(" ","") :
        bar_mark = str(xl.Cells(i,1).Value2)
        rStart = float(xl.Cells(i,2).Value2)/304.8
        rEnd = float(xl.Cells(i,3).Value2)/304.8
        spacing = float(xl.Cells(i,7).Value2)
        Size = str(xl.Cells(i,4).Value2)
        print(bar_mark)
        BarDict_TC[bar_mark] = [rStart, rEnd, spacing, Size]

BarDict_BC = {}
for i in range(1,200):
    i += 1
    if "BC" in str(xl.Cells(i, 1).Value2).replace(" ","") :
        bar_mark = str(xl.Cells(i,1).Value2)
        rStart = float(xl.Cells(i,2).Value2)/304.8
        rEnd = float(xl.Cells(i,3).Value2)/304.8
        Size = str(xl.Cells(i,4).Value2)
        spacing = float(xl.Cells(i,7).Value2)
        print(bar_mark)
        BarDict_BC[bar_mark] = [rStart, rEnd, spacing, Size]


# Get Elements from document

FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
for elem in FEC:
    if elem.Name == '1PA_WTF_SteelTower':
        wtf_steel = elem

rPlinth = wtf_steel.LookupParameter('rPlinth').AsDouble()
rBase = wtf_steel.LookupParameter('rBase').AsDouble()
hPlinth = wtf_steel.LookupParameter('hPlinth').AsDouble()
hBase = wtf_steel.LookupParameter('hBase').AsDouble()
hCone  = wtf_steel.LookupParameter('hCone').AsDouble()
hPit = wtf_steel.LookupParameter('hBottomVoid').AsDouble()
rPitIn = wtf_steel.LookupParameter('rVoidInner').AsDouble()
rPitOut = wtf_steel.LookupParameter('rVoidOuter').AsDouble()
slabSlope = ((hCone-hBase)/(rBase - rPlinth))
wGroutTop = wtf_steel.LookupParameter('wGroutTop').AsDouble()
rTower = wtf_steel.LookupParameter('rTower').AsDouble()

FEC_dl = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Lines).WhereElementIsNotElementType().ToElements()
# for elem in FEC_dl:
    # if elem.OwnerViewId == ElementId(6085766):
        # print(elem.Name)
        # doc.Delete(elem.Id)
        # print("Deleted")

# Sort BarDict by 'rEnd' (index 1 in the list)
sorted_BarDict_TR = OrderedDict(sorted(BarDict_TR.items(), key=lambda item: item[1][1]))
sorted_BarDict_BR = OrderedDict(sorted(BarDict_BR.items(), key=lambda item: item[1][1]))
sorted_BarDict_TC = OrderedDict(sorted(BarDict_TC.items(), key=lambda item: item[1][1]))
sorted_BarDict_BC = OrderedDict(sorted(BarDict_BC.items(), key=lambda item: item[1][1]))


FEC_TC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
TC_BM_List = []
for elem in FEC_TC:
    if "TC" in elem.LookupParameter('Schedule Mark').AsString():
        if elem.LookupParameter('Schedule Mark').AsString() not in TC_BM_List:
            TC_BM_List.append(elem.LookupParameter('Schedule Mark').AsString())
TC_BM_List.sort()       

FEC_BC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
BC_BM_List = []
for elem in FEC_BC:
    if "BC" in elem.LookupParameter('Schedule Mark').AsString():
        if elem.LookupParameter('Schedule Mark').AsString() not in BC_BM_List:
            BC_BM_List.append(elem.LookupParameter('Schedule Mark').AsString())
BC_BM_List.sort()




TC_Dict = {}
for i in range(9):
    TC_Dict["TC"+str(i+1)] = []
for bm in TC_BM_List:
        TC_Dict[str(bm)[:3]].append(bm)


BC_Dict = {}
for i in range(9):
    BC_Dict["BC"+str(i+1)] = []
print("**************")
print(BC_Dict)
for bm in BC_BM_List:
        BC_Dict[str(bm)[:3]].append(bm)
print("**************")
print(BC_Dict)
print("**************")





# Get the default text note type
text_note_type_id = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType)

t = Transaction(doc)
t.Start("Section Anotation")

### Draw Center Line
# Define Points for Center Line
cl_top = XYZ(0, 0, hPlinth+3)
cl_bot = XYZ(0, 0, -hPit-2)
# Create Center Line
CL = doc.Create.NewDetailCurve(view, Line.CreateBound(cl_top, cl_bot))
CL.LineStyle = doc.GetElement(ElementId(1018897))
PL = doc.Create.NewDetailCurve(view, Line.CreateBound(XYZ(-rPlinth, 0, hPlinth-1), XYZ(-rPlinth, 0, hPlinth-1.2)))
#place CL annotation at center line top
place_cl_annotation(cl_top)

#region############################################## TC ###############################################
hcounter = 0
hidLineList = []
for bar_mark, values in sorted_BarDict_TC.items():
    rStart = values[0]
    rEnd = values[1]
    spacing = str(values[2])[:-2]
    Size = values[3]
    # Define Points for Bar
    bar_top_end = XYZ(-rEnd, 0, hCone+1.5)
    if rEnd >= rPlinth:
        bar_bot_end = XYZ(-rEnd, 0, 1 + hBase + (rBase-rEnd)*slabSlope)
    # Create hidden line 
    endHidLine = doc.Create.NewDetailCurve(view, Line.CreateBound(bar_top_end, bar_bot_end))
    endHidLine.LineStyle = doc.GetElement(ElementId(1019189))
    hidLineList.append(endHidLine)
    # Create the dimension
    refArray = ReferenceArray()
    refArray.Append(CL.GeometryCurve.Reference)
    refArray.Append(endHidLine.GeometryCurve.Reference)
    dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(0,0,hPlinth+500/304.8+hcounter),XYZ(rEnd,0,hPlinth+500/304.8+hcounter)), refArray, doc.GetElement(ElementId(1018370)))
    dim.TextPosition = dim.TextPosition + XYZ(0, 0, -0.04750)
    tagBarMark = 0
    subElements = len(TC_Dict[bar_mark[:3]])
    if subElements == 1:
        tagBarMark = bar_mark
    elif subElements > 1:
        tagBarMark = "TC("+str(bar_mark)[2]+"01-"+str(bar_mark)[2]+"0"+str(subElements)+")"
    tagText = str(Size)+"-"+str(tagBarMark)+"-"+str(spacing)
    textXposition = -((rStart+rEnd)/2)-((len(tagText)*35/304.8)/2)
    tag_text_note = TextNote.Create(doc, view.Id, XYZ(textXposition,0,hCone+1.25), tagText, text_note_type_id)
    tag_text_note.TextNoteType = doc.GetElement(ElementId(1018389))
    hcounter += 120/304.8
    rMid = (rStart + rEnd)/2
    hMid = hCone -1.5 - ((rMid - rPlinth)*slabSlope)
    spacing_ft = float(spacing)/304.8
    triangle_detail(hCone, hBase,rPlinth,rBase,spacing_ft,textXposition,hMid)


refArray = ReferenceArray()
refArray.Append(PL.GeometryCurve.Reference)
for line in hidLineList:
    refArray.Append(line.GeometryCurve.Reference)
dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(0,0,hCone+1.5),XYZ(-rBase,0,hCone+500/304.8)), refArray, doc.GetElement(ElementId(1018370)))
# dim.TextPosition = dim.TextPosition + XYZ(0, 0, -0.04750)

# triangle detail




#endregion

#region############################################## BC ###############################################

hcounter = 0
hidLineList = []
for bar_mark, values in sorted_BarDict_BC.items():
    rStart = values[0]
    rEnd = values[1]
    spacing = str(values[2])[:-2]
    Size = values[3]

    # Define Points for Bar
    bar_bot_end = XYZ(-rEnd, 0, -hPit-0.5)
    if rEnd >= rPlinth:
        bar_top_end = XYZ(-rEnd, 0, -0.25)
    elif rEnd < rPlinth:
        bar_top_end = XYZ(-rEnd, 0, -hPit-0.25)
    # Create hidden line 
    endHidLine = doc.Create.NewDetailCurve(view, Line.CreateBound(bar_top_end, bar_bot_end))
    endHidLine.LineStyle = doc.GetElement(ElementId(1019189))
    hidLineList.append(endHidLine)
    # Create the dimension
    refArray = ReferenceArray()
    refArray.Append(CL.GeometryCurve.Reference)
    refArray.Append(endHidLine.GeometryCurve.Reference)
    dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(0,0,-hPit-1+hcounter),XYZ(rEnd,0,-hPit-1+hcounter)), refArray, doc.GetElement(ElementId(1018370)))
    dim.TextPosition = dim.TextPosition + XYZ(0, 0, -0.04750)
    subElements = len(BC_Dict[bar_mark[:3]])
    tagBarMark
    if subElements == 1:
        tagBarMark = bar_mark
    elif subElements > 1:
        tagBarMark = "BC("+str(bar_mark)[2]+"01-"+str(bar_mark)[2]+"0"+str(subElements)+")"
    tagText = str(Size)+"-"+str(tagBarMark)+"-"+str(spacing)
    textXposition = -((rStart+rEnd)/2)-((len(tagText)*35/304.8)/2)
    tag_text_note = TextNote.Create(doc, view.Id, XYZ(textXposition,0,-hPit-0.75), tagText, text_note_type_id)
    tag_text_note.TextNoteType = doc.GetElement(ElementId(1018389))

    hcounter -= 120/304.8

refArray = ReferenceArray()
refArray.Append(PL.GeometryCurve.Reference)
for line in hidLineList:
    refArray.Append(line.GeometryCurve.Reference)
dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(0,0,-hPit-0.5),XYZ(-rBase,0,-hPit-0.5)), refArray, doc.GetElement(ElementId(1018370)))
# dim.TextPosition = dim.TextPosition + XYZ(0, 0, -0.04750)
#endregion

#region############################################## TR ###############################################
hcounter = 0
hScounter = 0
rEndOld = 0
rStartOld = 0
hTagcounter = -120/304.8
hStartTagcounter = -120/304.8

for bar_mark, values in sorted_BarDict_TR.items():
    rStart = values[0]
    rEnd = values[1]
    count = values[2]
    Size = values[3]
    
    angle = str(float(360)/float(count))+"°"

    ####START 
    # Define Points for Bar
    bar_top_start = XYZ(rStart, 0, hPlinth-3)
    if rStart < rPlinth:
        bar_bot_start = XYZ(rStart, 0, hCone-1)
    elif rStart >= rPlinth:
        bar_bot_start = XYZ(rStart, 0, 1 + hBase + (rBase-rStart)*slabSlope)
    # Create hidden line 
    startHidLine = doc.Create.NewDetailCurve(view, Line.CreateBound(bar_top_start, bar_bot_start))
    startHidLine.LineStyle = doc.GetElement(ElementId(1019189))
    # Create the text note
    text_note = TextNote.Create(doc, view.Id, XYZ(bar_top_start.X+0.1,0,hPlinth-3), "START", text_note_type_id)
    text_note.TextNoteType = doc.GetElement(ElementId(1018389))
    # text_note.Coord = XYZ(text_note.Coord.X, text_note.Coord.Y-1, text_note.Coord.Z)
    ####
    if rStart <= rPlinth:
        if rStart != rStartOld:
            hStartTagcounter = hStartTagcounter-120/304.8
        tag_text_note = TextNote.Create(doc, view.Id, XYZ(rPlinth+1,0,hPlinth-3+hStartTagcounter), str(count)+"x"+str(Size)+"-"+str(bar_mark)+"-"+str(angle), text_note_type_id)
        tag_text_note.TextNoteType = doc.GetElement(ElementId(1018389))
        tag_text_note.AddLeader(TextNoteLeaderTypes.TNLT_STRAIGHT_L)
        tag_text_note.GetLeaders()[0].End = XYZ(rStart,0,+hPlinth-3+hStartTagcounter-45/304.8)
        hScounter += 120/304.8
        rStartOld = rStart
    elif rStart > rPlinth:
        tag_text_note = TextNote.Create(doc, view.Id, XYZ(bar_top_start.X+0.25,0,hPlinth-3+hTagcounter), str(count)+"x"+str(Size)+"-"+str(bar_mark)+"-"+str(angle), text_note_type_id)
        tag_text_note.TextNoteType = doc.GetElement(ElementId(1018389))
        tag_text_note.AddLeader(TextNoteLeaderTypes.TNLT_STRAIGHT_L)
        tag_text_note.GetLeaders()[0].End = XYZ(bar_top_start.X,0,hPlinth-3+hTagcounter-45/304.8)
        hScounter += 120/304.8
        rStartOld = rStart

    ####END 
    ####
    # Define Points for Bar
    bar_top_end = XYZ(rEnd, 0, hPlinth-3)
    if rEnd < rPlinth:
        bar_bot_end = XYZ(rEnd, 0, hCone-1)
    elif rEnd >= rPlinth:
        bar_bot_end = XYZ(rEnd, 0, 1 + hBase + (rBase-rEnd)*slabSlope)
    # Create hidden line 
    endHidLine = doc.Create.NewDetailCurve(view, Line.CreateBound(bar_top_end, bar_bot_end))
    endHidLine.LineStyle = doc.GetElement(ElementId(1019189))
    if rEnd != rEndOld:
        # Create the text note
        text_note = TextNote.Create(doc, view.Id, XYZ(bar_top_end.X-0.53,0,hPlinth-3),"END", text_note_type_id)
        text_note.TextNoteType = doc.GetElement(ElementId(1018389))
        # Create the dimension
        refArray = ReferenceArray()
        refArray.Append(CL.GeometryCurve.Reference)
        refArray.Append(endHidLine.GeometryCurve.Reference)
        dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(0,0,hPlinth+500/304.8+hcounter),XYZ(rEnd,0,hPlinth+500/304.8+hcounter)), refArray, doc.GetElement(ElementId(1018370)))
        dim.TextPosition = dim.TextPosition + XYZ(0, 0, -0.04750)
    if rEnd == rEndOld:
        hTagcounter = hTagcounter-120/304.8
    tag_text_note = TextNote.Create(doc, view.Id, XYZ(bar_top_end.X-2.25,0,hPlinth-3+hTagcounter), str(count)+"x"+str(Size)+"-"+str(bar_mark)+"-"+str(angle), text_note_type_id)
    tag_text_note.TextNoteType = doc.GetElement(ElementId(1018389))
    tag_text_note.AddLeader(TextNoteLeaderTypes.TNLT_STRAIGHT_R)
    tag_text_note.GetLeaders()[0].End = XYZ(bar_top_end.X,0,+hPlinth-3+hTagcounter-45/304.8)
    hcounter += 120/304.8
    rEndOld = rEnd
#endregion

#region############################################## BR ###############################################
hcounter = -120/304.8
rEndOld = 0
hTagcounter = -120/304.8

for bar_mark, values in sorted_BarDict_BR.items():
    rStart = values[0]
    rEnd = values[1]
    count = values[2]
    Size = values[3]
    
    angle = str(float(360)/float(count))+"°"

    ####START 
    # Define Points for Bar
    bar_bot_start = XYZ(rStart, 0, -hPit-1)
    if rStart < rPlinth:
        bar_top_start = XYZ(rStart, 0, 700/304.8)
    elif rStart >= rPlinth:
        bar_top_start = XYZ(rStart, 0, -1)
    # Create hidden line 
    startHidLine = doc.Create.NewDetailCurve(view, Line.CreateBound(bar_top_start, bar_bot_start))
    startHidLine.LineStyle = doc.GetElement(ElementId(1019189))
    # Create the text note
    text_note = TextNote.Create(doc, view.Id, XYZ(bar_top_start.X+0.1,0,-hPit -0.1), "START", text_note_type_id)
    text_note.TextNoteType = doc.GetElement(ElementId(1018389))
    # Create the dimension
    refArray = ReferenceArray()
    refArray.Append(CL.GeometryCurve.Reference)
    refArray.Append(startHidLine.GeometryCurve.Reference)
    if rStart <= rPlinth:
        dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(0,0,700/304.8-hcounter),XYZ(rStart,0,700/304.8-hcounter)), refArray, doc.GetElement(ElementId(1018370)))
        dim.TextPosition = dim.TextPosition + XYZ(0, 0, -0.04750)
        hStartTagcounter = -120/304.8
        if rEnd == rEndOld:
            hStartTagcounter = hStartTagcounter-120/304.8
        tag_text_note = TextNote.Create(doc, view.Id, XYZ(bar_top_start.X+0.25,0,-hPit-0.25+hStartTagcounter+45.5/304.8), str(count)+"x"+str(Size)+"-"+str(bar_mark)+"-"+str(angle), text_note_type_id)
        tag_text_note.TextNoteType = doc.GetElement(ElementId(1018389))
        tag_text_note.AddLeader(TextNoteLeaderTypes.TNLT_STRAIGHT_R)
        tag_text_note.GetLeaders()[0].End = XYZ(rStart,0,-hPit-0.25+hStartTagcounter)
        
    ####
    ####END 
    if rEnd != rEndOld:
        ####
        # Define Points for Bar
        bar_bot_end = XYZ(rEnd, 0, -hPit-1+hcounter)
        if rEnd < rPlinth:
            bar_top_end = XYZ(rEnd, 0, 0)
        elif rEnd >= rPlinth:
            bar_top_end = XYZ(rEnd, 0, -1)
        # Create hidden line 
        endHidLine = doc.Create.NewDetailCurve(view, Line.CreateBound(bar_top_end, bar_bot_end))
        endHidLine.LineStyle = doc.GetElement(ElementId(1019189))
        # Create the text note
        text_note = TextNote.Create(doc, view.Id, XYZ(bar_top_end.X-0.53,0,-hPit),"END", text_note_type_id)
        text_note.TextNoteType = doc.GetElement(ElementId(1018389))
        # Create the rebar tag text note
        # Create the dimension
        refArray = ReferenceArray()
        refArray.Append(CL.GeometryCurve.Reference)
        refArray.Append(endHidLine.GeometryCurve.Reference)
        dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(0,0,-hPit-1+hcounter),XYZ(rEnd,0,-hPit-2-hcounter)), refArray, doc.GetElement(ElementId(1018370)))
        dim.TextPosition = dim.TextPosition + XYZ(0, 0, -0.04750)    
    if rEnd == rEndOld:
        hTagcounter = hTagcounter-120/304.8
    tag_text_note = TextNote.Create(doc, view.Id, XYZ(bar_top_end.X-2.25,0,-hPit+hTagcounter), str(count)+"x"+str(Size)+"-"+str(bar_mark)+"-"+str(angle), text_note_type_id)
    tag_text_note.TextNoteType = doc.GetElement(ElementId(1018389))
    tag_text_note.AddLeader(TextNoteLeaderTypes.TNLT_STRAIGHT_R)
    tag_text_note.GetLeaders()[0].End = XYZ(bar_top_end.X,0,-hPit+hTagcounter-45/304.8)
    hcounter -= 120/304.8
    rEndOld = rEnd
#endregion

#region############################################## ST ################################################


# get all inputs form excel in a dictionary
BarDict_ST = {}
for i in range(1,200):
    i += 1
    if "ST" in str(xl.Cells(i, 1).Value2).replace(" ","") :
        bar_mark = str(xl.Cells(i,1).Value2)
        r = float(xl.Cells(i,2).Value2)/304.8
        noBars = float(xl.Cells(i,3).Value2)
        Size = str(xl.Cells(i,4).Value2)
        # print(bar_mark)
        BarDict_ST[bar_mark] = [r, noBars, Size]
        print(BarDict_ST[bar_mark])


# Get Elements from document

FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
for elem in FEC:
    if elem.Name == '1PA_WTF_SteelTower':
        wtf_steel = elem

rPlinth = wtf_steel.LookupParameter('rPlinth').AsDouble()
rBase = wtf_steel.LookupParameter('rBase').AsDouble()
hPlinth = wtf_steel.LookupParameter('hPlinth').AsDouble()
hBase = wtf_steel.LookupParameter('hBase').AsDouble()
hCone  = wtf_steel.LookupParameter('hCone').AsDouble()
hPit = wtf_steel.LookupParameter('hBottomVoid').AsDouble()
rPitIn = wtf_steel.LookupParameter('rVoidInner').AsDouble()
rPitOut = wtf_steel.LookupParameter('rVoidOuter').AsDouble()
slabSlope = ((hCone-hBase)/(rBase - rPlinth))

def base_height_at_radius(r):
    # Calculate the height of the base at a given radius
    if r <= rPlinth:
        return hPlinth
    elif r <= rBase and r > rPlinth:
        return hCone - (r - rPlinth) * slabSlope
    

# Collect all FamilySymbols in the category
FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DetailComponents).OfClass(FamilySymbol).ToElements()
typeDict = {}
for sym in FEC:
    if sym.FamilyName == "BVI-DET-Stool" :
        if sym.LookupParameter("Type Name").AsString() == 'Y16':
            Y16 = sym
            typeDict["16"] = Y16
            print("Y16 loaded")
        if sym.LookupParameter("Type Name").AsString() == 'Y20':
            Y20 = sym
            typeDict["20"] = Y20
            print("Y20 loaded")
        if sym.LookupParameter("Type Name").AsString() == 'Y25':
            Y25 = sym
            typeDict["25"] = Y25
            print("Y25 loaded")
        if sym.LookupParameter("Type Name").AsString() == 'Y32':
            Y32 = sym
            typeDict["32"] = Y32
            print("Y32 loaded")
        # if sym.LookupParameter("Type Name").AsString() == 'Y20*':
        #     Y20m = sym
        #     typeDict["20*"] = Y20m
        #     print("Y20* loaded")
        # if sym.LookupParameter("Type Name").AsString() == 'Y25*':
        #     Y25m = sym
        #     typeDict["25*"] = Y25m
        #     print("Y25* loaded")
        # if sym.LookupParameter("Type Name").AsString() == 'Y32*':
        #     Y32m = sym
        #     typeDict["32*"] = Y32m
        #     print("Y32* loaded")
        #     break

# Get the default text note type
text_note_type_id = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType)

for symbol in [Y16, Y20, Y25, Y32]:
    if isinstance(symbol, FamilySymbol) and not symbol.IsActive:
        symbol.Activate()
        doc.Regenerate()

############################################### ST ###############################################

cover = 50*0.00328084
STdimLineList = []
for bar_mark, values in BarDict_ST.items():
    r = values[0]
    noBars = values[1]
    Size = str(values[2]).strip('Y').strip('R')
    SizeFloat = float(str(values[2]).strip('Y').strip('R'))
    barRad = SizeFloat*0.00328084
    # Place detail component
    placementPointL = XYZ(r-barRad, 0, cover)#+(SizeFloat*0.00328084)/2)
    FamInstanceL = doc.Create.NewFamilyInstance(placementPointL, typeDict[Size], view)
    FamInstanceL.LookupParameter('hStool').Set((base_height_at_radius(r)) -cover*2)
    placementPointR = XYZ(-r+barRad, 0, cover)#+(SizeFloat*0.00328084)/2)
    FamInstanceR = doc.Create.NewFamilyInstance(placementPointR, typeDict[Size], view)
    FamInstanceR.flipHand()
    FamInstanceR.LookupParameter('hStool').Set((base_height_at_radius(r)) -cover*2)

    # Create center line 
    cl = doc.Create.NewDetailCurve(view, Line.CreateBound(XYZ(0,0,hBase/2), XYZ(0,0,(hBase/2)-0.1)))
    cl.LineStyle = doc.GetElement(ElementId(1019189))

    # Create hidden line 
    endHidLine = doc.Create.NewDetailCurve(view, Line.CreateBound(XYZ(r,0,hBase/2), XYZ(r,0,(hBase/2)-0.1)))
    endHidLine.LineStyle = doc.GetElement(ElementId(1019189))

    spacing = 360/noBars
    tagText =   str(noBars) + "x" + str(values[2])+"-"+str(bar_mark)+"-"+str(spacing)+"°"
    textXposition = r+0.3 #(len(tagText)*35/304.8)/2)
    tag_text_note = TextNote.Create(doc, view.Id, XYZ(textXposition,0,base_height_at_radius(r)-2), tagText, text_note_type_id)
    tag_text_note.TextNoteType = doc.GetElement(ElementId(1018389))
    #add leader to text note
    tag_text_note.AddLeader(TextNoteLeaderTypes.TNLT_STRAIGHT_L)
    for i in range(0, len(tag_text_note.GetLeaders())):
        tag_text_note.GetLeaders()[i].End = XYZ(r,0,base_height_at_radius(r)-2.15)
    
    pt1 = XYZ(r,0,1)
    pt2 = XYZ(r,0,1.01)
    dimLine = doc.Create.NewDetailCurve(view, Line.CreateBound(pt1, pt2))
    STdimLineList.append(dimLine)

#endregion##############################################################################################################################################################################################################################################

#region############################################## SF ###############################################

# get all inputs form excel in a list
BarList_SF = []
for i in range(1,200):
    i += 1
    if "SF" in str(xl.Cells(i, 1).Value2).replace(" ","") :
        bar_mark = str(xl.Cells(i,1).Value2)
        Size = str(xl.Cells(i,4).Value2)
        
#dimension 
topLine_p1 = XYZ(-rBase - 1,0,hBase - cover)
topLine_p2 = XYZ(-rBase - 1.05,0,hBase - cover)
botLine_p1 = XYZ(-rBase - 1   ,0,cover)
botLine_p2 = XYZ(-rBase - 1.05,0,cover)

refArray = ReferenceArray()
topLine = doc.Create.NewDetailCurve(view, Line.CreateBound(topLine_p1, topLine_p2))
botLine = doc.Create.NewDetailCurve(view, Line.CreateBound(botLine_p1, botLine_p2))
refArray.Append(topLine.GeometryCurve.Reference)
refArray.Append(botLine.GeometryCurve.Reference)

dim = doc.Create.NewDimension(view, Line.CreateBound(topLine_p1, botLine_p2), refArray, doc.GetElement(ElementId(1018370)))
text = str(Size)+"-"+str(bar_mark)#+"-"+str(spacing*304.8)[:-2]
dim.Below = text
#endregion

#region############################################## GR ###############################################

def get_grid_bar_for_text_tag(list_of_rebars):
    grid_bars = [r for r in list_of_rebars if "GR" in r.LookupParameter("Schedule Mark").AsString()]
    if not grid_bars:
        return None
    grid_bars.sort(key=lambda r: r.LookupParameter("Schedule Mark").AsString())
    return grid_bars[-1]

FEC_GR = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()

grids = []
for elem in FEC_GR:
    if "GR" in elem.LookupParameter("Schedule Mark").AsString():
        grids.append(elem)

def grid_bar_tag(rebar, suff):
    bar_mark = rebar.LookupParameter("Schedule Mark").AsString()
    bar_size = rebar.LookupParameter("Bar Diameter").AsDouble()*304.8
    spacing = rebar.LookupParameter("Rebar Spacing").AsValueString()
    tag_text = "Y{}-{}(a-{})-{}".format(str(bar_size)[:-2], str(bar_mark)[:-1], suff, spacing)
    return tag_text


BarList_GR = []
GRdimLineList = []
for i in range(1,200):
    i += 1
    if "GR" in str(xl.Cells(i, 1).Value2).replace(" ","") :
        bar_mark = str(xl.Cells(i,1).Value2)
        Size = str(xl.Cells(i,4).Value2)
        height = float(xl.Cells(i,6).Value2)/304.8
        spacing = int(xl.Cells(i,7).Value2)
        suffixBar = get_grid_bar_for_text_tag(grids)


        tag_text = "{}-{}(a-{})-{}".format(Size, str(bar_mark), suffixBar.LookupParameter("Schedule Mark").AsString()[-1], spacing)
        text_note = TextNote.Create(doc, view.Id, XYZ(200/304.8,0,height-0.5), tag_text, text_note_type_id)
        text_note.HorizontalAlignment = HorizontalTextAlignment.Left
        text_note.VerticalAlignment = VerticalTextAlignment.Top

        leader = text_note.AddLeader(TextNoteLeaderTypes.TNLT_STRAIGHT_R)
        leader.End = XYZ(95/304.8,0,height)

        pt1 = XYZ(95/304.8,0,height)
        pt2 = XYZ(100/304.8,0,height)
        dimLine = doc.Create.NewDetailCurve(view, Line.CreateBound(pt1, pt2))
        GRdimLineList.append(dimLine)

#endregion

#region############################################## PF ###############################################

for i in range(1,200):
    if "PF" in str(xl.Cells(i, 1).Value2).replace(" ","") :
        bar_mark = str(xl.Cells(i,1).Value2)
        Size = str(xl.Cells(i,4).Value2)
        spacing = str(xl.Cells(i,7).Value2)
        
#dimension 
topLine_p1 = XYZ(-rPlinth - 1    ,0,hPlinth - cover)
topLine_p2 = XYZ(-rPlinth - 1.05 ,0,hPlinth - cover)
botLine_p1 = XYZ(-rPlinth - 1    ,0,hCone)
botLine_p2 = XYZ(-rPlinth - 1.05 ,0,hCone)

refArray = ReferenceArray()
topLine = doc.Create.NewDetailCurve(view, Line.CreateBound(topLine_p1, topLine_p2))
botLine = doc.Create.NewDetailCurve(view, Line.CreateBound(botLine_p1, botLine_p2))
refArray.Append(topLine.GeometryCurve.Reference)
refArray.Append(botLine.GeometryCurve.Reference)

dim = doc.Create.NewDimension(view, Line.CreateBound(topLine_p1, botLine_p2), refArray, doc.GetElement(ElementId(1018370)))
text = str(Size)+"-"+str(bar_mark)+"-"+str(spacing)[:-2]
dim.ValueOverride = text
#endregion

#region############################################## PH ###############################################

for i in range(1,200):
    if "PH" in str(xl.Cells(i, 1).Value2).replace(" ","") :
        bar_mark = str(xl.Cells(i,1).Value2)
        Size = str(xl.Cells(i,4).Value2)
        spacing = str(xl.Cells(i,7).Value2)

rightX = rTower +wGroutTop/2 + cover

#dimension 
LeftLine_p1 = XYZ(-rPlinth ,0,hPlinth +0.51)
LeftLine_p2 = XYZ(-rPlinth ,0,hPlinth +0.5)
RightLine_p1 = XYZ(-rightX ,0,hPlinth +0.51)
RightLine_p2 = XYZ(-rightX ,0,hPlinth +0.5)

refArray = ReferenceArray()
topLine = doc.Create.NewDetailCurve(view, Line.CreateBound(LeftLine_p1, LeftLine_p2))
botLine = doc.Create.NewDetailCurve(view, Line.CreateBound(RightLine_p1, RightLine_p2))
refArray.Append(topLine.GeometryCurve.Reference)
refArray.Append(botLine.GeometryCurve.Reference)

dim = doc.Create.NewDimension(view, Line.CreateBound(LeftLine_p1+XYZ(0,0,0.1), RightLine_p1+XYZ(0,0,0.1)), refArray, doc.GetElement(ElementId(1018370)))
text = str(Size)+"-"+str(bar_mark)#+"-"+str(spacing*304.8)[:-2]
dim.ValueOverride = text
#endregion

FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
HPtagList = []
PCtagList = []
RAtagList = []
GRdimList = []
for bar in FEC:
    if bar.LookupParameter("Mark").AsString() == "SECTION HAIR PINS":
        HPtagList.append(bar)
    elif "PC" in bar.LookupParameter("Schedule Mark").AsString():
        PCtagList.append(bar)
    elif "RA" in bar.LookupParameter("Schedule Mark").AsString():
        RAtagList.append(bar)
    elif "GR" in bar.LookupParameter("Schedule Mark").AsString():
        GRdimList.append(bar)

#region############################################## HP ###############################################

for bar in HPtagList:
    if get_rebar_midpoint(bar).X < 0:
        HPtagList.remove(bar)
    
for bar in HPtagList:
    tag = IndependentTag.Create(
        doc,
        ElementId(3488323),
        view.Id,
        get_rebar_tag_reference(bar),
        True,
        TagOrientation.Horizontal,
        XYZ(rPlinth+1.5, 0, get_rebar_midpoint(bar).Z))

    tag.TagHeadPosition = XYZ(rPlinth+1.5, 0, get_rebar_midpoint(bar).Z)
    tag.LeaderEndCondition = LeaderEndCondition.Free

    tag_ref = tag.GetTaggedReferences()[0]
    tag.SetLeaderEnd( tag_ref, XYZ(rPlinth-cover-(25/304.8), 0, get_rebar_midpoint(bar).Z) )

#endregion

#region############################################## PC ###############################################

Tag_pc_bars_in_view(PCtagList, view)

#endregion

#region############################################## RA ###############################################

Tag_ra_bars_in_view(RAtagList, view)

#endregion


#Grind dimension lines for GR bars
refArray = ReferenceArray()

wtf_face_ref = get_face_reference_by_geometry_id(wtf_steel, 104)
if wtf_face_ref is not None:
    refArray.Append(wtf_face_ref)

for line in GRdimLineList:
    refArray.Append(line.GeometryCurve.Reference)

dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(-0.5,0,hPlinth),XYZ(-0.5,0,0)), refArray, doc.GetElement(ElementId(1018370)))

#Stool dimension lines for ST bars
refArray = ReferenceArray()
cl_ref = CL.GeometryCurve.Reference
refArray.Append(cl_ref)
pt1 = XYZ(rBase,0,1)
pt2 = XYZ(rBase,0,1.01)
line = doc.Create.NewDetailCurve(view, Line.CreateBound(pt1, pt2))
refArray.Append(line.GeometryCurve.Reference)
for line in STdimLineList:
    refArray.Append(line.GeometryCurve.Reference)

dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(0,0,1),XYZ(rBase,0,1)), refArray, doc.GetElement(ElementId(1018370)))




t.Commit()
#close excel
workbook.Close(False)
excel.Quit()

