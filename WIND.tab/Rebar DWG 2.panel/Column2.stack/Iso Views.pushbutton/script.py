# -*- coding: utf-8 -*-
import clr
import math
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
import Autodesk.Revit.DB.Structure as Structure
from collections import OrderedDict
from pyrevit import forms
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel


uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

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

    positive_y = [pt for pt in intersections if pt.Y > 0]
    if positive_y:
        return max(positive_y, key=lambda p: p.Y)

    return intersections[0]

def get_safe_tag_head_position(rebar_element, plane_normal, plane_origin=XYZ(0, 0, 0)):
    intersection_pt = get_rebar_plane_intersection(rebar_element, plane_normal, plane_origin)
    if intersection_pt is not None:
        return intersection_pt
    return get_rebar_midpoint(rebar_element)

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

def get_rebar_tag_reference(rebar_element):
    if rebar_element is None or not rebar_element.IsValidObject:
        raise Exception("Invalid rebar element passed to get_rebar_tag_reference")

    subelements = rebar_element.GetSubelements()
    if subelements and len(subelements) > 0:
        return subelements[0].GetReference()

    return Reference(rebar_element)

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

def grid_bar_tag(rebar, suff):
    bar_mark = rebar.LookupParameter("Schedule Mark").AsString()
    bar_size = rebar.LookupParameter("Bar Diameter").AsDouble()*304.8
    spacing = rebar.LookupParameter("Rebar Spacing").AsValueString()
    tag_text = "Y{}-{}(a-{})-{}".format(str(bar_size)[:-2], str(bar_mark)[:-1], suff, spacing)
    return tag_text

def get_grid_bar_for_text_tag(list_of_rebars):
    grid_bars = [r for r in list_of_rebars if "GR" in r.LookupParameter("Schedule Mark").AsString()]
    if not grid_bars:
        return None
    grid_bars.sort(key=lambda r: r.LookupParameter("Schedule Mark").AsString())
    return grid_bars[-1]


def Tag_pc_bars_in_view(tag_bars, view):
    PCtagBars = [bar for bar in tag_bars if "PC" in bar.LookupParameter("Schedule Mark").AsString()]
    PCbarDict = {}
    for bar in PCtagBars:
        bar_mark = bar.LookupParameter("Schedule Mark").AsString()
        if bar_mark not in PCbarDict:
            PCbarDict[bar_mark] = []
        PCbarDict[bar_mark].append(bar)

    # Remove duplicates by Schedule Mark, keeping the one with highest Y at x=0 intersection.
    selected_pc_bars = {}
    for bar_mark, bars in PCbarDict.items():
        max_y = float('-inf')
        bar_to_keep = None
        for bar in bars:
            intersection_pt = get_rebar_plane_intersection(bar, XYZ(1, 0, 0), XYZ(0, 0, 0))
            if intersection_pt is not None and intersection_pt.Y > max_y:
                max_y = intersection_pt.Y
                bar_to_keep = bar
        if bar_to_keep is None and len(bars) > 0:
            bar_to_keep = bars[0]

        selected_pc_bars[bar_mark] = bar_to_keep
        for bar in bars:
            if bar != bar_to_keep and bar in tag_bars:
                tag_bars.remove(bar)

    # Remove grid bars from tag_bars if both endpoints are on x=0 or y=0.
    for bar in list(tag_bars):
        bar_mark = bar.LookupParameter("Schedule Mark").AsString()
        if "GR" in bar_mark:
            curves = list(bar.GetCenterlineCurves(
                False, False, False,
                Structure.MultiplanarOption.IncludeOnlyPlanarCurves,
                0
            ))
            if curves:
                curve = curves[0]
                pts = list(curve.Tessellate())
                if len(pts) >= 2:
                    p1 = pts[0]
                    p2 = pts[-1]
                    if (abs(p1.X) < 1e-9 and abs(p2.X) < 1e-9) or (abs(p1.Y) < 1e-9 and abs(p2.Y) < 1e-9):
                        tag_bars.remove(bar)

    # PC bars are tagged separately, then skipped in the general tagging loop.
    for key in selected_pc_bars.keys():
        bar = selected_pc_bars[key]
        if bar.LookupParameter("Schedule Mark").AsString()[-2:] == "00" or bar.LookupParameter("Schedule Mark").AsString()[-2:] == "01":
            tagType = ElementId(3488323)
            leaderLen = 3
        else:
            tagType = ElementId(3488323) #9338727)
            leaderLen = 3                # 100 / 304.8
        tag = IndependentTag.Create(
                doc,
                tagType,
                view.Id,
                get_rebar_tag_reference(bar),
                    True,
                TagOrientation.Horizontal,
                get_rebar_midpoint(bar)
        )
        tag.TagHeadPosition = get_safe_tag_head_position(bar, XYZ(1, 0, 0), XYZ(-leaderLen, 0, 0))
        tag.LeaderEndCondition = LeaderEndCondition.Free

        tag_ref = tag.GetTaggedReferences()[0]
        tag.SetLeaderEnd(tag_ref, get_safe_tag_head_position(bar, XYZ(1, 0, 0), XYZ(-32 / 304.8, 0, 0)))

    return tag_bars

FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
for element in FEC:
    if element.Name == "1PA_WTF_SteelTower":
        WTF = element
        break

rPlinth = WTF.LookupParameter("rPlinth").AsDouble()
hPlinth = WTF.LookupParameter("hPlinth").AsDouble()

#Isometric Views



PLINTH_REINFORCEMENT_ISOMETRIC_VIEW = doc.GetElement(ElementId(966797))



text_note_type_id = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType)

t= Transaction(doc)
t.Start("Annotate Rebar DWG2 iso views")

#region################TOP_BURSTING_REBAR_1_ISOMETRIC_VIEW##########################################################################################################################################################
TOP_BURSTING_REBAR_1_ISOMETRIC_VIEW  = doc.GetElement(ElementId(3095802))
view = TOP_BURSTING_REBAR_1_ISOMETRIC_VIEW
# collect all rebar visible in TOP_BURSTING_REBAR_1_ISOMETRIC_VIEW in a list
collector = FilteredElementCollector(doc, TOP_BURSTING_REBAR_1_ISOMETRIC_VIEW.Id).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
tag_bars = list(collector)  
#keep all bar with "GR" in their Bar Mark and only one bar for each unique Bar Mark from bars not containing "GR" in their Bar Mark
final_rebars = []
for rebar in tag_bars:
    bar_mark = rebar.LookupParameter("Schedule Mark").AsString()
    if "GR" in bar_mark:
        final_rebars.append(rebar)
    elif "PC" in bar_mark:
        final_rebars.append(rebar)
    else:
        if not any(r.LookupParameter("Schedule Mark").AsString() == bar_mark for r in final_rebars):
            final_rebars.append(rebar)

tag_bars = final_rebars

tag_bars = Tag_pc_bars_in_view(tag_bars, view)

#tagging all other bars
text_note_created = False
for bar in tag_bars:
    bar_mark = bar.LookupParameter("Schedule Mark").AsString()

    if "GR" in bar_mark:
        tag = IndependentTag.Create(
                doc,
                ElementId(9387342),
                view.Id,
                get_rebar_tag_reference(bar),
                True,
                TagOrientation.Horizontal,
                get_rebar_midpoint(bar)
        )
        tag.TagHeadPosition = get_safe_tag_head_position(bar, XYZ(0, 1, 0), XYZ(0, 0, 0))
        print("tag created for bar mark: {}".format(bar_mark))
    elif "PC" not in bar_mark:
        tag = IndependentTag.Create(
                doc,
                ElementId(3488323),
                view.Id,
                get_rebar_tag_reference(bar),
                True,
                TagOrientation.Horizontal,
                get_rebar_midpoint(bar)
        )
        print("tag created for bar mark: {}".format(bar_mark))
        if "PH" in bar_mark:
            tag.TagHeadPosition = get_safe_tag_head_position(bar, XYZ(1, 0, 0), XYZ(-2, 0, 0))
            tag.LeaderEndCondition = LeaderEndCondition.Free

            tag_ref = tag.GetTaggedReferences()[0]
            tag.SetLeaderEnd( tag_ref, get_safe_tag_head_position(bar, XYZ(1, 0, 0), XYZ(0, 0, 0)))
        
        if "PF" in bar_mark:
            tag.TagHeadPosition = get_safe_tag_head_position(bar, XYZ(1, 0, 0), XYZ(-2, 0, 0))
            tag.LeaderEndCondition = LeaderEndCondition.Free

            tag_ref = tag.GetTaggedReferences()[0]
            tag.SetLeaderEnd( tag_ref, get_safe_tag_head_position(bar, XYZ(1, 0, 0), XYZ(0, 0, 0)) )

        if "HP" in bar_mark:
            tag.TagHeadPosition =  XYZ(rPlinth+2, 0, get_rebar_midpoint(bar).Z-2)

        if "PV" in bar_mark:
            SecBoxMinZ = view.GetSectionBox().Min.Z
            tag.TagHeadPosition =  get_safe_tag_head_position(bar, XYZ(0, 0, 1), XYZ(0, 0, SecBoxMinZ)) + XYZ(0,-2, 0)
            tag.LeaderEndCondition = LeaderEndCondition.Free

            tag_ref = tag.GetTaggedReferences()[0]
            tag.SetLeaderEnd( tag_ref, get_safe_tag_head_position(bar, XYZ(0, 0, 1), XYZ(0, 0, SecBoxMinZ)) )

        if not text_note_created:
            grid_bar = get_grid_bar_for_text_tag(tag_bars)
            if grid_bar is not None:
                tag_text = grid_bar_tag(grid_bar,grid_bar.LookupParameter("Schedule Mark").AsString()[-1])
                text_note = TextNote.Create(doc, view.Id,get_rebar_midpoint(grid_bar)+ XYZ(-2,0,0), tag_text, text_note_type_id)
                text_note.HorizontalAlignment = HorizontalTextAlignment.Left
                text_note.VerticalAlignment = VerticalTextAlignment.Top

                leader = text_note.AddLeader(TextNoteLeaderTypes.TNLT_STRAIGHT_R)
                leader.End = get_rebar_midpoint(grid_bar)
                text_note_created = True


print("Finished annotating TOP_BURSTING_REBAR_1_ISOMETRIC_VIEW")
#endregion


#region################TOP_BURSTING_REBAR_2_ISOMETRIC_VIEW##########################################################################################################################################################
TOP_BURSTING_REBAR_2_ISOMETRIC_VIEW  = doc.GetElement(ElementId(3096005))
view = TOP_BURSTING_REBAR_2_ISOMETRIC_VIEW
# collect all rebar visible in TOP_BURSTING_REBAR_2_ISOMETRIC_VIEW in a list
collector = FilteredElementCollector(doc, TOP_BURSTING_REBAR_2_ISOMETRIC_VIEW.Id).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
tag_bars = list(collector)  
#keep all bar with "GR" in their Bar Mark and only one bar for each unique Bar Mark from bars not containing "GR" in their Bar Mark
final_rebars = []
for rebar in tag_bars:
    bar_mark = rebar.LookupParameter("Schedule Mark").AsString()
    if "GR" in bar_mark:
        final_rebars.append(rebar)
    elif "PC" in bar_mark:
        final_rebars.append(rebar)
    else:
        if not any(r.LookupParameter("Schedule Mark").AsString() == bar_mark for r in final_rebars):
            final_rebars.append(rebar)

tag_bars = final_rebars

#tag PC bars    
tag_bars = Tag_pc_bars_in_view(tag_bars, view)

# tag top radials
# collect TR bars by Schedule Mark, keeping the one with the maximum start-point X.
print("starting to tag TR bars in TOP_BURSTING_REBAR_2_ISOMETRIC_VIEW")
TR_tagBars = {}
TrbarMarks = []
for bar in tag_bars:
    bar_mark = bar.LookupParameter("Schedule Mark").AsString()
    if "TR" in bar_mark and bar_mark not in TrbarMarks:
        TrbarMarks.append(bar_mark)
        TR_tagBars[bar_mark] = bar
for mark in TrbarMarks:
    max_x = 0
    for bar in tag_bars:
        bar_mark = bar.LookupParameter("Schedule Mark").AsString()
        if bar_mark == mark:
            start_pt, end_pt = get_rebar_endpoints(bar)
            if start_pt is not None and start_pt.X > max_x:
                max_x = start_pt.X
                TR_tagBars[mark] = bar

#remove all dupicate TR bars with the same bar mark except the one with maximum start-point X in TR_tagbars
filtered_tag_bars = []
for bar in tag_bars:
    bar_mark = bar.LookupParameter("Schedule Mark").AsString()
    if "TR" not in bar_mark or TR_tagBars.get(bar_mark) == bar:
        filtered_tag_bars.append(bar)
tag_bars = filtered_tag_bars

for key in TR_tagBars.keys():
    bar = TR_tagBars[key]
    tag = IndependentTag.Create(
        doc,
        ElementId(3488323),
        view.Id,
        get_rebar_tag_reference(bar),
        True,
        TagOrientation.Horizontal,
        get_rebar_midpoint(bar)
    )
    print("tag created for bar mark: {}".format(key))
    tag.TagHeadPosition =  get_rebar_endpoints(bar)[1] + XYZ(2, -2, 0)
    tag.LeaderEndCondition = LeaderEndCondition.Free
    tag_ref = tag.GetTaggedReferences()[0]
    tag.SetLeaderEnd( tag_ref, get_rebar_endpoints(bar)[1] )
print("finished tagging TR bars in TOP_BURSTING_REBAR_2_ISOMETRIC_VIEW")

#tagging all other bars
text_note_created = False
for bar in tag_bars:
    bar_mark = bar.LookupParameter("Schedule Mark").AsString()

    if "GR" in bar_mark:
        tag = IndependentTag.Create(
                doc,
                ElementId(9387342),
                view.Id,
                get_rebar_tag_reference(bar),
                True,
                TagOrientation.Horizontal,
                get_rebar_midpoint(bar)
        )
        tag.TagHeadPosition = get_safe_tag_head_position(bar, XYZ(0, 1, 0), XYZ(0, 0, 0))
        print("tag created for bar mark: {}".format(bar_mark))
    elif "PC" not in bar_mark and "TR" not in bar_mark:
        tag = IndependentTag.Create(
                doc,
                ElementId(3488323),
                view.Id,
                get_rebar_tag_reference(bar),
                True,
                TagOrientation.Horizontal,
                get_rebar_midpoint(bar)
        )
        print("tag created for bar mark: {}".format(bar_mark))
        if "PH" in bar_mark:
            tag.TagHeadPosition = get_safe_tag_head_position(bar, XYZ(1, 0, 0), XYZ(-2, 0, 0))
            tag.LeaderEndCondition = LeaderEndCondition.Free

            tag_ref = tag.GetTaggedReferences()[0]
            tag.SetLeaderEnd( tag_ref, get_safe_tag_head_position(bar, XYZ(1, 0, 0), XYZ(0, 0, 0)))
        
        if "PF" in bar_mark:
            tag.TagHeadPosition = get_safe_tag_head_position(bar, XYZ(1, 0, 0), XYZ(-2, 0, 0))
            tag.LeaderEndCondition = LeaderEndCondition.Free

            tag_ref = tag.GetTaggedReferences()[0]
            tag.SetLeaderEnd( tag_ref, get_safe_tag_head_position(bar, XYZ(1, 0, 0), XYZ(0, 0, 0)) )

        if "HP" in bar_mark:
            tag.TagHeadPosition =  XYZ(rPlinth+2, 0, get_rebar_midpoint(bar).Z-2)

        if "PV" in bar_mark:
            SecBoxMinZ = view.GetSectionBox().Min.Z
            tag.TagHeadPosition =  get_safe_tag_head_position(bar, XYZ(0, 0, 1), XYZ(0, 0, SecBoxMinZ)) + XYZ(0,-2, 0)
            tag.LeaderEndCondition = LeaderEndCondition.Free

            tag_ref = tag.GetTaggedReferences()[0]
            tag.SetLeaderEnd( tag_ref, get_safe_tag_head_position(bar, XYZ(0, 0, 1), XYZ(0, 0, SecBoxMinZ)) )

        if not text_note_created:
            grid_bar = get_grid_bar_for_text_tag(tag_bars)
            if grid_bar is not None:
                tag_text = grid_bar_tag(grid_bar,grid_bar.LookupParameter("Schedule Mark").AsString()[-1])
                text_note = TextNote.Create(doc, view.Id,get_rebar_midpoint(grid_bar)+ XYZ(-2,0,0), tag_text, text_note_type_id)
                text_note.HorizontalAlignment = HorizontalTextAlignment.Left
                text_note.VerticalAlignment = VerticalTextAlignment.Top

                leader = text_note.AddLeader(TextNoteLeaderTypes.TNLT_STRAIGHT_R)
                leader.End = get_rebar_midpoint(grid_bar)
                text_note_created = True


print("Finished annotating TOP_BURSTING_REBAR_2_ISOMETRIC_VIEW")
#endregion


#region################SURFACE_REINFORCEMENT_3D_VIEW##########################################################################################################################################################
SURFACE_REINFORCEMENT_3D_VIEW  = doc.GetElement(ElementId(603339))
view = SURFACE_REINFORCEMENT_3D_VIEW
# collect all rebar visible in SURFACE_REINFORCEMENT_3D_VIEW in a list
collector = FilteredElementCollector(doc, SURFACE_REINFORCEMENT_3D_VIEW.Id).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
tag_bars = list(collector)  
#keep all bar with "GR" in their Bar Mark and only one bar for each unique Bar Mark from bars not containing "GR" in their Bar Mark
final_rebars = []
for rebar in tag_bars:
    bar_mark = rebar.LookupParameter("Schedule Mark").AsString()
    if "GR" in bar_mark:
        final_rebars.append(rebar)
    elif "PC" in bar_mark:
        final_rebars.append(rebar)
    else:
        if not any(r.LookupParameter("Schedule Mark").AsString() == bar_mark for r in final_rebars):
            final_rebars.append(rebar)

tag_bars = final_rebars

tag_bars = Tag_pc_bars_in_view(tag_bars, view)

# tag top radials
# collect BR bars by Schedule Mark, keeping the one with the maximum start-point X.
print("starting to tag BR bars in TOP_BURSTING_REBAR_2_ISOMETRIC_VIEW")
BR_tagBars = {}
BRbarMarks = []
for bar in tag_bars:
    bar_mark = bar.LookupParameter("Schedule Mark").AsString()
    if "BR" in bar_mark and bar_mark not in BRbarMarks:
        BRbarMarks.append(bar_mark)
        BR_tagBars[bar_mark] = bar
for mark in BRbarMarks:
    max_x = 0
    for bar in tag_bars:
        bar_mark = bar.LookupParameter("Schedule Mark").AsString()
        if bar_mark == mark:
            start_pt, end_pt = get_rebar_endpoints(bar)
            if start_pt is not None and start_pt.X > max_x:
                max_x = start_pt.X
                BR_tagBars[mark] = bar

# remove all duplicate BR bars with the same bar mark, keeping only the one in BR_tagBars
filtered_tag_bars = []
for bar in tag_bars:
    bar_mark = bar.LookupParameter("Schedule Mark").AsString()
    if "BR" not in bar_mark or BR_tagBars.get(bar_mark) == bar:
        filtered_tag_bars.append(bar)
tag_bars = filtered_tag_bars

for key in BR_tagBars.keys():
    bar = BR_tagBars[key]
    tag = IndependentTag.Create(
        doc,
        ElementId(3488323),
        view.Id,
        get_rebar_tag_reference(bar),
        True,
        TagOrientation.Horizontal,
        get_rebar_midpoint(bar)
    )
    print("tag created for bar mark: {}".format(key))
    tag.TagHeadPosition =  get_rebar_endpoints(bar)[0] + XYZ(2, -2, 0)
    tag.LeaderEndCondition = LeaderEndCondition.Free
    tag_ref = tag.GetTaggedReferences()[0]
    tag.SetLeaderEnd( tag_ref, get_rebar_endpoints(bar)[0] )
print("finished tagging BR bars in TOP_BURSTING_REBAR_2_ISOMETRIC_VIEW")

#tagging all other bars
text_note_created = False
for bar in tag_bars:
    bar_mark = bar.LookupParameter("Schedule Mark").AsString()

    if "GR" in bar_mark:
        tag = IndependentTag.Create(
                doc,
                ElementId(9387342),
                view.Id,
                get_rebar_tag_reference(bar),
                True,
                TagOrientation.Horizontal,
                get_rebar_midpoint(bar)
        )
        tag.TagHeadPosition = get_safe_tag_head_position(bar, XYZ(0, 1, 0), XYZ(0, 0, 0))
        print("tag created for bar mark: {}".format(bar_mark))
    elif "PC" not in bar_mark and "BR" not in bar_mark:
        tag = IndependentTag.Create(
                doc,
                ElementId(3488323),
                view.Id,
                get_rebar_tag_reference(bar),
                True,
                TagOrientation.Horizontal,
                get_rebar_midpoint(bar)
        )
        print("tag created for bar mark: {}".format(bar_mark))
        if "PH" in bar_mark:
            tag.TagHeadPosition = get_safe_tag_head_position(bar, XYZ(1, 0, 0), XYZ(-2, 0, 0))
            tag.LeaderEndCondition = LeaderEndCondition.Free

            tag_ref = tag.GetTaggedReferences()[0]
            tag.SetLeaderEnd( tag_ref, get_safe_tag_head_position(bar, XYZ(1, 0, 0), XYZ(0, 0, 0)))
        
        if "PF" in bar_mark:
            tag.TagHeadPosition = get_safe_tag_head_position(bar, XYZ(1, 0, 0), XYZ(-2, 0, 0))
            tag.LeaderEndCondition = LeaderEndCondition.Free

            tag_ref = tag.GetTaggedReferences()[0]
            tag.SetLeaderEnd( tag_ref, get_safe_tag_head_position(bar, XYZ(1, 0, 0), XYZ(0, 0, 0)) )

        if "HP" in bar_mark:
            tag.TagHeadPosition =  XYZ(rPlinth+2, 0, get_rebar_midpoint(bar).Z-2)

        if "PV" in bar_mark:
            SecBoxMaxZ = view.GetSectionBox().Max.Z
            tag.TagHeadPosition =  get_safe_tag_head_position(bar, XYZ(0, 0, 1), XYZ(0, 0, SecBoxMaxZ)) + XYZ(3,-2, 0)
            tag.LeaderEndCondition = LeaderEndCondition.Free

            tag_ref = tag.GetTaggedReferences()[0]
            tag.SetLeaderEnd( tag_ref, get_safe_tag_head_position(bar, XYZ(0, 0, 1), XYZ(0, 0, SecBoxMaxZ)) )

        if not text_note_created:
            grid_bar = get_grid_bar_for_text_tag(tag_bars)
            if grid_bar is not None:
                tag_text = grid_bar_tag(grid_bar,grid_bar.LookupParameter("Schedule Mark").AsString()[-1])
                text_note = TextNote.Create(doc, view.Id,get_rebar_midpoint(grid_bar)+ XYZ(-2,0,0), tag_text, text_note_type_id)
                text_note.HorizontalAlignment = HorizontalTextAlignment.Left
                text_note.VerticalAlignment = VerticalTextAlignment.Top

                leader = text_note.AddLeader(TextNoteLeaderTypes.TNLT_STRAIGHT_R)
                leader.End = get_rebar_midpoint(grid_bar)
                text_note_created = True


print("Finished annotating SURFACE_REINFORCEMENT_3D_VIEW")
#endregion


#region################BOTTOM_BURSTING_REBAR_ISOMETRIC_VIEW##########################################################################################################################################################
BOTTOM_BURSTING_REBAR_ISOMETRIC_VIEW = doc.GetElement(ElementId(603329))
view = BOTTOM_BURSTING_REBAR_ISOMETRIC_VIEW
# collect all rebar visible in BOTTOM_BURSTING_REBAR_ISOMETRIC_VIEW in a list
collector = FilteredElementCollector(doc, BOTTOM_BURSTING_REBAR_ISOMETRIC_VIEW.Id).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
tag_bars = list(collector)  
#keep all bar with "GR" in their Bar Mark and only one bar for each unique Bar Mark from bars not containing "GR" in their Bar Mark
final_rebars = []
for rebar in tag_bars:
    bar_mark = rebar.LookupParameter("Schedule Mark").AsString()
    if "GR" in bar_mark:
        final_rebars.append(rebar)
    elif "PC" in bar_mark:
        final_rebars.append(rebar)
    else:
        if not any(r.LookupParameter("Schedule Mark").AsString() == bar_mark for r in final_rebars):
            final_rebars.append(rebar)

tag_bars = final_rebars

tag_bars = Tag_pc_bars_in_view(tag_bars, view)

# tag top radials
# collect BR bars by Schedule Mark, keeping the one with the maximum start-point X.
print("starting to tag BR bars in BOTTOM_BURSTING_REBAR_ISOMETRIC_VIEW")
BR_tagBars = {}
BRbarMarks = []
for bar in tag_bars:
    bar_mark = bar.LookupParameter("Schedule Mark").AsString()
    if "BR" in bar_mark and bar_mark not in BRbarMarks:
        BRbarMarks.append(bar_mark)
        BR_tagBars[bar_mark] = bar
for mark in BRbarMarks:
    max_x = 0
    for bar in tag_bars:
        bar_mark = bar.LookupParameter("Schedule Mark").AsString()
        if bar_mark == mark:
            start_pt, end_pt = get_rebar_endpoints(bar)
            if start_pt is not None and start_pt.X > max_x:
                max_x = start_pt.X
                BR_tagBars[mark] = bar

# remove all duplicate BR bars with the same bar mark, keeping only the one in BR_tagBars
filtered_tag_bars = []
for bar in tag_bars:
    bar_mark = bar.LookupParameter("Schedule Mark").AsString()
    if "BR" not in bar_mark or BR_tagBars.get(bar_mark) == bar:
        filtered_tag_bars.append(bar)
tag_bars = filtered_tag_bars

for key in BR_tagBars.keys():
    bar = BR_tagBars[key]
    tag = IndependentTag.Create(
        doc,
        ElementId(3488323),
        view.Id,
        get_rebar_tag_reference(bar),
        True,
        TagOrientation.Horizontal,
        get_rebar_midpoint(bar)
    )
    print("tag created for bar mark: {}".format(key))
    tag.TagHeadPosition =  get_rebar_endpoints(bar)[0] + XYZ(2, -2, 0)
    tag.LeaderEndCondition = LeaderEndCondition.Free
    tag_ref = tag.GetTaggedReferences()[0]
    tag.SetLeaderEnd( tag_ref, get_rebar_endpoints(bar)[0] )
print("finished tagging BR bars in BOTTOM_BURSTING_REBAR_ISOMETRIC_VIEW")

#tagging all other bars
text_note_created = False
for bar in tag_bars:
    bar_mark = bar.LookupParameter("Schedule Mark").AsString()

    if "GR" in bar_mark:
        tag = IndependentTag.Create(
                doc,
                ElementId(9387342),
                view.Id,
                get_rebar_tag_reference(bar),
                True,
                TagOrientation.Horizontal,
                get_rebar_midpoint(bar)
        )
        tag.TagHeadPosition = get_safe_tag_head_position(bar, XYZ(0, 1, 0), XYZ(0, 0, 0))
        print("tag created for bar mark: {}".format(bar_mark))
    elif "PC" not in bar_mark and "BR" not in bar_mark:
        tag = IndependentTag.Create(
                doc,
                ElementId(3488323),
                view.Id,
                get_rebar_tag_reference(bar),
                True,
                TagOrientation.Horizontal,
                get_rebar_midpoint(bar)
        )
        print("tag created for bar mark: {}".format(bar_mark))
        if "PH" in bar_mark:
            tag.TagHeadPosition = get_safe_tag_head_position(bar, XYZ(1, 0, 0), XYZ(-2, 0, 0))
            tag.LeaderEndCondition = LeaderEndCondition.Free

            tag_ref = tag.GetTaggedReferences()[0]
            tag.SetLeaderEnd( tag_ref, get_safe_tag_head_position(bar, XYZ(1, 0, 0), XYZ(0, 0, 0)))
        
        if "PF" in bar_mark:
            tag.TagHeadPosition = get_safe_tag_head_position(bar, XYZ(1, 0, 0), XYZ(-2, 0, 0))
            tag.LeaderEndCondition = LeaderEndCondition.Free

            tag_ref = tag.GetTaggedReferences()[0]
            tag.SetLeaderEnd( tag_ref, get_safe_tag_head_position(bar, XYZ(1, 0, 0), XYZ(0, 0, 0)) )

        if "HP" in bar_mark:
            tag.TagHeadPosition =  XYZ(rPlinth+2, 0, get_rebar_midpoint(bar).Z-2)

        if "PV" in bar_mark:
            SecBoxMaxZ = view.GetSectionBox().Max.Z
            tag.TagHeadPosition =  get_safe_tag_head_position(bar, XYZ(0, 0, 1), XYZ(0, 0, SecBoxMaxZ)) + XYZ(3,-2, 0)
            tag.LeaderEndCondition = LeaderEndCondition.Free

            tag_ref = tag.GetTaggedReferences()[0]
            tag.SetLeaderEnd( tag_ref, get_safe_tag_head_position(bar, XYZ(0, 0, 1), XYZ(0, 0, SecBoxMaxZ)) )

        if not text_note_created:
            grid_bar = get_grid_bar_for_text_tag(tag_bars)
            if grid_bar is not None:
                tag_text = grid_bar_tag(grid_bar,grid_bar.LookupParameter("Schedule Mark").AsString()[-1])
                text_note = TextNote.Create(doc, view.Id,get_rebar_midpoint(grid_bar)+ XYZ(-2,0,0), tag_text, text_note_type_id)
                text_note.HorizontalAlignment = HorizontalTextAlignment.Left
                text_note.VerticalAlignment = VerticalTextAlignment.Top

                leader = text_note.AddLeader(TextNoteLeaderTypes.TNLT_STRAIGHT_R)
                leader.End = get_rebar_midpoint(grid_bar)
                text_note_created = True


print("Finished annotating BOTTOM_BURSTING_REBAR_ISOMETRIC_VIEW")
#endregion


t.Commit()

