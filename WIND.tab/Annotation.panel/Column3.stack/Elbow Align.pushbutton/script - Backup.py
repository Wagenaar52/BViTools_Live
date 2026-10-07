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

# 100 mm target length in model space (Revit internal units are feet).
leaderLength = 100.0 / 304.8

# Horizontal in the active view plane is the view's right direction.
right_dir = view.RightDirection
if right_dir.IsZeroLength():
    right_dir = XYZ.BasisX
right_dir = right_dir.Normalize()
up_dir = view.UpDirection
if up_dir.IsZeroLength():
    up_dir = XYZ.BasisY
up_dir = up_dir.Normalize()
view_normal = view.ViewDirection
if view_normal.IsZeroLength():
    view_normal = XYZ.BasisZ
view_normal = view_normal.Normalize()

t = Transaction(doc)
t.Start("Set leader end")

#draw model line


def bbox_corners(bb):
    mn = bb.Min
    mx = bb.Max
    return [
        XYZ(mn.X, mn.Y, mn.Z), XYZ(mn.X, mn.Y, mx.Z), XYZ(mn.X, mx.Y, mn.Z), XYZ(mn.X, mx.Y, mx.Z),
        XYZ(mx.X, mn.Y, mn.Z), XYZ(mx.X, mn.Y, mx.Z), XYZ(mx.X, mx.Y, mn.Z), XYZ(mx.X, mx.Y, mx.Z)
    ]

def left_right_middle_from_bbox(element, active_view, right_axis, up_axis, normal_axis):
    bb = element.get_BoundingBox(active_view)

    corners = bbox_corners(bb)
    r_vals = [p.DotProduct(right_axis) for p in corners]
    u_vals = [p.DotProduct(up_axis) for p in corners]

    left_r = min(r_vals)
    right_r = max(r_vals)
    mid_u = 0.5 * (min(u_vals) + max(u_vals))


    left_mid = right_axis.Multiply(left_r) + up_axis.Multiply(mid_u) 
    right_mid = right_axis.Multiply(right_r) + up_axis.Multiply(mid_u) 
    return left_mid, right_mid


def get_tag_text_bounds_without_leader(tag, tag_ref, active_view, right_axis, up_axis, normal_axis):
    original_leader_end = None
    had_leader = hasattr(tag, "HasLeader") and tag.HasLeader

    if hasattr(tag, "GetLeaderEnd"):
        try:
            original_leader_end = tag.GetLeaderEnd(tag_ref)
        except Exception:
            original_leader_end = None

    if had_leader:
        tag.HasLeader = False
        doc.Regenerate()

    left_mid, right_mid = left_right_middle_from_bbox(tag, active_view, right_axis, up_axis, normal_axis)

    if had_leader:
        tag.HasLeader = True
        doc.Regenerate()
        if original_leader_end is not None and hasattr(tag, "SetLeaderEnd"):
            tag.SetLeaderEnd(tag_ref, original_leader_end)
            doc.Regenerate()

    return left_mid, right_mid


for element in selected_elements:
    # TextNote
    if hasattr(element, "GetLeaders"):
        leaders = element.GetLeaders()
        if leaders is None or len(leaders) == 0:
            continue

        leader = leaders[0]
        elbow = leader.Elbow
        leaderEnd = leader.End
        Anchor = leader.Anchor
        horVector = element.BaseDirection

        element.VerticalAlignment = VerticalTextAlignment.Top
        element.HorizontalAlignment = HorizontalTextAlignment.Left

        element.LeaderLeftAttachment = element.LeaderLeftAttachment.TopLine
        element.LeaderRightAttachment = element.LeaderRightAttachment.TopLine

        dotProduct = horVector.DotProduct(leaderEnd - Anchor)
        if dotProduct > 0:
            horVector = horVector*-1
            element.HorizontalAlignment = HorizontalTextAlignment.Right
        leader.Elbow = Anchor - horVector.Multiply(leaderLength)
        

    # Tag  
    if hasattr(element, "TagHeadPosition") and hasattr(element, "GetTaggedReferences") and hasattr(element, "SetLeaderElbow"):
        if hasattr(element, "HasLeader") and not element.HasLeader:
            continue

        refs = element.GetTaggedReferences()
        if refs is None or len(refs) == 0:
            continue

        leftPoint, rightPoint = get_tag_text_bounds_without_leader(
            element,
            refs[0],
            view,
            right_dir,
            up_dir,
            view_normal,
        )
        
        horVector = (rightPoint - leftPoint).Normalize()

        leaderEnd = element.GetLeaderEnd(refs[0])
        
        dotProduct = horVector.DotProduct(leaderEnd - leftPoint)
        if dotProduct > 0:
            horVector = horVector*-1
            leftPoint, rightPoint = rightPoint, leftPoint
        
        
        element.SetLeaderElbow(refs[0], leftPoint - horVector.Multiply(leaderLength))

t.Commit()


