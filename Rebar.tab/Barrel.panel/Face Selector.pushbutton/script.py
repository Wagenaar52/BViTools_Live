# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import * 
from Autodesk.Revit.UI.Selection import ObjectType
import Functions as func
from pyrevit import revit, forms, script
import clr
clr.AddReference("System")
from System import Int64
from System.Collections.Generic import List
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

class SupressWarnings(IFailuresPreprocessor):
    
    def PreprocessFailures(self, failuresAccessor):
        try:
            failures = failuresAccessor.GetFailureMessages()
            for failure in failures:
                severity = failure.GetSeverity()
                description = failure.GetDescriptionText()
                fail_Id = failure.GetFailureDefinitionId()

                if severity == FailureSeverity.Warning:
                    failuresAccessor.DeleteWarning(failure)
        except:
            import traceback
            print(traceback.format_exc())
        
        return FailureProcessingResult.Continue

##### VARIABLES ################################################################################

cover = 40 * 0.00328084
Spacing = 300 #mm

try:
    element_ref = uidoc.Selection.PickObject(ObjectType.Element, "Select element")
    element = doc.GetElement(element_ref.ElementId)
except Exception as exc:
    if "aborted the pick operation" in str(exc).lower():
        script.exit()
    raise




front_face_Id = 147
top_face_Id = 138
bot_face_Id = 127
back_face_Id = 181
left_face_Id = 132
right_face_Id = 142


def get_average_z_of_face_curveloops(face):
    curve_loops = face.GetEdgesAsCurveLoops()
    total_z = 0
    count = 0

    for loop in curve_loops:
        for curve in loop:
            total_z += curve.GetEndPoint(0).Z + curve.GetEndPoint(1).Z
            count += 2

    return total_z / count if count > 0 else None

def get_face_normal(face):
    bbox_uv = face.GetBoundingBox()
    midpoint = UV(
        (bbox_uv.Min.U + bbox_uv.Max.U) / 2.0,
        (bbox_uv.Min.V + bbox_uv.Max.V) / 2.0,
    )
    return face.ComputeNormal(midpoint).Normalize()

def get_face_points(face):
    points = []

    for loop in face.GetEdgesAsCurveLoops():
        for curve in loop:
            points.append(curve.GetEndPoint(0))
            points.append(curve.GetEndPoint(1))

    return points

def points_are_close(point_a, point_b, tol=1e-6):
    return point_a.DistanceTo(point_b) <= tol

def face_shares_points(face, reference_face, tol=1e-6):
    face_points = get_face_points(face)
    reference_points = get_face_points(reference_face)

    for face_point in face_points:
        for reference_point in reference_points:
            if points_are_close(face_point, reference_point, tol):
                return True

    return False

def get_face_center(face):
    points = get_face_points(face)
    if not points:
        return None

    total_x = 0.0
    total_y = 0.0
    total_z = 0.0
    for point in points:
        total_x += point.X
        total_y += point.Y
        total_z += point.Z

    point_count = float(len(points))
    return XYZ(total_x / point_count, total_y / point_count, total_z / point_count)

def get_adaptive_points(element):
    adaptive_points = []

    try:
        point_ids = AdaptiveComponentInstanceUtils.GetInstancePlacementPointElementRefIds(element)
    except Exception:
        point_ids = None

    if not point_ids:
        return adaptive_points

    for point_id in point_ids:
        point_element = doc.GetElement(point_id)
        if point_element is not None and hasattr(point_element, "Position"):
            adaptive_points.append(point_element.Position)

    return adaptive_points

def get_adaptive_plan_axes(element):
    adaptive_points = get_adaptive_points(element)
    if len(adaptive_points) < 2:
        return None, None, None

    origin = adaptive_points[0]
    point_2 = adaptive_points[1]
    local_y = XYZ(point_2.X - origin.X, point_2.Y - origin.Y, 0.0)
    if local_y.GetLength() == 0:
        return None, None, None

    local_y = local_y.Normalize()
    local_x = local_y.CrossProduct(XYZ.BasisZ)
    if local_x.GetLength() == 0:
        return None, None, None

    return origin, local_x.Normalize(), local_y

def is_point_on_face(point, face, tol=1e-6):
    # Check if the point is on the plane of the face
    n = face.GetSurface().Normal.Normalize()
    signed_dist = n.DotProduct(point - face.GetSurface().Origin)
    if abs(signed_dist) > tol:
        return False

    # Check if the point is within the face boundaries
    projection = face.Project(point)
    if projection is None:
        return False
    uv = projection.UVPoint
    return face.IsInside(uv)

def collect_top_bottom_face(element):
    options = Options()
    geometry = element.get_Geometry(options)
    face_candidates = []

    def collect_faces_from_geo(geo_obj):
        if isinstance(geo_obj, Face):
            face_candidates.append(geo_obj)
        elif isinstance(geo_obj, Solid) and geo_obj.Volume > 0:
            for face in geo_obj.Faces:
                face_candidates.append(face)

    for gobj in geometry:
        collect_faces_from_geo(gobj)
        if isinstance(gobj, GeometryInstance):
            for iobj in gobj.GetInstanceGeometry():
                collect_faces_from_geo(iobj)

    if not face_candidates:
        return None, None

    sorted_areas = sorted([face.Area for face in face_candidates], reverse=True)
    area_threshold = sorted_areas[min(6, len(sorted_areas) - 1)]
    ZList = []
    faceCan = []
    for f in face_candidates:
        if f.Area >= area_threshold:
            average_z = get_average_z_of_face_curveloops(f)
            if average_z is not None:
                ZList.append(average_z)
                faceCan.append(f)

    if not ZList:
        return None, None

    maxZ = max(ZList)
    minZ = min(ZList)
    top_face = None
    bot_face = None

    for f in faceCan:
        average_z = get_average_z_of_face_curveloops(f)
        if average_z == maxZ:
            top_face = f
        elif average_z == minZ:
            bot_face = f

    return top_face, bot_face

def collect_front_back_face(element):
    options = Options()
    geometry = element.get_Geometry(options)
    face_candidates = []

    def collect_faces_from_geo(geo_obj):
        if isinstance(geo_obj, Face):
            face_candidates.append(geo_obj)
        elif isinstance(geo_obj, Solid) and geo_obj.Volume > 0:
            for face in geo_obj.Faces:
                face_candidates.append(face)

    for gobj in geometry:
        collect_faces_from_geo(gobj)
        if isinstance(gobj, GeometryInstance):
            for iobj in gobj.GetInstanceGeometry():
                collect_faces_from_geo(iobj)

    if not face_candidates:
        return None, None

    front_face = None
    back_face = None

    adaptive_points = get_adaptive_points(element)
    if len(adaptive_points) >= 2:
        front_point = adaptive_points[0]
        back_point = adaptive_points[1]

        for f in face_candidates:
            if front_face is None and is_point_on_face(front_point, f):
                front_face = f
            elif back_face is None and is_point_on_face(back_point, f):
                back_face = f

        if front_face is not None and back_face is not None:
            return front_face, back_face

    facing_orientation = getattr(element, "FacingOrientation", None)
    if facing_orientation is None:
        return None, None

    face_scores = []
    for f in face_candidates:
        try:
            normal = get_face_normal(f)
        except Exception:
            continue

        if abs(normal.Z) > 0.5:
            continue

        alignment = normal.DotProduct(facing_orientation)
        face_scores.append((alignment, f.Area, f))

    if not face_scores:
        return None, None

    front_face = max(face_scores, key=lambda item: (item[0], item[1]))[2]
    back_face = min(face_scores, key=lambda item: (item[0], -item[1]))[2]

    return front_face, back_face

def collect_left_right_face(element, front_face, back_face, top_face, bot_face):
    options = Options()
    geometry = element.get_Geometry(options)
    face_candidates = []

    def collect_faces_from_geo(geo_obj):
        if isinstance(geo_obj, Face):
            face_candidates.append(geo_obj)
        elif isinstance(geo_obj, Solid) and geo_obj.Volume > 0:
            for face in geo_obj.Faces:
                face_candidates.append(face)

    for gobj in geometry:
        collect_faces_from_geo(gobj)
        if isinstance(gobj, GeometryInstance):
            for iobj in gobj.GetInstanceGeometry():
                collect_faces_from_geo(iobj)

    if not face_candidates:
        return None, None

    if front_face is None or back_face is None or top_face is None:
        return None, None

    try:
        front_normal = get_face_normal(front_face)
        top_normal = get_face_normal(top_face)
    except Exception:
        return None, None

    origin, local_x, local_y = get_adaptive_plan_axes(element)
    if local_x is None:
        side_axis = front_normal.CrossProduct(top_normal)
        if side_axis.GetLength() == 0:
            return None, None
        side_axis = side_axis.Normalize()
    else:
        side_axis = local_x

    candidate_faces = []
    for f in face_candidates:
        if f.Id == front_face.Id or f.Id == back_face.Id:
            continue

        if bot_face is not None and f.Id == bot_face.Id:
            continue

        if top_face is not None and f.Id == top_face.Id:
            continue

        shares_front = face_shares_points(f, front_face)
        shares_back = face_shares_points(f, back_face)
        if not (shares_front and shares_back):
            continue

        try:
            normal = get_face_normal(f)
        except Exception:
            continue

        if abs(normal.DotProduct(front_normal)) > 0.4:
            continue

        if abs(normal.DotProduct(top_normal)) > 0.4:
            continue

        face_center = get_face_center(f)
        if face_center is None:
            continue

        if origin is None:
            offset = face_center.DotProduct(side_axis)
        else:
            offset = (face_center - origin).DotProduct(side_axis)
        candidate_faces.append((offset, f.Area, f))

    if len(candidate_faces) < 2:
        return None, None

    left_face = min(candidate_faces, key=lambda item: (item[0], -item[1]))[2]
    right_face = max(candidate_faces, key=lambda item: (item[0], item[1]))[2]

    return left_face, right_face

top_face, bot_face = collect_top_bottom_face(element)
front_face, back_face = collect_front_back_face(element)
left_face, right_face = collect_left_right_face(element, front_face, back_face, top_face, bot_face)

if top_face is None or bot_face is None:
    raise Exception("Unable to identify top and bottom faces for the selected element.")

if front_face is None or back_face is None:
    raise Exception("Unable to identify front and back faces for the selected element.")

if left_face is None or right_face is None:
    raise Exception("Unable to identify left and right faces for the selected element.")

print("Top Face Id: " + str(top_face.Id))
print("Bottom Face Id: " + str(bot_face.Id))
print("Front Face Id: " + str(front_face.Id))
print("Back Face Id: " + str(back_face.Id))
print("Left Face Id: " + str(left_face.Id))
print("Right Face Id: " + str(right_face.Id))   
