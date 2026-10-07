# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import * 
import Functions as func
from pyrevit import revit, forms, script
import clr
import math
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

element = doc.GetElement(ElementId(Int64(552179)))

hwfront_left_Id = 314
hwfront_right_Id = 316
hwback_left_Id = 295
hwback_right_Id = 297
hwTop_front_Id = 306
hwTop_back_Id = 292
hwback_back_Id = 282
hwfront_front_Id = 301


cwfront_front_Id = 319
cwfront_left_Id = 321
cwfront_right_Id = 323
cwback_back_Id = 325


botSlabHeight = element.LookupParameter("Flr_T").AsDouble()
topSlabHeight = element.LookupParameter("Roof_T").AsDouble()
Cul_Haunch_Bot_D = element.LookupParameter("Cul_Haunch_Bot_D").AsDouble()

HeadW1_H = element.LookupParameter("HeadW1_H").AsDouble()
HeadW2_H = element.LookupParameter("HeadW2_H").AsDouble()
HeadW1_T = element.LookupParameter("HeadW1_T").AsDouble()
HeadW2_T = element.LookupParameter("HeadW2_T").AsDouble()

###### MARKS ##################################################################################

Headwall_20 = "Head Wall Horizontal"
Headwall_38 = "Head Wall U-bars"
Headwall_60 = "Head Wall stirrups"
HeadwallDownStream_20 = "Down Stream Head Wall Horizontal"
HeadwallDownStream_38 = "Down Stream Head Wall U-bars"
HeadwallDownStream_60 = "Down Stream Head Wall stirrups"

#####Rebar Shape ###############################################################################

def getRebarShapeByName(name):
    rebar_shape = FilteredElementCollector(doc).OfClass(RebarShape).WhereElementIsElementType().ToElements()
    for r_shape in rebar_shape:
        if r_shape.LookupParameter("Type Name").AsString() == name:
            return r_shape
    return None

##### Rebar type ###############################################################################
    
all_rebar_types = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsElementType().ToElements()

def barTypeBySize(size):
    bar_type = next((rebar_type for rebar_type in all_rebar_types 
                 if rebar_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == size), None)
    return bar_type

################################################################################################
sc_20 = getRebarShapeByName("20")
sc_38 = getRebarShapeByName("38")
sc_55 = getRebarShapeByName("55")
sc_60 = getRebarShapeByName("60")
barType16 = barTypeBySize("Y16") # 16mm rebar
barType12 = barTypeBySize("Y12") # 12mm rebar
barSize = 16 * 0.00328084 # 16mm rebar

################################################################################################

def collect_faces_from_geo(geo_obj):
    if isinstance(geo_obj, Face):
        face_candidates.append(geo_obj)
    elif isinstance(geo_obj, Solid) and geo_obj.Volume > 0:
        for face in geo_obj.Faces:
            face_candidates.append(face)

def collect_curves_from_geo(geo_obj):
    if isinstance(geo_obj, Curve):
        curve_candidates.append(geo_obj)
    elif isinstance(geo_obj, Solid) and geo_obj.Volume > 0:
        for edge in geo_obj.Edges:
            edge_curve = edge.AsCurve()
            if edge_curve is not None:
                curve_candidates.append(edge_curve)

def is_line_in_plane(line, plane, tol=1e-6):
    n = plane.Normal.Normalize()
    v = line.Direction.Normalize()

    # 1) Line direction must be perpendicular to plane normal
    if abs(n.DotProduct(v)) > tol:
        return False

    # 2) A point on the line must satisfy plane equation
    p0 = line.GetEndPoint(0) if line.IsBound else line.Origin
    signed_dist = n.DotProduct(p0 - plane.Origin)

    return abs(signed_dist) <= tol

def is_point_on_face(point, face, tol=1e-9):
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

def get_bottom_point_of_line(line):
    if line.IsBound:
        p1 = line.GetEndPoint(0)
        p2 = line.GetEndPoint(1)
        return p1 if p1.Z < p2.Z else p2
    else:
        return line.Origin

def get_top_point_of_line(line):
    if line.IsBound:
        p1 = line.GetEndPoint(0)
        p2 = line.GetEndPoint(1)
        return p1 if p1.Z > p2.Z else p2
    else:
        return line.Origin

def is_line_vertical(line, tol=1e-9):
    direction = line.Direction.Normalize()
    vertical = XYZ.BasisZ
    return abs(direction.DotProduct(vertical)) > 1 - tol

def draw_plane_in_view(plane, size=3.0):
    if plane is None:
        return

    normal = plane.Normal.Normalize()
    origin = plane.Origin

    if abs(normal.DotProduct(XYZ.BasisZ)) < 0.95:
        ref_axis = XYZ.BasisZ
    else:
        ref_axis = XYZ.BasisX

    axis_u = normal.CrossProduct(ref_axis).Normalize()
    axis_v = normal.CrossProduct(axis_u).Normalize()

    p1 = origin + axis_u * size + axis_v * size
    p2 = origin - axis_u * size + axis_v * size
    p3 = origin - axis_u * size - axis_v * size
    p4 = origin + axis_u * size - axis_v * size

    sketch_plane = SketchPlane.Create(doc, plane)
    doc.Create.NewModelCurve(Line.CreateBound(p1, p2), sketch_plane)
    doc.Create.NewModelCurve(Line.CreateBound(p2, p3), sketch_plane)
    doc.Create.NewModelCurve(Line.CreateBound(p3, p4), sketch_plane)
    doc.Create.NewModelCurve(Line.CreateBound(p4, p1), sketch_plane)

def draw_line_in_view(start , end):
    if start is None or end is None:
        return
    doc.Create.NewModelCurve(Line.CreateBound(start, end), SketchPlane.Create(doc, Plane.CreateByThreePoints(start, end, XYZ(0,0,0))))

def get_average_z_of_face_curveloops(face):
    curve_loops = face.GetEdgesAsCurveLoops()
    total_z = 0
    count = 0
    for loop in curve_loops:
        for curve in loop:
            total_z += curve.GetEndPoint(0).Z + curve.GetEndPoint(1).Z
            count += 2
    return total_z / count if count > 0 else None

def collect_top_face(element):
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

    sixth_maxArea = sorted([face.Area for face in face_candidates], reverse=True)[6]
    ZList = []
    faceCan = []
    for f in face_candidates:
        if f.Area > sixth_maxArea:
             ZList.append(get_average_z_of_face_curveloops(f))
             faceCan.append(f)
    maxZ = max(ZList)
    for f in faceCan:
        if get_average_z_of_face_curveloops(f) == maxZ:
            top_face = f

    return top_face

def angle_to_horizontal_two_points(point1, point2):
    vec = (point2 - point1)
    horizontal_vec = XYZ(vec.X, vec.Y, 0)
    angle = horizontal_vec.AngleTo(vec) 
    return angle

def sc20_by2points(rebar_p1, rebar_p2, Plane, element, barType =barType16, document = doc, sc_20 = sc_20, ModelLine = False):
        #place curves
        curve1 = Line.CreateBound(rebar_p1, rebar_p2)
        geomPlane = Plane.CreateByThreePoints(rebar_p1, rebar_p2, Plane.Origin )
        sketch = SketchPlane.Create(doc, geomPlane)

        if ModelLine == True:
            model_line = doc.Create.NewModelCurve(curve1, sketch)
        else:
           #### Cast the list to IList<Curve>
            curve_list20 = List[Curve]([curve1])
            
            #### Bluid ####################################################################        
            rebar = Structure.Rebar.CreateFromCurvesAndShape(doc, 
                                                sc_20, 
                                                barType, 
                                                element,
                                                Plane.Normal,  
                                                curve_list20, 
                                                BarTerminationsData(doc))
            
        return rebar

print("Family geometry properties:")

options = Options()
geometry = element.get_Geometry(options)
curve_candidates = []
face_candidates = []

for gobj in geometry:
    collect_curves_from_geo(gobj)
    collect_faces_from_geo(gobj)
    if isinstance(gobj, GeometryInstance):
        for iobj in gobj.GetInstanceGeometry():
            collect_curves_from_geo(iobj)
            collect_faces_from_geo(iobj)

print("Number of curves:", len(curve_candidates))
print("Number of faces:", len(face_candidates))
print("#"*150)

for f in face_candidates:
    if f.Id == front_face_Id:
        front_face = f
    if f.Id == top_face_Id:
        top_face = f
    if f.Id == bot_face_Id:
        bot_face = f
    if f.Id == back_face_Id:
        back_face = f
    if f.Id == left_face_Id:
        left_face = f
    if f.Id == right_face_Id:
        right_face = f

print("Front face Id:", front_face.Id)
print("Top face Id:", top_face.Id)  
print("Bottom face Id:", bot_face.Id)
print("Back face Id:", back_face.Id)
print("Left face Id:", left_face.Id)
print("Right face Id:", right_face.Id)


def collect_corner_lines(curve_candidates, front_face, back_face, left_face, right_face):

    FrontrebarAnchorLine = []
    BackrebarAnchorLine = []
    for c in curve_candidates:
        if is_line_in_plane(c, front_face.GetSurface()):
            if abs(c.Direction.Normalize().Z) > 0.9:
                FrontrebarAnchorLine.append(c)
        if is_line_in_plane(c, back_face.GetSurface()):
            if abs(c.Direction.Normalize().Z) > 0.9:
                BackrebarAnchorLine.append(c)

    print("Number of rebar anchor lines:", len(FrontrebarAnchorLine))
    print("Number of rebar anchor lines:", len(BackrebarAnchorLine))
    print("#"*150)

    # remove lines that are not on the top of the element, but head wall lines in the front face plane
    maxZfront = max(max(line.GetEndPoint(0).Z, line.GetEndPoint(1).Z) for line in FrontrebarAnchorLine)
    for line in FrontrebarAnchorLine:
        if line.GetEndPoint(0).Z== maxZfront or line.GetEndPoint(1).Z == maxZfront:
            FrontrebarAnchorLine.remove(line)

    maxZback = max(max(line.GetEndPoint(0).Z, line.GetEndPoint(1).Z) for line in BackrebarAnchorLine)
    for line in BackrebarAnchorLine:
        if line.GetEndPoint(0).Z== maxZback or line.GetEndPoint(1).Z == maxZback:
            BackrebarAnchorLine.remove(line)

    #remove second line
    maxZfront = max(max(line.GetEndPoint(0).Z, line.GetEndPoint(1).Z) for line in FrontrebarAnchorLine)
    for line in FrontrebarAnchorLine:
        if line.GetEndPoint(0).Z== maxZfront or line.GetEndPoint(1).Z == maxZfront:
            FrontrebarAnchorLine.remove(line)

    maxZback = max(max(line.GetEndPoint(0).Z, line.GetEndPoint(1).Z) for line in BackrebarAnchorLine)
    for line in BackrebarAnchorLine:
        if line.GetEndPoint(0).Z== maxZback or line.GetEndPoint(1).Z == maxZback:
            BackrebarAnchorLine.remove(line)

    print("Number of rebar anchor lines after removing head wall lines:", len(FrontrebarAnchorLine))
    print("Number of rebar anchor lines after removing head wall lines:", len(BackrebarAnchorLine))
    print("#"*150)

    leftLineFront = None
    rightLineFront = None
    leftLineBack = None
    rightLineBack = None
    # at front get side faces and set to left_face and right_face
    min_line_length = min(c.Length for c in FrontrebarAnchorLine)
    for line in FrontrebarAnchorLine:
        if line.Length > min_line_length*1.01 and is_line_in_plane(line, left_face.GetSurface()): # if the line is more than 1% longer than the shortest line, we consider it as an outer line
            leftLineFront = line
        elif line.Length > min_line_length*1.01 and is_line_in_plane(line, right_face.GetSurface()):
            rightLineFront = line
    # at back get side faces and set to left_face and right_face
    min_line_length = min(c.Length for c in BackrebarAnchorLine)
    for line in BackrebarAnchorLine:
        if line.Length > min_line_length*1.01 and is_line_in_plane(line, left_face.GetSurface()): # if the line is more than 1% longer than the shortest line, we consider it as an outer line
            leftLineBack = line
        elif line.Length > min_line_length*1.01 and is_line_in_plane(line, right_face.GetSurface()):
            rightLineBack = line

    return leftLineFront, rightLineFront, leftLineBack, rightLineBack

def get_faces_above_top_face(face_candidates, top_face):
    faces_above = []
    for face in face_candidates:
        if get_average_z_of_face_curveloops(face) > get_average_z_of_face_curveloops(top_face):
            faces_above.append(face)
    return faces_above

def get_points_of_curveloop(curve_loop):
    points = []
    for curve in curve_loop:
        points.append(curve.GetEndPoint(0))
        points.append(curve.GetEndPoint(1))
    return points

leftLineFront, rightLineFront, leftLineBack, rightLineBack = collect_corner_lines(curve_candidates, front_face, back_face, left_face, right_face)

print(get_faces_above_top_face(face_candidates, top_face))
print("$$$$$$$$$$$$$$$ ")
collect_corner_lines(curve_candidates, front_face, back_face, left_face, right_face)
leftTransverseLine = Line.CreateBound(get_top_point_of_line(leftLineFront), get_top_point_of_line(leftLineBack))
rightTransverseLine = Line.CreateBound(get_top_point_of_line(rightLineFront), get_top_point_of_line(rightLineBack))
frontTransverseLine = Line.CreateBound(get_top_point_of_line(leftLineFront), get_top_point_of_line(rightLineFront))
backTransverseLine = Line.CreateBound(get_top_point_of_line(leftLineBack), get_top_point_of_line(rightLineBack))

print(leftTransverseLine, rightTransverseLine, frontTransverseLine, backTransverseLine)


for face in get_faces_above_top_face(face_candidates, top_face):
    if is_line_in_plane(frontTransverseLine, face.GetSurface()):
        hwfront_front = face
    elif is_line_in_plane(backTransverseLine, face.GetSurface()):
        hwback_back = face
    elif is_line_in_plane(leftTransverseLine, face.GetSurface()):
        if face.GetSurface().Origin.DistanceTo(get_top_point_of_line(leftLineFront)) < face.GetSurface().Origin.DistanceTo(get_top_point_of_line(leftLineBack)):
            hwfront_left = face
            print("hwfront left found")
        else:
            hwback_left = face
            print("hwback left found")

    elif is_line_in_plane(rightTransverseLine, face.GetSurface()):
        if face.GetSurface().Origin.DistanceTo(get_top_point_of_line(leftLineFront)) < face.GetSurface().Origin.DistanceTo(get_top_point_of_line(leftLineBack)):
            hwfront_right = face
            print("hwfront right found")
        else:
            hwback_right = face
            print("hwback right found")

for face in face_candidates:
    if face.Id == hwfront_left_Id:
        hwfront_left = face
    elif face.Id == hwfront_right_Id:
        hwfront_right = face
    elif face.Id == hwback_left_Id:
        hwback_left = face
    elif face.Id == hwback_right_Id:
        hwback_right = face
    elif face.Id == hwback_back_Id:
        hwback_back = face
    elif face.Id == hwfront_front_Id:
        hwfront_front = face



faceList = []
for face in get_faces_above_top_face(face_candidates, top_face):
    if abs(face.GetSurface().Normal.Normalize().DotProduct(XYZ.BasisZ)) > 0.9:
        faceList.append(face)

faceList = sorted(faceList, key=lambda face: get_average_z_of_face_curveloops(face), reverse=True)

if len(faceList) < 2:
    raise Exception("Unable to identify both head wall top faces.")

if faceList[0].GetSurface().Origin.DistanceTo(get_top_point_of_line(leftLineFront)) < faceList[1].GetSurface().Origin.DistanceTo(get_top_point_of_line(leftLineFront)):
    hwTop_front = faceList[0]
    hwTop_back = faceList[1]
else:
    hwTop_front = faceList[1]
    hwTop_back = faceList[0]

for face in face_candidates:
    if face.Id == hwTop_front_Id:
        hwTop_front = face
    if face.Id == hwTop_back_Id:
        hwTop_back = face

t = Transaction(doc, "Create Headwall Rebar")
t.Start()

#region -  Front Head Wall front longitudinal rebar 20
##################################################################################################################################################################
top_point = get_top_point_of_line(leftLineFront)
vertical_vector = leftLineFront.Direction.Normalize()
horizontal_vector = frontTransverseLine.Direction.Normalize()
# rebar = Structure.Rebar.CreateFromRebarShape(doc, sc_20, barType16, element, top_point, horizontal_vector,vertical_vector)
rebar = sc20_by2points(top_point, get_top_point_of_line(rightLineFront), hwTop_front.GetSurface(), element, barType16, doc, sc_20)
rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(Spacing * 0.00328084, 5,True , True, True) #SetLayoutAsMaximumSpacing(double spacing,double arrayLength,bool barsOnNormalSide,bool includeFirstBar,bool includeLastBar)
shape_accessor = rebar.GetShapeDrivenAccessor()
shape_accessor.UseRebarConstraintsToProduceVaryingBars = True
rebar.LookupParameter("Mark").Set(Headwall_20)

handleList = rebar.GetRebarConstraintsManager().GetAllHandles()
conman = rebar.GetRebarConstraintsManager()

for handle in handleList:
    print("Starting handle:" + str(handle.GetHandleName()))

    if handle.GetHandleName() == "Out of Plane Extent":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == hwTop_front.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(-barSize)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover -barSize
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Out of Plane Extent - Handle processed #####")

    if handle.GetHandleName() == "Start of Bar":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == hwfront_left.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(0)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Start of Bar - Handle processed #####")   

    if handle.GetHandleName() == "End of Bar":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == right_face.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(0)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### End of Bar - Handle processed #####") 

    if handle.GetHandleName() == "Bar Plane":
        tempFace = Plane.CreateByNormalAndOrigin(hwTop_front.GetSurface().Normal, top_point + XYZ(0,0,1) * ( + cover))
        constraint = RebarConstraint.CreateConstraintToSurface(handle, tempFace)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()
        print("##### Bar Plane - Handle processed #####") 

    if handle.GetHandleName() == "Bar Segment 1":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == front_face.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(- barSize)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover - barSize
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Bar Segment 1 - Handle processed #####") 

####Front Head Wall back longitudinal rebar 20 ##############################################################################################################################################################
vertical_vector = leftLineFront.Direction.Normalize()
horizontal_vector = frontTransverseLine.Direction.Normalize()
top_point = get_top_point_of_line(leftLineFront) + leftTransverseLine.Direction.Normalize() * (HeadW1_T - cover)
topRightPoint = get_top_point_of_line(rightLineFront) + rightTransverseLine.Direction.Normalize() * (HeadW1_T - cover)
hwTop_backSur = Plane.CreateByNormalAndOrigin(hwTop_front.GetSurface().Normal, hwTop_front.GetSurface().Origin + hwTop_front.GetSurface().Normal * (HeadW1_T - cover))
# rebar = Structure.Rebar.CreateFromRebarShape(doc, sc_20, barType16, element, top_point, horizontal_vector,vertical_vector)
rebar = sc20_by2points(top_point, topRightPoint, hwTop_backSur, element, barType16, doc, sc_20)
rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(Spacing * 0.00328084, 5,True , True, True) #SetLayoutAsMaximumSpacing(double spacing,double arrayLength,bool barsOnNormalSide,bool includeFirstBar,bool includeLastBar)
shape_accessor = rebar.GetShapeDrivenAccessor()
shape_accessor.UseRebarConstraintsToProduceVaryingBars = True
rebar.LookupParameter("Mark").Set(Headwall_20)

handleList = rebar.GetRebarConstraintsManager().GetAllHandles()
conman = rebar.GetRebarConstraintsManager()

for handle in handleList:
    print("Starting handle:" + str(handle.GetHandleName()))

    if handle.GetHandleName() == "Out of Plane Extent":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == hwTop_front.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(- barSize)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover -barSize
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Out of Plane Extent - Handle processed #####")

    if handle.GetHandleName() == "Start of Bar":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == hwfront_left.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(0)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Start of Bar - Handle processed #####")   

    if handle.GetHandleName() == "End of Bar":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == right_face.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(0)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### End of Bar - Handle processed #####") 

    if handle.GetHandleName() == "Bar Plane":
        tempFace = Plane.CreateByNormalAndOrigin(hwTop_front.GetSurface().Normal, top_point + XYZ(0,0,1) * ( + cover))
        constraint = RebarConstraint.CreateConstraintToSurface(handle, tempFace)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()
        print("##### Bar Plane - Handle processed #####") 

    if handle.GetHandleName() == "Bar Segment 1":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == front_face.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(-HeadW1_T+ 2*cover + 2*barSize)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover - barSize + HeadW1_T
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Bar Segment 1 - Handle processed #####") 


##################################################################################################################################################################
#endregion

#region -  Back Head Wall front longitudinal rebar 20
##################################################################################################################################################################
top_point = get_top_point_of_line(leftLineBack)
vertical_vector = leftLineBack.Direction.Normalize()
horizontal_vector = backTransverseLine.Direction.Normalize()
# rebar = Structure.Rebar.CreateFromRebarShape(doc, sc_20, barType16, element, top_point, horizontal_vector,vertical_vector)
rebar = sc20_by2points(top_point, get_top_point_of_line(rightLineBack), hwTop_back.GetSurface(), element, barType16, doc, sc_20)
rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(Spacing * 0.00328084, 5,True , True, True) #SetLayoutAsMaximumSpacing(double spacing,double arrayLength,bool barsOnNormalSide,bool includeFirstBar,bool includeLastBar)
shape_accessor = rebar.GetShapeDrivenAccessor()
shape_accessor.UseRebarConstraintsToProduceVaryingBars = True
rebar.LookupParameter("Mark").Set(HeadwallDownStream_20)

handleList = rebar.GetRebarConstraintsManager().GetAllHandles()
conman = rebar.GetRebarConstraintsManager()

for handle in handleList:
    print("Starting handle:" + str(handle.GetHandleName()))

    if handle.GetHandleName() == "Out of Plane Extent":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == hwTop_back.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(0)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Out of Plane Extent - Handle processed #####")

    if handle.GetHandleName() == "Start of Bar":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == hwback_left.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(0)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Start of Bar - Handle processed #####")   

    if handle.GetHandleName() == "End of Bar":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == right_face.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(0)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### End of Bar - Handle processed #####") 

    if handle.GetHandleName() == "Bar Plane":
        tempFace = Plane.CreateByNormalAndOrigin(hwTop_back.GetSurface().Normal, top_point + XYZ(0,0,1) * ( + cover))
        constraint = RebarConstraint.CreateConstraintToSurface(handle, tempFace)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()
        print("##### Bar Plane - Handle processed #####") 

    if handle.GetHandleName() == "Bar Segment 1":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == back_face.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(- barSize)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover - barSize
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Bar Segment 1 - Handle processed #####") 

####Back Head Wall back longitudinal rebar 20 ##############################################################################################################################################################
vertical_vector = leftLineBack.Direction.Normalize()
horizontal_vector = backTransverseLine.Direction.Normalize()
top_point = get_top_point_of_line(leftLineBack) + leftTransverseLine.Direction.Normalize() * (HeadW1_T - cover)
topRightPoint = get_top_point_of_line(rightLineBack) + rightTransverseLine.Direction.Normalize() * (HeadW1_T - cover)
hwback_backSur1 = Plane.CreateByNormalAndOrigin(hwTop_back.GetSurface().Normal, hwTop_back.GetSurface().Origin + hwTop_back.GetSurface().Normal * (- cover))
# rebar = Structure.Rebar.CreateFromRebarShape(doc, sc_20, barType16, element, top_point, horizontal_vector,vertical_vector)
rebar = sc20_by2points(top_point, topRightPoint, hwback_backSur1, element, barType16, doc, sc_20)
rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(Spacing * 0.00328084, 5,True , True, True) #SetLayoutAsMaximumSpacing(double spacing,double arrayLength,bool barsOnNormalSide,bool includeFirstBar,bool includeLastBar)
shape_accessor = rebar.GetShapeDrivenAccessor()
shape_accessor.UseRebarConstraintsToProduceVaryingBars = True
rebar.LookupParameter("Mark").Set(HeadwallDownStream_20)

handleList = rebar.GetRebarConstraintsManager().GetAllHandles()
conman = rebar.GetRebarConstraintsManager()

for handle in handleList:
    print("Starting handle:" + str(handle.GetHandleName()))

    if handle.GetHandleName() == "Out of Plane Extent":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == hwTop_back.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(0)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Out of Plane Extent - Handle processed #####")

    if handle.GetHandleName() == "Start of Bar":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == hwback_left.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(0)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Start of Bar - Handle processed #####")   

    if handle.GetHandleName() == "End of Bar":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == right_face.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(0)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### End of Bar - Handle processed #####") 

    if handle.GetHandleName() == "Bar Plane":
        tempFace = Plane.CreateByNormalAndOrigin(hwTop_back.GetSurface().Normal, top_point + XYZ(0,0,1) * ( + cover))
        constraint = RebarConstraint.CreateConstraintToSurface(handle, tempFace)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()
        print("##### Bar Plane - Handle processed #####") 

    if handle.GetHandleName() == "Bar Segment 1":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == back_face.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(2*cover + barSize*2 - HeadW2_T)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover - barSize*2 + HeadW2_T
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Bar Segment 1 - Handle processed #####") 


##################################################################################################################################################################
#endregion

print("#"*100)

#region -  Front Head Wall rebar 60
##################################################################################################################################################################
top_point = get_top_point_of_line(leftLineFront)
vertical_vector = leftLineFront.Direction.Normalize()
horizontal_vector = XYZ(leftTransverseLine.Direction.Normalize().X, leftTransverseLine.Direction.Normalize().Y, 0)
rebar = Structure.Rebar.CreateFromRebarShape(doc, sc_60, barType12, element, top_point, horizontal_vector,vertical_vector)
# rebar = sc20_by2points(top_point, get_top_point_of_line(rightLineFront), hwTop_front.GetSurface(), element, barType16, doc, sc_20)
rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(Spacing * 0.00328084, 5,True , True, True) #SetLayoutAsMaximumSpacing(double spacing,double arrayLength,bool barsOnNormalSide,bool includeFirstBar,bool includeLastBar)
shape_accessor = rebar.GetShapeDrivenAccessor()
shape_accessor.UseRebarConstraintsToProduceVaryingBars = True
rebar.LookupParameter("Mark").Set(Headwall_60)
rebar.LookupParameter("A").Set(HeadW1_H + topSlabHeight - 2*cover)
rebar.LookupParameter("B").Set(HeadW1_T - cover)

handleList = rebar.GetRebarConstraintsManager().GetAllHandles()
conman = rebar.GetRebarConstraintsManager()

for handle in handleList:
    print("Starting handle:" + str(handle.GetHandleName()))

    if handle.GetHandleName() == "Out of Plane Extent":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == hwfront_right.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(0)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover 
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Out of Plane Extent - Handle processed #####")

    # if handle.GetHandleName() == "Start of Bar":
    #     constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
    #     for const in constraint:
    #         conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
    #         if conSur.Id == hwfront_left.Id :
    #             constraint = const
    #             break
    #     if constraint and constraint.IsToCover():
    #         constraint.SetDistanceToTargetCover(0)
    #         conman.SetPreferredConstraint(constraint)
    #         doc.Regenerate()
    #     elif constraint and constraint.IsToHostFaceOrCover(): 
    #         new_offset = -cover
    #         constraint.SetDistanceToTargetHostFace(new_offset)
    #         conman.SetPreferredConstraint(constraint)
    #         doc.Regenerate()
    #     print("##### Start of Bar - Handle processed #####")   

    if handle.GetHandleName() == "Start of Bar":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            print(conSur.Id)
            if conSur.Id == front_face.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(0)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = cover
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Start of Bar - Handle processed #####") 

    if handle.GetHandleName() == "Bar Plane":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == hwfront_left.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(0)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Bar Plane - Handle processed #####") 

    if handle.GetHandleName() == "Bar Segment 1":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == hwTop_front.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(0)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover - barSize
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Bar Segment 1 - Handle processed #####") 

    if handle.GetHandleName() == "Bar Segment 2":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == hwfront_front.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(0)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover 
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Bar Segment 2 - Handle processed #####") 

    # if handle.GetHandleName() == "Bar Segment 3":
    #     constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
    #     for const in constraint:
    #         conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
    #         if conSur.Id == hwTop_front.Id :
    #             constraint = const
    #             break
    #     if constraint and constraint.IsToCover():
    #         constraint.SetDistanceToTargetCover(HeadW1_H - cover)
    #         conman.SetPreferredConstraint(constraint)
    #         doc.Regenerate()
    #     elif constraint and constraint.IsToHostFaceOrCover(): 
    #         new_offset = cover  - HeadW1_H
    #         constraint.SetDistanceToTargetHostFace(new_offset)
    #         conman.SetPreferredConstraint(constraint)
    #         doc.Regenerate()
    #     print("##### Bar Segment 3 - Handle processed #####") 

    # if handle.GetHandleName() == "Bar Segment 4":
    #     constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
    #     for const in constraint:
    #         conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
    #         if conSur.Id == hwfront_front.Id :
    #             constraint = const
    #             break
    #         conman.SetPreferredConstraint(constraint)
    #         doc.Regenerate()
    #     if constraint and constraint.IsToCover():
    #         constraint.SetDistanceToTargetCover(200*0.003)    # HeadW1_T )# -2*cover)
    #         conman.SetPreferredConstraint(constraint)
    #         doc.Regenerate()
    #     elif constraint and constraint.IsToHostFaceOrCover(): 
    #         new_offset = cover  - HeadW1_T
    #         constraint.SetDistanceToTargetHostFace(new_offset)
    #         conman.SetPreferredConstraint(constraint)
    #         doc.Regenerate()
    #     print("##### Bar Segment 4 - Handle processed #####") 

##################################################################################################################################################################
#endregion

#region -  Back Head Wall rebar 60
##################################################################################################################################################################
top_point = get_top_point_of_line(leftLineBack)
vertical_vector = leftLineFront.Direction.Normalize()
horizontal_vector = XYZ(leftTransverseLine.Direction.Normalize().X, leftTransverseLine.Direction.Normalize().Y, 0)
rebar = Structure.Rebar.CreateFromRebarShape(doc, sc_60, barType12, element, top_point, -horizontal_vector,vertical_vector)
# rebar = sc20_by2points(top_point, get_top_point_of_line(rightLineFront), hwTop_front.GetSurface(), element, barType16, doc, sc_20)
rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(Spacing * 0.00328084, 5,True , True, True) #SetLayoutAsMaximumSpacing(double spacing,double arrayLength,bool barsOnNormalSide,bool includeFirstBar,bool includeLastBar)
shape_accessor = rebar.GetShapeDrivenAccessor()
shape_accessor.UseRebarConstraintsToProduceVaryingBars = True
rebar.LookupParameter("Mark").Set(HeadwallDownStream_60)
rebar.LookupParameter("A").Set(HeadW2_H + topSlabHeight - 2*cover)
rebar.LookupParameter("B").Set(HeadW2_T - cover)

handleList = rebar.GetRebarConstraintsManager().GetAllHandles()
conman = rebar.GetRebarConstraintsManager()

for handle in handleList:
    print("Starting handle:" + str(handle.GetHandleName()))

    if handle.GetHandleName() == "Out of Plane Extent":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == hwback_right.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(0)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover 
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Out of Plane Extent - Handle processed #####")

    if handle.GetHandleName() == "Start of Bar":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            print(conSur.Id)
            if conSur.Id == back_face.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(0)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = cover
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Start of Bar - Handle processed #####") 

    if handle.GetHandleName() == "Bar Plane":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == hwback_left.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(0)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Bar Plane - Handle processed #####") 

    if handle.GetHandleName() == "Bar Segment 1":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == hwTop_back.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(0)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover - barSize
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Bar Segment 1 - Handle processed #####") 

    if handle.GetHandleName() == "Bar Segment 2":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == hwback_back.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(0)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover 
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Bar Segment 2 - Handle processed #####") 

##################################################################################################################################################################
#endregion

print("#"*100)

#region -  ubars on front head wall
##################################################################################################################################################################
top_point = get_top_point_of_line(leftLineFront)
vertical_vector = frontTransverseLine.Direction.Normalize()
horizontal_vector = XYZ.BasisZ.CrossProduct(vertical_vector)                      

seg2Plane = Plane.CreateByNormalAndOrigin(vertical_vector, top_point)
offseg2Plane = Plane.CreateByNormalAndOrigin(vertical_vector, top_point + vertical_vector.Normalize() * (barSize*50))

rebar = Structure.Rebar.CreateFromRebarShape(doc, sc_38, barType16, element, top_point, horizontal_vector,vertical_vector)
# rebar = sc20_by2points(top_point, get_top_point_of_line(rightLineFront), hwTop_front.GetSurface(), element, barType16, doc, sc_20)
rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(Spacing * 0.00328084, 5,False , True, True) #SetLayoutAsMaximumSpacing(double spacing,double arrayLength,bool barsOnNormalSide,bool includeFirstBar,bool includeLastBar)
shape_accessor = rebar.GetShapeDrivenAccessor()
shape_accessor.UseRebarConstraintsToProduceVaryingBars = True
rebar.LookupParameter("Mark").Set(Headwall_38)
rebar.LookupParameter("B").Set(HeadW1_T - cover*2 - barSize*2)
rebar.LookupParameter("A").Set(barSize*50)

rebarNormalVec = rebar.GetShapeDrivenAccessor().GetDistributionPath().Direction 
handleList = rebar.GetRebarConstraintsManager().GetAllHandles()
conman = rebar.GetRebarConstraintsManager()

for handle in handleList:
    print("Starting handle:" + str(handle.GetHandleName()))

    if handle.GetHandleName() == "Out of Plane Extent":
        tempFace = Plane.CreateByNormalAndOrigin(rebarNormalVec, hwTop_front.Origin + rebarNormalVec.Normalize() * (- cover - barSize*2))
        constraint = RebarConstraint.CreateConstraintToSurface(handle, tempFace)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()
        print("$$$"*50)
        print("##### Out of Plane Extent - Handle processed #####")

    if handle.GetHandleName() == "Start of Bar":
        constraint = RebarConstraint.CreateConstraintToSurface(handle, offseg2Plane)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()
        print("##### Start of Bar - Handle processed #####")   

    if handle.GetHandleName() == "End of Bar":
        constraint = RebarConstraint.CreateConstraintToSurface(handle, offseg2Plane)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()
        print("##### End of Bar - Handle processed #####") 

    if handle.GetHandleName() == "Bar Segment 1":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == front_face.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(- barSize)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover - barSize
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Bar Segment 1 - Handle processed #####") 

    if handle.GetHandleName() == "Bar Segment 2":
        constraint = RebarConstraint.CreateConstraintToSurface(handle, seg2Plane)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()
        print("##### Bar Segment 2 - Handle processed #####") 


##################################################################################################################################################################
# RIGTH SIDE OF FRONT HEAD WALL
top_point = get_top_point_of_line(rightLineFront)
vertical_vector = -frontTransverseLine.Direction.Normalize()
horizontal_vector = -XYZ.BasisZ.CrossProduct(vertical_vector)                      

seg2Plane = Plane.CreateByNormalAndOrigin(vertical_vector, top_point)
offseg2Plane = Plane.CreateByNormalAndOrigin(vertical_vector, top_point + vertical_vector.Normalize() * (barSize*50))

rebar = Structure.Rebar.CreateFromRebarShape(doc, sc_38, barType16, element, top_point, horizontal_vector,vertical_vector)
# rebar = sc20_by2points(top_point, get_top_point_of_line(rightLineFront), hwTop_front.GetSurface(), element, barType16, doc, sc_20)
rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(Spacing * 0.00328084, 5,False , True, True) #SetLayoutAsMaximumSpacing(double spacing,double arrayLength,bool barsOnNormalSide,bool includeFirstBar,bool includeLastBar)
shape_accessor = rebar.GetShapeDrivenAccessor()
shape_accessor.UseRebarConstraintsToProduceVaryingBars = True
rebar.LookupParameter("Mark").Set(Headwall_38)
rebar.LookupParameter("B").Set(HeadW1_T - cover*2 - barSize*2)
rebar.LookupParameter("A").Set(barSize*50)

rebarNormalVec = rebar.GetShapeDrivenAccessor().GetDistributionPath().Direction 
handleList = rebar.GetRebarConstraintsManager().GetAllHandles()
conman = rebar.GetRebarConstraintsManager()

for handle in handleList:
    print("Starting handle:" + str(handle.GetHandleName()))

    if handle.GetHandleName() == "Out of Plane Extent":
        tempFace = Plane.CreateByNormalAndOrigin(rebarNormalVec, hwTop_front.Origin + rebarNormalVec.Normalize() * -(- cover - barSize*2))
        constraint = RebarConstraint.CreateConstraintToSurface(handle, tempFace)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()
        print("$$$"*50)
        print("##### Out of Plane Extent - Handle processed #####")

    if handle.GetHandleName() == "Start of Bar":
        constraint = RebarConstraint.CreateConstraintToSurface(handle, offseg2Plane)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()
        print("##### Start of Bar - Handle processed #####")   

    if handle.GetHandleName() == "End of Bar":
        constraint = RebarConstraint.CreateConstraintToSurface(handle, offseg2Plane)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()
        print("##### End of Bar - Handle processed #####") 

    if handle.GetHandleName() == "Bar Segment 1":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == front_face.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(- barSize)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover - barSize
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Bar Segment 1 - Handle processed #####") 

    if handle.GetHandleName() == "Bar Segment 2":
        constraint = RebarConstraint.CreateConstraintToSurface(handle, seg2Plane)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()
        print("##### Bar Segment 2 - Handle processed #####") 

#endregion

#region -  ubars on back head wall
##################################################################################################################################################################
top_point = get_top_point_of_line(leftLineBack)
vertical_vector = backTransverseLine.Direction.Normalize()
horizontal_vector = -XYZ.BasisZ.CrossProduct(vertical_vector)                      

seg2Plane = Plane.CreateByNormalAndOrigin(vertical_vector, top_point)
offseg2Plane = Plane.CreateByNormalAndOrigin(vertical_vector, top_point + vertical_vector.Normalize() * (barSize*50))

rebar = Structure.Rebar.CreateFromRebarShape(doc, sc_38, barType16, element, top_point, horizontal_vector,vertical_vector)
# rebar = sc20_by2points(top_point, get_top_point_of_line(rightLineFront), hwTop_front.GetSurface(), element, barType16, doc, sc_20)
rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(Spacing * 0.00328084, 5,False , True, True) #SetLayoutAsMaximumSpacing(double spacing,double arrayLength,bool barsOnNormalSide,bool includeFirstBar,bool includeLastBar)
shape_accessor = rebar.GetShapeDrivenAccessor()
shape_accessor.UseRebarConstraintsToProduceVaryingBars = True
rebar.LookupParameter("Mark").Set(HeadwallDownStream_38)
rebar.LookupParameter("B").Set(HeadW1_T - cover*2 - barSize*2)
rebar.LookupParameter("A").Set(barSize*50)

rebarNormalVec = rebar.GetShapeDrivenAccessor().GetDistributionPath().Direction 
handleList = rebar.GetRebarConstraintsManager().GetAllHandles()
conman = rebar.GetRebarConstraintsManager()

for handle in handleList:
    print("Starting handle:" + str(handle.GetHandleName()))

    if handle.GetHandleName() == "Out of Plane Extent":
        tempFace = Plane.CreateByNormalAndOrigin(rebarNormalVec, hwTop_back.Origin + rebarNormalVec.Normalize() * -(- cover - barSize*2))
        constraint = RebarConstraint.CreateConstraintToSurface(handle, tempFace)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()
        print("$$$"*50)
        print("##### Out of Plane Extent - Handle processed #####")

    if handle.GetHandleName() == "Start of Bar":
        constraint = RebarConstraint.CreateConstraintToSurface(handle, offseg2Plane)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()
        print("##### Start of Bar - Handle processed #####")   

    if handle.GetHandleName() == "End of Bar":
        constraint = RebarConstraint.CreateConstraintToSurface(handle, offseg2Plane)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()
        print("##### End of Bar - Handle processed #####") 

    if handle.GetHandleName() == "Bar Segment 1":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == back_face.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(- barSize)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover - barSize
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Bar Segment 1 - Handle processed #####") 

    if handle.GetHandleName() == "Bar Segment 2":
        constraint = RebarConstraint.CreateConstraintToSurface(handle, seg2Plane)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()
        print("##### Bar Segment 2 - Handle processed #####") 


##################################################################################################################################################################
# RIGTH SIDE OF BACK HEAD WALL
top_point = get_top_point_of_line(rightLineBack)
vertical_vector = -backTransverseLine.Direction.Normalize()
horizontal_vector = XYZ.BasisZ.CrossProduct(vertical_vector)                      

seg2Plane = Plane.CreateByNormalAndOrigin(vertical_vector, top_point)
offseg2Plane = Plane.CreateByNormalAndOrigin(vertical_vector, top_point + vertical_vector.Normalize() * (barSize*50))

rebar = Structure.Rebar.CreateFromRebarShape(doc, sc_38, barType16, element, top_point, horizontal_vector,vertical_vector)
# rebar = sc20_by2points(top_point, get_top_point_of_line(rightLineFront), hwTop_front.GetSurface(), element, barType16, doc, sc_20)
rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(Spacing * 0.00328084, 5,False , True, True) #SetLayoutAsMaximumSpacing(double spacing,double arrayLength,bool barsOnNormalSide,bool includeFirstBar,bool includeLastBar)
shape_accessor = rebar.GetShapeDrivenAccessor()
shape_accessor.UseRebarConstraintsToProduceVaryingBars = True
rebar.LookupParameter("Mark").Set(HeadwallDownStream_38)
rebar.LookupParameter("B").Set(HeadW1_T - cover*2 - barSize*2)
rebar.LookupParameter("A").Set(barSize*50)

rebarNormalVec = rebar.GetShapeDrivenAccessor().GetDistributionPath().Direction 
handleList = rebar.GetRebarConstraintsManager().GetAllHandles()
conman = rebar.GetRebarConstraintsManager()

for handle in handleList:
    print("Starting handle:" + str(handle.GetHandleName()))

    if handle.GetHandleName() == "Out of Plane Extent":
        tempFace = Plane.CreateByNormalAndOrigin(rebarNormalVec, hwTop_back.Origin + rebarNormalVec.Normalize() * (- cover - barSize*2))
        constraint = RebarConstraint.CreateConstraintToSurface(handle, tempFace)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()
        print("$$$"*50)
        print("##### Out of Plane Extent - Handle processed #####")

    if handle.GetHandleName() == "Start of Bar":
        constraint = RebarConstraint.CreateConstraintToSurface(handle, offseg2Plane)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()
        print("##### Start of Bar - Handle processed #####")   

    if handle.GetHandleName() == "End of Bar":
        constraint = RebarConstraint.CreateConstraintToSurface(handle, offseg2Plane)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()
        print("##### End of Bar - Handle processed #####") 

    if handle.GetHandleName() == "Bar Segment 1":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == back_face.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(- barSize)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover - barSize
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Bar Segment 1 - Handle processed #####") 

    if handle.GetHandleName() == "Bar Segment 2":
        constraint = RebarConstraint.CreateConstraintToSurface(handle, seg2Plane)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()
        print("##### Bar Segment 2 - Handle processed #####") 

#endregion


# supress warnings ####################################################################
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SupressWarnings())
t.SetFailureHandlingOptions(failHandler)

t.Commit()

print("#"*100)
print("#"*100)
print("DONE")
print("#"*100)