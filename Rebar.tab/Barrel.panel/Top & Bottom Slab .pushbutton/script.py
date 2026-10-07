# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import * 
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

element = doc.GetElement(ElementId(Int64(552179)))
# botslab = doc.GetElement(ElementId(Int64(508812)))
# slab = doc.GetElement(ElementId(Int64(502467)))

front_face_Id = 147
top_face_Id = 138
bot_face_Id = 127
back_face_Id = 181
left_face_Id = 132
right_face_Id = 142

botSlabHeight = element.LookupParameter("Flr_T").AsDouble()
topSlabHeight = element.LookupParameter("Roof_T").AsDouble()
Cul_Haunch_Bot_D = element.LookupParameter("Cul_Haunch_Bot_D").AsDouble()

###### MARKS ##################################################################################

TopSlab_Longitudinal_top = "TopSlab Longitudinal top"
TopSlab_Longitudinal_Bot = "TopSlab Longitudinal Bot"
TopSlab_Transverse_top = "TopSlab Transverse top"
TopSlab_Transverse_bot = "TopSlab Transverse bot"

BottomSlab_Longitudinal_top = "BottomSlab Longitudinal top"
BottomSlab_Longitudinal_bot = "BottomSlab Longitudinal bot"
BottomSlab_Transverse_top = "BottomSlab Transverse top"
BottomSlab_Transverse_bot = "BottomSlab Transverse bot"

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
sc_55 = getRebarShapeByName("55")
barType16 = barTypeBySize("Y16") # 16mm rebar
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

def is_line_in_plane(line, plane, tol=1e-9):
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


leftLineFront, rightLineFront, leftLineBack, rightLineBack = collect_corner_lines(curve_candidates, front_face, back_face, left_face, right_face)

t = Transaction(doc, "Create Slab Rebar")
t.Start()

#region -  Top slab longitudinal rebar 20
##################################################################################################################################################################

top_point = get_top_point_of_line(leftLineFront)
vertical_vector = leftLineFront.Direction.Normalize()
horizontal_vector = vertical_vector.CrossProduct(front_face.ComputeNormal(UV(0,0))).Normalize()
rebar = Structure.Rebar.CreateFromRebarShape(doc, sc_20, barType16, element, top_point, horizontal_vector, vertical_vector)
rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(Spacing * 0.00328084, 5,False , True, True) #SetLayoutAsMaximumSpacing(double spacing,double arrayLength,bool barsOnNormalSide,bool includeFirstBar,bool includeLastBar)
shape_accessor = rebar.GetShapeDrivenAccessor()
shape_accessor.UseRebarConstraintsToProduceVaryingBars = True
rebar.LookupParameter("Mark").Set(TopSlab_Longitudinal_top)

#set rebar constraint handles to the curve loop
handleList = rebar.GetRebarConstraintsManager().GetAllHandles()
conman = rebar.GetRebarConstraintsManager()
conSurface = left_face.GetSurface()

#colect all lines that go through bottom_point
rebarVec = None
for c in curve_candidates:
        if (c.GetEndPoint(0).DistanceTo(top_point) < 1e-6 or c.GetEndPoint(1).DistanceTo(top_point) < 1e-6):
            if not is_line_in_plane(c, front_face.GetSurface()):
                rebarVec = c.Direction.Normalize()
                rebarVecZ = abs(c.GetEndPoint(0).Z - c.GetEndPoint(1).Z)
                rebarVecPlanLength = (XYZ(c.GetEndPoint(0).X - c.GetEndPoint(1).X, c.GetEndPoint(0).Y - c.GetEndPoint(1).Y, 0)).GetLength()

horVec = XYZ(horizontal_vector.X, horizontal_vector.Y, 0).Normalize()
norVec = rebarVec.CrossProduct(horVec).Normalize()

# conSurface = Plane.CreateByNormalAndOrigin(conSurface.Normal, conSurface.Origin + conSurface.Normal * cover)
# conSurface3 = Plane.CreateByNormalAndOrigin(conSurface.Normal, conSurface.Origin + conSurface.Normal * (-600*0.00328084  - cover))

# #get the top handle between bar segment 1 and 3
# for handle in handleList:
#     if handle.GetHandleName() == "Bar Segment 1":
#         handle1_Z = handle.GetHandleSurface().Origin.Z
#         for handle in handleList:
#             if handle.GetHandleName() == "Bar Segment 3":
#                 handle3_Z = handle.GetHandleSurface().Origin.Z
#                 if handle1_Z > handle3_Z:
#                     handle1_ontop = True
#                 else:            
#                     handle1_ontop = False


for handle in handleList:
    print("Starting handle:" + str(handle.GetHandleName()))

    if handle.GetHandleName() == "Out of Plane Extent":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)[2]
        conSurface = element.GetGeometryObjectFromReference(constraint.GetTargetHostFaceReference())
        if constraint and constraint.IsToCover():
            new_offset = -cover
            constraint.SetDistanceToTargetCover(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover
            constraint.SetDistanceToTargetHostFace(new_offset)
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Out of Plane Extent - Handle processed #####")

    if handle.GetHandleName() == "Start of Bar":
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
        print("##### Start of Bar - Handle processed #####")   

    if handle.GetHandleName() == "End of Bar":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == left_face.Id:
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
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == front_face.Id:
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
        # constraint = conman.GetCurrentConstraintOnHandle(handle)
        # constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)[0]
        top_faceTL20 = Plane.CreateByNormalAndOrigin(top_face.GetSurface().Normal, top_face.GetSurface().Origin + XYZ(0,0,1) * (- cover))
        # bot_face = Plane.CreateByNormalAndOrigin(norVec, bottom_point + XYZ(0,0,1) * (- botSlabHeight -Cul_Haunch_Bot_D + cover ))

        constraint = RebarConstraint.CreateConstraintToSurface(handle, top_faceTL20)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()
        print("##### Bar Segment 1 - Handle processed #####") 

    # if handle.GetHandleName() == "Bar Segment 3":
    #     top_face = Plane.CreateByNormalAndOrigin(norVec, bottom_point + XYZ(0,0,1) * (cuvertHeight - botSlabHeight -Cul_Haunch_Bot_D - cover + botPointZ_corrector))
    #     bot_face = Plane.CreateByNormalAndOrigin(norVec, bottom_point + XYZ(0,0,1) * (- botSlabHeight -Cul_Haunch_Bot_D + cover + botPointZ_corrector))

    #     if handle1_ontop:
    #         faceTemp3 = bot_face
    #     else:
    #         faceTemp3 = top_face

    #     constraint = RebarConstraint.CreateConstraintToSurface(handle, faceTemp3)
    #     conman.SetPreferredConstraint(constraint)
    #     doc.Regenerate()

    #     if constraint and constraint.IsToCover():
    #         new_offsetBot = -cover
    #         constraint.SetDistanceToTargetCover(new_offsetBot)
    #         # cons = RebarConstraint.CreateConstraintToSurface(handle, plane)
    #         conman.SetPreferredConstraint(constraint)
    #         doc.Regenerate()
    #     elif constraint and constraint.IsToHostFaceOrCover(): 
    #         new_offsetBot = -cover
    #         constraint.SetDistanceToTargetHostFace(new_offsetBot)
    #         # cons = RebarConstraint.CreateConstraintToSurface(handle, plane)
    #         constraint.SetDistanceToTargetHostFace(new_offsetBot)
    #         conman.SetPreferredConstraint(constraint)
    #         doc.Regenerate()
    #     print("##### Bar Segment 3 - Handle processed #####")

    # if handle.GetHandleName() == "Bar Segment 2":
    #     # constraint = conman.GetCurrentConstraintOnHandle(handle)
    #     constraint = RebarConstraint.CreateConstraintToSurface(handle, conSurface)
    #     # constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)[0]
    #     # print("Current constraint type:", constraint)
    #     # print("Is ToCover constraint:",  constraint.IsToCover())
    #     # print("Is ToHostFace constraint:", constraint.IsToHostFaceOrCover())
    #     if constraint and constraint.IsToCover():
    #         new_offset = -cover  
    #         constraint.SetDistanceToTargetCover(new_offset)
    #         # cons = RebarConstraint.CreateConstraintToSurface(handle, plane)
    #         conman.SetPreferredConstraint(constraint)
    #         doc.Regenerate()
    #         # print("Updated:", constraint.GetDistanceToTargetCover())
    #     elif constraint and constraint.IsToHostFaceOrCover(): 
    #         new_offset = -cover  
    #         constraint.SetDistanceToTargetHostFace(new_offset)
    #         # cons = RebarConstraint.CreateConstraintToSurface(handle, plane)
    #         constraint.SetDistanceToTargetHostFace(new_offset)
    #         conman.SetPreferredConstraint(constraint)
    #         doc.Regenerate()
    #     conman.SetPreferredConstraint(constraint)
    #     doc.Regenerate()
    #     print("##### Bar Segment 2 - Handle processed #####")   

##################################################################################################################################################################
#endregion

#region -  Top slab longitudinal rebar 55
##################################################################################################################################################################

top_point = get_top_point_of_line(leftLineFront)
vertical_vector = leftLineFront.Direction.Normalize()
horizontal_vector = vertical_vector.CrossProduct(front_face.ComputeNormal(UV(0,0))).Normalize()

for c in curve_candidates:
    if c.GetEndPoint(0).DistanceTo(top_point) < 1e-6 or c.GetEndPoint(1).DistanceTo(top_point) < 1e-6:
        if not is_line_in_plane(c, front_face.GetSurface()):
            rebarVec = c.Direction.Normalize()

SlopeVec = rebarVec.CrossProduct(horizontal_vector).Normalize()
rebarSlopeVec = SlopeVec.CrossProduct(horizontal_vector).Normalize()
slopePlane = Plane.CreateByThreePoints(top_point, top_point + rebarSlopeVec, top_point + horizontal_vector)

rebar = Structure.Rebar.CreateFromRebarShape(doc, sc_55, barType16, element, top_point+XYZ(0,0,-1)*(topSlabHeight - cover), horizontal_vector, vertical_vector)
rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(Spacing * 0.00328084, 5,False , True, True) #SetLayoutAsMaximumSpacing(double spacing,double arrayLength,bool barsOnNormalSide,bool includeFirstBar,bool includeLastBar)
shape_accessor = rebar.GetShapeDrivenAccessor()
shape_accessor.UseRebarConstraintsToProduceVaryingBars = True
rebar.LookupParameter("Mark").Set(TopSlab_Longitudinal_Bot)

rebar.LookupParameter("B").Set(topSlabHeight - 2*cover)
rebar.LookupParameter("D").Set(topSlabHeight - 2*cover)

#set rebar constraint handles to the curve loop
handleList = rebar.GetRebarConstraintsManager().GetAllHandles()
conman = rebar.GetRebarConstraintsManager()
conSurface = left_face.GetSurface()

#colect all lines that go through bottom_point
rebarVec = None
for c in curve_candidates:
        if (c.GetEndPoint(0).DistanceTo(top_point) < 1e-6 or c.GetEndPoint(1).DistanceTo(top_point) < 1e-6):
            if not is_line_in_plane(c, front_face.GetSurface()):
                rebarVec = c.Direction.Normalize()
                rebarVecZ = abs(c.GetEndPoint(0).Z - c.GetEndPoint(1).Z)
                rebarVecPlanLength = (XYZ(c.GetEndPoint(0).X - c.GetEndPoint(1).X, c.GetEndPoint(0).Y - c.GetEndPoint(1).Y, 0)).GetLength()

# horVec = XYZ(horizontal_vector.X, horizontal_vector.Y, 0).Normalize()
# norVec = rebarVec.CrossProduct(horVec).Normalize()

for handle in handleList:
    print("Starting handle:" + str(handle.GetHandleName()))

    if handle.GetHandleName() == "Out of Plane Extent":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == back_face_Id:
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(0)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover
            constraint.SetDistanceToTargetHostFace(new_offset)
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Out of Plane Extent - Handle processed #####")

    if handle.GetHandleName() == "Start of Bar":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == left_face.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(-50*barSize)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover - barSize*50
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Start of Bar - Handle processed #####")   

    if handle.GetHandleName() == "End of Bar":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == right_face.Id:
                constraint = const
                print("###"*200)
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(-50*barSize)
            conman.SetPreferredConstraint(constraint)
            print("constraint Set")
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover - barSize*50
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### End of Bar - Handle processed #####") 

    if handle.GetHandleName() == "Bar Plane":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == front_face.Id:
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
        top_faceTL55 = Plane.CreateByNormalAndOrigin(slopePlane.Normal, slopePlane.Origin + XYZ(0,0,1) * (  - cover))
        constraint = RebarConstraint.CreateConstraintToSurface(handle, top_faceTL55)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()

        print("##### Bar Segment 1 - Handle processed #####")

    if handle.GetHandleName() == "Bar Segment 3":
        top_faceTL55 = Plane.CreateByNormalAndOrigin(slopePlane.Normal, slopePlane.Origin + XYZ(0,0,1) * ( - topSlabHeight  + cover))
        constraint = RebarConstraint.CreateConstraintToSurface(handle, top_faceTL55)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()

        print("##### Bar Segment 3 - Handle processed #####")

    if handle.GetHandleName() == "Bar Segment 5":
        top_faceTL55 = Plane.CreateByNormalAndOrigin(slopePlane.Normal, slopePlane.Origin + XYZ(0,0,1) * (  - cover))
        constraint = RebarConstraint.CreateConstraintToSurface(handle, top_faceTL55)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()

        print("##### Bar Segment 5 - Handle processed #####")

    if handle.GetHandleName() == "Bar Segment 2":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == left_face.Id:
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

    if handle.GetHandleName() == "Bar Segment 4":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == right_face.Id:
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
        print("##### Bar Segment 4 - Handle processed #####")  

#edit bars in the rebar set by rotating the each bar segment to the top face
top_point_front_left = get_top_point_of_line(leftLineFront)
top_point_front_right = get_top_point_of_line(rightLineFront)
top_point_back_left = get_top_point_of_line(leftLineBack)
top_point_back_right = get_top_point_of_line(rightLineBack)

if top_point_front_left.Z + top_point_front_right.Z < top_point_back_left.Z + top_point_back_right.Z or top_point_front_left.Z + top_point_back_left.Z < top_point_front_right.Z + top_point_back_right.Z:
    rotDir = -1
else:
    rotDir = 1

if top_point_front_left.Z + top_point_front_right.Z < top_point_back_left.Z + top_point_back_right.Z and top_point_front_left.Z + top_point_back_left.Z < top_point_front_right.Z + top_point_back_right.Z:
    rotDir = -1
elif top_point_front_left.Z + top_point_front_right.Z < top_point_back_left.Z + top_point_back_right.Z and top_point_front_left.Z + top_point_back_left.Z > top_point_front_right.Z + top_point_back_right.Z:
    rotDir = 1

rotAxis = Line.CreateBound(top_point_front_left, top_point_back_left).Direction
rotAngleStart = angle_to_horizontal_two_points(top_point_front_left, top_point_front_right)*rotDir
rotAngleEnd = angle_to_horizontal_two_points(top_point_back_left, top_point_back_right)*rotDir


rebar.LookupParameter("B").Set(topSlabHeight - 2*cover)
rebar.LookupParameter("D").Set(topSlabHeight - 2*cover)

for bar in range(rebar.NumberOfBarPositions):
    bar_transform = Transform.CreateRotationAtPoint(rotAxis, rotAngleStart + (rotAngleEnd - rotAngleStart) * bar / (rebar.NumberOfBarPositions - 1), top_point_front_left)
    rebar.MoveBarInSet(bar, bar_transform)




##################################################################################################################################################################
#endregion

#region -  Top slab transverse rebar 55
##################################################################################################################################################################
top_point = get_top_point_of_line(rightLineFront)
vertical_vector = rightLineFront.Direction.Normalize()
horizontal_vector = vertical_vector.CrossProduct(right_face.ComputeNormal(UV(0,0))).Normalize()

for c in curve_candidates:
    if c.GetEndPoint(0).DistanceTo(top_point) < 1e-6 or c.GetEndPoint(1).DistanceTo(top_point) < 1e-6:
        if not is_line_in_plane(c, right_face.GetSurface()):
            rebarVec = c.Direction.Normalize()

SlopeVec = rebarVec.CrossProduct(horizontal_vector).Normalize()
rebarSlopeVec = SlopeVec.CrossProduct(horizontal_vector).Normalize()
slopePlane = Plane.CreateByThreePoints(top_point, top_point + rebarSlopeVec, top_point + horizontal_vector)

rebar = Structure.Rebar.CreateFromRebarShape(doc, sc_55, barType16, element, top_point+XYZ(0,0,1)*( - topSlabHeight - cover), -horizontal_vector, -vertical_vector)
rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(Spacing * 0.00328084, 5,True , True, True) #SetLayoutAsMaximumSpacing(double spacing,double arrayLength,bool barsOnNormalSide,bool includeFirstBar,bool includeLastBar)
shape_accessor = rebar.GetShapeDrivenAccessor()
shape_accessor.UseRebarConstraintsToProduceVaryingBars = True
rebar.LookupParameter("Mark").Set(TopSlab_Transverse_bot)

rebar.LookupParameter("A").Set(barSize*50)
rebar.LookupParameter("B").Set(botSlabHeight - 2*cover - 2*barSize)
rebar.LookupParameter("D").Set(botSlabHeight - 2*cover - 2*barSize)
rebar.LookupParameter("E").Set(barSize*50)

#set rebar constraint handles to the curve loop
handleList = rebar.GetRebarConstraintsManager().GetAllHandles()
conman = rebar.GetRebarConstraintsManager()
conSurface = left_face.GetSurface()

for handle in handleList:
    print("Starting handle:" + str(handle.GetHandleName()))

    if handle.GetHandleName() == "Out of Plane Extent":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == left_face_Id:
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(0)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover
            constraint.SetDistanceToTargetHostFace(new_offset)
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Out of Plane Extent - Handle processed #####")

    if handle.GetHandleName() == "Start of Bar":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == front_face.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(-50*barSize)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover - barSize*50
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Start of Bar - Handle processed #####")   

    if handle.GetHandleName() == "End of Bar":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == back_face.Id:
                constraint = const
                print("###"*200)
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(-50*barSize)
            conman.SetPreferredConstraint(constraint)
            print("constraint Set")
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover - barSize*50
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### End of Bar - Handle processed #####") 

    if handle.GetHandleName() == "Bar Plane":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == right_face.Id:
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
        bot_faceTL55 = Plane.CreateByNormalAndOrigin(slopePlane.Normal, slopePlane.Origin + XYZ(0,0,1) * ( - cover -barSize))
        constraint = RebarConstraint.CreateConstraintToSurface(handle, bot_faceTL55)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()

        print("##### Bar Segment 1 - Handle processed #####")

    if handle.GetHandleName() == "Bar Segment 3":
        bot_faceTL55 = Plane.CreateByNormalAndOrigin(slopePlane.Normal, slopePlane.Origin + XYZ(0,0,1) * ( - topSlabHeight   + cover + barSize))
        constraint = RebarConstraint.CreateConstraintToSurface(handle, bot_faceTL55)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()

        print("##### Bar Segment 3 - Handle processed #####")

    if handle.GetHandleName() == "Bar Segment 5":
        bot_faceTL55 = Plane.CreateByNormalAndOrigin(slopePlane.Normal, slopePlane.Origin + XYZ(0,0,1) * ( - cover - barSize))
        constraint = RebarConstraint.CreateConstraintToSurface(handle, bot_faceTL55)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()

        print("##### Bar Segment 5 - Handle processed #####")

    if handle.GetHandleName() == "Bar Segment 2":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == front_face.Id:
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

    if handle.GetHandleName() == "Bar Segment 4":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == back_face.Id:
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
        print("##### Bar Segment 4 - Handle processed #####")  

#edit bars in the rebar set by rotating the each bar segment to the top face
top_point_front_left = get_top_point_of_line(leftLineFront)
top_point_front_right = get_top_point_of_line(rightLineFront)
top_point_back_left = get_top_point_of_line(leftLineBack)
top_point_back_right = get_top_point_of_line(rightLineBack)

if top_point_front_left.Z + top_point_front_right.Z < top_point_back_right.Z + top_point_back_left .Z:
    rotDir = 1
else:
    rotDir = -1

rotAxis = Line.CreateBound(top_point_front_left+ XYZ(0,0,1) * (-topSlabHeight), top_point_front_right+ XYZ(0,0,1) * (-topSlabHeight)).Direction
rotAngleStart = angle_to_horizontal_two_points(top_point_front_right, top_point_back_right)*rotDir
rotAngleEnd = angle_to_horizontal_two_points(top_point_front_left, top_point_back_left)*rotDir

for bar in range(rebar.NumberOfBarPositions):
    bar_transform = Transform.CreateRotationAtPoint(rotAxis, rotAngleStart + (rotAngleEnd - rotAngleStart) * bar / (rebar.NumberOfBarPositions - 1), top_point_front_left)
    rebar.MoveBarInSet(bar, bar_transform)

##################################################################################################################################################################
#endregion

#region -  Top slab transverse rebar 20
##################################################################################################################################################################
top_point = get_top_point_of_line(rightLineFront)

rebarVec = None
for c in curve_candidates:
        if (c.GetEndPoint(0).DistanceTo(top_point) < 1e-6 or c.GetEndPoint(1).DistanceTo(top_point) < 1e-6):
            if not is_line_in_plane(c, right_face.GetSurface()) and not is_line_vertical(c):
                rebarVec = c.Direction.Normalize()
                rebarVecZ = abs(c.GetEndPoint(0).Z - c.GetEndPoint(1).Z)
                rebarVecPlanLength = (XYZ(c.GetEndPoint(0).X - c.GetEndPoint(1).X, c.GetEndPoint(0).Y - c.GetEndPoint(1).Y, 0)).GetLength()

for c in curve_candidates:
        if (c.GetEndPoint(0).DistanceTo(top_point) < 1e-6 or c.GetEndPoint(1).DistanceTo(top_point) < 1e-6):
            if not is_line_in_plane(c, front_face.GetSurface()):
                horizontal_vector = -c.Direction.Normalize()
                rebarLine = c

horizontal_vector = XYZ(horizontal_vector.X, horizontal_vector.Y, 0).Normalize()
rebarLine = Line.CreateBound(rebarLine.GetEndPoint(0), rebarLine.GetEndPoint(1))
#IList of rebarLine
rebarCurve_list = List[Curve]([rebarLine])
rebar = Structure.Rebar.CreateFromCurvesAndShape(doc, sc_20, barType16, element, right_face.GetSurface().Normal, rebarCurve_list, BarTerminationsData(doc))
rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(Spacing * 0.00328084, 5,False , True, True) #SetLayoutAsMaximumSpacing(double spacing,double arrayLength,bool barsOnNormalSide,bool includeFirstBar,bool includeLastBar)
shape_accessor = rebar.GetShapeDrivenAccessor()
shape_accessor.UseRebarConstraintsToProduceVaryingBars = True
rebar.LookupParameter("Mark").Set(TopSlab_Transverse_top)

handleList = rebar.GetRebarConstraintsManager().GetAllHandles()
conman = rebar.GetRebarConstraintsManager()

for handle in handleList:
    # print("Starting handle:" + str(handle.GetHandleName()))

    if handle.GetHandleName() == "Out of Plane Extent":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        # conSurface = element.GetGeometryObjectFromReference(constraint.GetTargetHostFaceReference())
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == left_face.Id:
                constraint = const
                break
        if constraint and constraint.IsToCover():
            new_offset = -cover
            constraint.SetDistanceToTargetCover(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        # print("##### Out of Plane Extent - Handle processed #####")

    if handle.GetHandleName() == "Start of Bar":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == back_face_Id:
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
            if conSur.Id == front_face_Id:
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
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == right_face.Id:
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
            if conSur.Id == top_face.Id:
                constraint = const
                break

        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover( - barSize)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset =  - barSize
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Bar Segment 1 - Handle processed #####") 


        # draw_plane_in_view(bot_face.GetSurface())
        # bot_faceBS1_BT20 = Plane.CreateByNormalAndOrigin(bot_face.GetSurface().Normal, bot_face.GetSurface().Origin + XYZ(0,0,1) * ( botSlabHeight - cover ))
        # constraint = RebarConstraint.CreateConstraintToSurface(handle, bot_faceBS1_BT20)
        # conman.SetPreferredConstraint(constraint)
        # doc.Regenerate()
        # print("##### Bar Segment 1 - Handle processed #####") 

#endregion

#region -  Bot slab longitudinal rebar 20
##################################################################################################################################################################
bot_point = get_bottom_point_of_line(leftLineFront)
vertical_vector = leftLineFront.Direction.Normalize()
horizontal_vector = vertical_vector.CrossProduct(front_face.ComputeNormal(UV(0,0))).Normalize()
rebar = Structure.Rebar.CreateFromRebarShape(doc, sc_20, barType16, element, bot_point, horizontal_vector, vertical_vector)
rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(Spacing * 0.00328084, 5,False , True, True) #SetLayoutAsMaximumSpacing(double spacing,double arrayLength,bool barsOnNormalSide,bool includeFirstBar,bool includeLastBar)
shape_accessor = rebar.GetShapeDrivenAccessor()
shape_accessor.UseRebarConstraintsToProduceVaryingBars = True
rebar.LookupParameter("Mark").Set(BottomSlab_Longitudinal_top)

#set rebar constraint handles to the curve loop
handleList = rebar.GetRebarConstraintsManager().GetAllHandles()
conman = rebar.GetRebarConstraintsManager()
conSurface = left_face.GetSurface()

#colect all lines that go through bottom_point
rebarVec = None
for c in curve_candidates:
        if (c.GetEndPoint(0).DistanceTo(bot_point) < 1e-6 or c.GetEndPoint(1).DistanceTo(bot_point) < 1e-6):
            if not is_line_in_plane(c, front_face.GetSurface()):
                rebarVec = c.Direction.Normalize()
                rebarVecZ = abs(c.GetEndPoint(0).Z - c.GetEndPoint(1).Z)
                rebarVecPlanLength = (XYZ(c.GetEndPoint(0).X - c.GetEndPoint(1).X, c.GetEndPoint(0).Y - c.GetEndPoint(1).Y, 0)).GetLength()

# horVec = XYZ(horizontal_vector.X, horizontal_vector.Y, 0).Normalize()
# norVec = rebarVec.CrossProduct(horVec).Normalize()

for handle in handleList:
    print("Starting handle:" + str(handle.GetHandleName()))

    if handle.GetHandleName() == "Out of Plane Extent":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)[2]
        conSurface = element.GetGeometryObjectFromReference(constraint.GetTargetHostFaceReference())
        if constraint and constraint.IsToCover():
            new_offset = -cover
            constraint.SetDistanceToTargetCover(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover
            constraint.SetDistanceToTargetHostFace(new_offset)
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Out of Plane Extent - Handle processed #####")

    if handle.GetHandleName() == "Start of Bar":
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
        print("##### Start of Bar - Handle processed #####")   

    if handle.GetHandleName() == "End of Bar":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == left_face.Id:
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
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == front_face.Id:
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
        # constraint = conman.GetCurrentConstraintOnHandle(handle)
        # constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)[0]
        bot_faceBS20 = Plane.CreateByNormalAndOrigin(bot_face.GetSurface().Normal, bot_face.GetSurface().Origin + XYZ(0,0,1) * (-cover + botSlabHeight))
        # bot_face = Plane.CreateByNormalAndOrigin(norVec, bottom_point + XYZ(0,0,1) * (- botSlabHeight -Cul_Haunch_Bot_D + cover ))

        constraint = RebarConstraint.CreateConstraintToSurface(handle, bot_faceBS20)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()
        print("##### Bar Segment 1 - Handle processed #####") 
#endregion

#region -  Bot slab longitudinal rebar 55 
#################################################################################################################################################################
top_point = get_bottom_point_of_line(leftLineFront)
vertical_vector = leftLineFront.Direction.Normalize()
horizontal_vector = vertical_vector.CrossProduct(front_face.ComputeNormal(UV(0,0))).Normalize()

for c in curve_candidates:
    if c.GetEndPoint(0).DistanceTo(top_point) < 1e-6 or c.GetEndPoint(1).DistanceTo(top_point) < 1e-6:
        if not is_line_in_plane(c, front_face.GetSurface()):
            rebarVec = c.Direction.Normalize()

SlopeVec = rebarVec.CrossProduct(horizontal_vector).Normalize()
rebarSlopeVec = SlopeVec.CrossProduct(horizontal_vector).Normalize()
slopePlane = Plane.CreateByThreePoints(top_point, top_point + rebarSlopeVec, top_point + horizontal_vector)

rebar = Structure.Rebar.CreateFromRebarShape(doc, sc_55, barType16, element, top_point+XYZ(0,0,-1)*(topSlabHeight - cover), horizontal_vector, vertical_vector)
rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(Spacing * 0.00328084, 5,False , True, True) #SetLayoutAsMaximumSpacing(double spacing,double arrayLength,bool barsOnNormalSide,bool includeFirstBar,bool includeLastBar)
shape_accessor = rebar.GetShapeDrivenAccessor()
shape_accessor.UseRebarConstraintsToProduceVaryingBars = True
rebar.LookupParameter("Mark").Set(BottomSlab_Longitudinal_bot)

rebar.LookupParameter("B").Set(topSlabHeight - 2*cover)
rebar.LookupParameter("D").Set(topSlabHeight - 2*cover)

#set rebar constraint handles to the curve loop
handleList = rebar.GetRebarConstraintsManager().GetAllHandles()
conman = rebar.GetRebarConstraintsManager()
conSurface = left_face.GetSurface()

#colect all lines that go through bottom_point
rebarVec = None
for c in curve_candidates:
        if (c.GetEndPoint(0).DistanceTo(top_point) < 1e-6 or c.GetEndPoint(1).DistanceTo(top_point) < 1e-6):
            if not is_line_in_plane(c, front_face.GetSurface()):
                rebarVec = c.Direction.Normalize()
                rebarVecZ = abs(c.GetEndPoint(0).Z - c.GetEndPoint(1).Z)
                rebarVecPlanLength = (XYZ(c.GetEndPoint(0).X - c.GetEndPoint(1).X, c.GetEndPoint(0).Y - c.GetEndPoint(1).Y, 0)).GetLength()

# horVec = XYZ(horizontal_vector.X, horizontal_vector.Y, 0).Normalize()
# norVec = rebarVec.CrossProduct(horVec).Normalize()

for handle in handleList:
    print("Starting handle:" + str(handle.GetHandleName()))

    if handle.GetHandleName() == "Out of Plane Extent":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == back_face_Id:
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(0)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover
            constraint.SetDistanceToTargetHostFace(new_offset)
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Out of Plane Extent - Handle processed #####")

    if handle.GetHandleName() == "Start of Bar":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == left_face.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(-50*barSize)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover - barSize*50
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Start of Bar - Handle processed #####")   

    if handle.GetHandleName() == "End of Bar":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == right_face.Id:
                constraint = const
                print("###"*200)
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(-50*barSize)
            conman.SetPreferredConstraint(constraint)
            print("constraint Set")
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover - barSize*50
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### End of Bar - Handle processed #####") 

    if handle.GetHandleName() == "Bar Plane":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == front_face.Id:
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
        bot_faceTL55 = Plane.CreateByNormalAndOrigin(slopePlane.Normal, slopePlane.Origin + XYZ(0,0,1) * ( + botSlabHeight - cover))
        constraint = RebarConstraint.CreateConstraintToSurface(handle, bot_faceTL55)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()

        print("##### Bar Segment 1 - Handle processed #####")

    if handle.GetHandleName() == "Bar Segment 3":
        bot_faceTL55 = Plane.CreateByNormalAndOrigin(slopePlane.Normal, slopePlane.Origin + XYZ(0,0,1) * (   + cover))
        constraint = RebarConstraint.CreateConstraintToSurface(handle, bot_faceTL55)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()

        print("##### Bar Segment 3 - Handle processed #####")

    if handle.GetHandleName() == "Bar Segment 5":
        bot_faceTL55 = Plane.CreateByNormalAndOrigin(slopePlane.Normal, slopePlane.Origin + XYZ(0,0,1) * ( + botSlabHeight - cover))
        constraint = RebarConstraint.CreateConstraintToSurface(handle, bot_faceTL55)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()

        print("##### Bar Segment 5 - Handle processed #####")

    if handle.GetHandleName() == "Bar Segment 2":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == left_face.Id:
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

    if handle.GetHandleName() == "Bar Segment 4":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == right_face.Id:
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
        print("##### Bar Segment 4 - Handle processed #####")  

#edit bars in the rebar set by rotating the each bar segment to the top face
bottom_point_front_left = get_bottom_point_of_line(leftLineFront)
bottom_point_front_right = get_bottom_point_of_line(rightLineFront)
bottom_point_back_left = get_bottom_point_of_line(leftLineBack)
bottom_point_back_right = get_bottom_point_of_line(rightLineBack)

if top_point_front_left.Z + top_point_front_right.Z < top_point_back_left.Z + top_point_back_right.Z or top_point_front_left.Z + top_point_back_left.Z < top_point_front_right.Z + top_point_back_right.Z:
    rotDir = -1
else:
    rotDir = 1

if top_point_front_left.Z + top_point_front_right.Z < top_point_back_left.Z + top_point_back_right.Z and top_point_front_left.Z + top_point_back_left.Z < top_point_front_right.Z + top_point_back_right.Z:
    rotDir = -1
elif top_point_front_left.Z + top_point_front_right.Z < top_point_back_left.Z + top_point_back_right.Z and top_point_front_left.Z + top_point_back_left.Z > top_point_front_right.Z + top_point_back_right.Z:
    rotDir = 1


rotAxis = Line.CreateBound(bottom_point_front_left, bottom_point_back_left).Direction
rotAngleStart = angle_to_horizontal_two_points(bottom_point_front_left, bottom_point_front_right)*rotDir
rotAngleEnd = angle_to_horizontal_two_points(bottom_point_back_left, bottom_point_back_right)*rotDir


rebar.LookupParameter("B").Set(botSlabHeight - 2*cover)
rebar.LookupParameter("D").Set(botSlabHeight - 2*cover)

for bar in range(rebar.NumberOfBarPositions):
    bar_transform = Transform.CreateRotationAtPoint(rotAxis, rotAngleStart + (rotAngleEnd - rotAngleStart) * bar / (rebar.NumberOfBarPositions - 1), bottom_point_front_left)
    rebar.MoveBarInSet(bar, bar_transform)

##################################################################################################################################################################
#endregion

#region -  Bot slab transverse rebar 20 
##################################################################################################################################################################
bot_point = get_bottom_point_of_line(rightLineFront)

rebarVec = None
for c in curve_candidates:
        if (c.GetEndPoint(0).DistanceTo(bot_point) < 1e-6 or c.GetEndPoint(1).DistanceTo(bot_point) < 1e-6):
            if not is_line_in_plane(c, right_face.GetSurface()) and not is_line_vertical(c):
                rebarVec = c.Direction.Normalize()
                rebarVecZ = abs(c.GetEndPoint(0).Z - c.GetEndPoint(1).Z)
                rebarVecPlanLength = (XYZ(c.GetEndPoint(0).X - c.GetEndPoint(1).X, c.GetEndPoint(0).Y - c.GetEndPoint(1).Y, 0)).GetLength()

for c in curve_candidates:
        if (c.GetEndPoint(0).DistanceTo(bot_point) < 1e-6 or c.GetEndPoint(1).DistanceTo(bot_point) < 1e-6):
            if not is_line_in_plane(c, front_face.GetSurface()):
                horizontal_vector = -c.Direction.Normalize()
                rebarLine = c

horizontal_vector = XYZ(horizontal_vector.X, horizontal_vector.Y, 0).Normalize()
rebarLine = Line.CreateBound(rebarLine.GetEndPoint(0), rebarLine.GetEndPoint(1))
#IList of rebarLine
rebarCurve_list = List[Curve]([rebarLine])
rebar = Structure.Rebar.CreateFromCurvesAndShape(doc, sc_20, barType16, element, right_face.GetSurface().Normal, rebarCurve_list, BarTerminationsData(doc))
rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(Spacing * 0.00328084, 5,False , True, True) #SetLayoutAsMaximumSpacing(double spacing,double arrayLength,bool barsOnNormalSide,bool includeFirstBar,bool includeLastBar)
shape_accessor = rebar.GetShapeDrivenAccessor()
shape_accessor.UseRebarConstraintsToProduceVaryingBars = True
rebar.LookupParameter("Mark").Set(BottomSlab_Transverse_top)

handleList = rebar.GetRebarConstraintsManager().GetAllHandles()
conman = rebar.GetRebarConstraintsManager()

for handle in handleList:
    print("Starting handle:" + str(handle.GetHandleName()))

    if handle.GetHandleName() == "Out of Plane Extent":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        # conSurface = element.GetGeometryObjectFromReference(constraint.GetTargetHostFaceReference())
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == left_face.Id:
                constraint = const
                break
        if constraint and constraint.IsToCover():
            new_offset = -cover
            constraint.SetDistanceToTargetCover(new_offset)
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
            if conSur.Id == back_face_Id:
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
            if conSur.Id == front_face_Id:
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
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == right_face.Id:
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
            if conSur.Id == bot_face.Id:
                constraint = const
                break

        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(2*cover + 1.5*barSize - botSlabHeight)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = cover - botSlabHeight
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Bar Segment 1 - Handle processed #####") 


        # draw_plane_in_view(bot_face.GetSurface())
        # bot_faceBS1_BT20 = Plane.CreateByNormalAndOrigin(bot_face.GetSurface().Normal, bot_face.GetSurface().Origin + XYZ(0,0,1) * ( botSlabHeight - cover ))
        # constraint = RebarConstraint.CreateConstraintToSurface(handle, bot_faceBS1_BT20)
        # conman.SetPreferredConstraint(constraint)
        # doc.Regenerate()
        # print("##### Bar Segment 1 - Handle processed #####") 


##################################################################################################################################################################
#endregion

#region -  Bot slab transverse rebar 55 
##################################################################################################################################################################
bot_point = get_bottom_point_of_line(rightLineFront)
vertical_vector = rightLineFront.Direction.Normalize()
horizontal_vector = vertical_vector.CrossProduct(right_face.ComputeNormal(UV(0,0))).Normalize()

for c in curve_candidates:
    if c.GetEndPoint(0).DistanceTo(bot_point) < 1e-6 or c.GetEndPoint(1).DistanceTo(bot_point) < 1e-6:
        if not is_line_in_plane(c, right_face.GetSurface()):
            rebarVec = c.Direction.Normalize()

SlopeVec = rebarVec.CrossProduct(horizontal_vector).Normalize()
rebarSlopeVec = SlopeVec.CrossProduct(horizontal_vector).Normalize()
slopePlane = Plane.CreateByThreePoints(bot_point, bot_point + rebarSlopeVec, bot_point + horizontal_vector)

rebar = Structure.Rebar.CreateFromRebarShape(doc, sc_55, barType16, element, bot_point+XYZ(0,0,1)*(  cover), -horizontal_vector, -vertical_vector)
rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(Spacing * 0.00328084, 5,True , True, True) #SetLayoutAsMaximumSpacing(double spacing,double arrayLength,bool barsOnNormalSide,bool includeFirstBar,bool includeLastBar)
shape_accessor = rebar.GetShapeDrivenAccessor()
shape_accessor.UseRebarConstraintsToProduceVaryingBars = True
rebar.LookupParameter("Mark").Set(BottomSlab_Transverse_bot)

rebar.LookupParameter("B").Set(botSlabHeight - 2*cover - 2*barSize)
rebar.LookupParameter("D").Set(botSlabHeight - 2*cover - 2*barSize)

#set rebar constraint handles to the curve loop
handleList = rebar.GetRebarConstraintsManager().GetAllHandles()
conman = rebar.GetRebarConstraintsManager()
conSurface = left_face.GetSurface()

for handle in handleList:
    print("Starting handle:" + str(handle.GetHandleName()))

    if handle.GetHandleName() == "Out of Plane Extent":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == left_face_Id:
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(0)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover
            constraint.SetDistanceToTargetHostFace(new_offset)
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Out of Plane Extent - Handle processed #####")

    if handle.GetHandleName() == "Start of Bar":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == front_face.Id :
                constraint = const
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(-50*barSize)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover - barSize*50
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### Start of Bar - Handle processed #####")   

    if handle.GetHandleName() == "End of Bar":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == back_face.Id:
                constraint = const
                print("###"*200)
                break
        if constraint and constraint.IsToCover():
            constraint.SetDistanceToTargetCover(-50*barSize)
            conman.SetPreferredConstraint(constraint)
            print("constraint Set")
            doc.Regenerate()
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -cover - barSize*50
            constraint.SetDistanceToTargetHostFace(new_offset)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
        print("##### End of Bar - Handle processed #####") 

    if handle.GetHandleName() == "Bar Plane":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == right_face.Id:
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
        bot_faceTL55 = Plane.CreateByNormalAndOrigin(slopePlane.Normal, slopePlane.Origin + XYZ(0,0,1) * ( + botSlabHeight - cover - barSize))
        constraint = RebarConstraint.CreateConstraintToSurface(handle, bot_faceTL55)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()

        print("##### Bar Segment 1 - Handle processed #####")

    if handle.GetHandleName() == "Bar Segment 3":
        bot_faceTL55 = Plane.CreateByNormalAndOrigin(slopePlane.Normal, slopePlane.Origin + XYZ(0,0,1) * (   + cover + barSize))
        constraint = RebarConstraint.CreateConstraintToSurface(handle, bot_faceTL55)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()

        print("##### Bar Segment 3 - Handle processed #####")

    if handle.GetHandleName() == "Bar Segment 5":
        bot_faceTL55 = Plane.CreateByNormalAndOrigin(slopePlane.Normal, slopePlane.Origin + XYZ(0,0,1) * ( + botSlabHeight - cover - barSize))
        constraint = RebarConstraint.CreateConstraintToSurface(handle, bot_faceTL55)
        conman.SetPreferredConstraint(constraint)
        doc.Regenerate()

        print("##### Bar Segment 5 - Handle processed #####")

    if handle.GetHandleName() == "Bar Segment 2":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == front_face.Id:
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

    if handle.GetHandleName() == "Bar Segment 4":
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
        for const in constraint:
            conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
            if conSur.Id == back_face.Id:
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
        print("##### Bar Segment 4 - Handle processed #####")  

#edit bars in the rebar set by rotating the each bar segment to the top face
bottom_point_front_left = get_bottom_point_of_line(leftLineFront)
bottom_point_front_right = get_bottom_point_of_line(rightLineFront)
bottom_point_back_left = get_bottom_point_of_line(leftLineBack)
bottom_point_back_right = get_bottom_point_of_line(rightLineBack)

if bottom_point_front_left.Z + bottom_point_front_right.Z < bottom_point_back_right.Z + bottom_point_back_left .Z:
    rotDir = 1
else:
    rotDir = -1

rotAxis = Line.CreateBound(bottom_point_front_left, bottom_point_front_right).Direction
rotAngleStart = angle_to_horizontal_two_points(bottom_point_front_right, bottom_point_back_right)*rotDir
rotAngleEnd = angle_to_horizontal_two_points(bottom_point_front_left, bottom_point_back_left)*rotDir



for bar in range(rebar.NumberOfBarPositions):
    bar_transform = Transform.CreateRotationAtPoint(rotAxis, rotAngleStart + (rotAngleEnd - rotAngleStart) * bar / (rebar.NumberOfBarPositions - 1), bottom_point_front_left)
    rebar.MoveBarInSet(bar, bar_transform)

##################################################################################################################################################################
#endregion

print("#"*100)


# supress warnings ####################################################################
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SupressWarnings())
t.SetFailureHandlingOptions(failHandler)

t.Commit()

print("#"*100)
print("#"*100)
print("DONE")
print("#"*100)