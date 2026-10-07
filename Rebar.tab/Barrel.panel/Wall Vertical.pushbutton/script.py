# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import * 
import Functions as func
from pyrevit import revit, forms, script
import clr
clr.AddReference("System")
from System import Int64
from System.Collections.Generic import List
import math
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
view = doc.ActiveView

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
ClipSpacing_X = 1000 #mm
ClipSpacing_Y = 1000 #mm


element = doc.GetElement(ElementId(Int64(552179)))

front_face_Id = 147
top_face_Id = 138
bot_face_Id = 127
back_face_Id = 181
left_face_Id = 132
right_face_Id = 142

botSlabHeight = element.LookupParameter("Flr_T").AsDouble()
topSlabHeight = element.LookupParameter("Roof_T").AsDouble()

Cul_Haunch_Top_D = element.LookupParameter("Cul_Haunch_Top_D").AsDouble()
Cul_Haunch_Top_W = element.LookupParameter("Cul_Haunch_Top_W").AsDouble()

Cul_Haunch_Bot_D = element.LookupParameter("Cul_Haunch_Bot_D").AsDouble()
Cul_Haunch_Bot_W = element.LookupParameter("Cul_Haunch_Bot_W").AsDouble()

Cell_Wall_T_Ext = element.LookupParameter("Cell_Wall_T_Ext").AsDouble()
Cell_Wall_T_Int = element.LookupParameter("Cell_Wall_T_Int").AsDouble()

Cell_H = element.LookupParameter("Cell_H").AsDouble()

skew = element.LookupParameter("Inlet_Skew").AsDouble()

###### MARKS ##################################################################################

Wall_Horizontal_U_bar_Ext = "Wall Horizontal U-bar Ext"
Wall_Horizontal_U_bar_Int = "Wall Horizontal U-bar Int"
Wall_Horizontal = "Wall Horizontal"
Wall_Vertical = "Wall Vertical"
Wall_Clips_external = "Wall Clips External"
Wall_Clips_internal = "Wall Clips Internal"

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
sc_38 = getRebarShapeByName("38")
sc_20 = getRebarShapeByName("20")
sc_85 = getRebarShapeByName("85")
barType16 = barTypeBySize("Y16") # 16mm rebar
barSize = 16 * 0.00328084 # 16mm rebar

barType10 = barTypeBySize("Y10") # 10mm rebar   
wallClipBarSize = 10 * 0.00328084 # 10mm rebar

################################################################################################

options = Options()
geometry = element.get_Geometry(options)
curve_candidates = []
face_candidates = []

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

def get_bottom_point_of_line(line):
    if line.IsBound:
        p1 = line.GetEndPoint(0)
        p2 = line.GetEndPoint(1)
        return p1 if p1.Z < p2.Z else p2
    else:
        return line.Origin

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

def line_plane_intersection(line, plane, tol=1e-6):
    n = plane.Normal.Normalize()

    if line.IsBound:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        direction = (p1 - p0).Normalize()
        line_length = p0.DistanceTo(p1)
    else:
        p0 = line.Origin
        direction = line.Direction.Normalize()
        line_length = None

    denom = n.DotProduct(direction)

    # Parallel (or almost parallel) -> no unique intersection
    if abs(denom) < tol:
        return None

    t = n.DotProduct(plane.Origin - p0) / denom
    intersection = p0 + t * direction

    # If bounded line segment, ensure intersection lies on segment
    if line.IsBound:
        if t < -tol or t > line_length + tol:
            return None

    return intersection

def point_plane_distance(point, plane):
    n = plane.Normal.Normalize()
    signed = n.DotProduct(point - plane.Origin)
    return abs(signed)  # unsigned distance

def get_distance_from_line_to_plane(line, plane):
    if line.IsBound:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        return min(point_plane_distance(p0, plane), point_plane_distance(p1, plane))
    else:
        return point_plane_distance(line.Origin, plane)

def is_point_in_plane(plane, point, tol=1e-9):
    n = plane.Normal.Normalize()
    signed_dist = n.DotProduct(point - plane.Origin)
    return abs(signed_dist) <= tol

for gobj in geometry:
    collect_curves_from_geo(gobj)
    if isinstance(gobj, GeometryInstance):
        for iobj in gobj.GetInstanceGeometry():
            collect_curves_from_geo(iobj)
            collect_faces_from_geo(iobj)

print("Family geometry properties:")
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




t = Transaction(doc, "Create Vertical Rebar")
t.Start()



rebarAnchorLine = []
for c in curve_candidates:
    if is_line_in_plane(c, front_face.GetSurface()):
        if abs(c.Direction.Normalize().Z) > 0.9:
            rebarAnchorLine.append(c)

print("Number of rebar anchor lines:", len(rebarAnchorLine))
print("#"*150)

# remove lines that are not on the top of the element, but head wall lines in the front face plane
maxZ = max(max(line.GetEndPoint(0).Z, line.GetEndPoint(1).Z) for line in rebarAnchorLine)
# print(maxZ)
for line in rebarAnchorLine:
    if line.GetEndPoint(0).Z== maxZ or line.GetEndPoint(1).Z == maxZ:
        rebarAnchorLine.remove(line)
#remove second line
maxZ = max(max(line.GetEndPoint(0).Z, line.GetEndPoint(1).Z) for line in rebarAnchorLine)
cuvertHeight = max(line.Length for line in rebarAnchorLine)
# print("Cuvet height:", cuvertHeight)
# print(maxZ)
for line in rebarAnchorLine:
    if line.GetEndPoint(0).Z== maxZ or line.GetEndPoint(1).Z == maxZ:
        rebarAnchorLine.remove(line)

print("Number of rebar anchor lines after removing head wall lines:", len(rebarAnchorLine))
print("#"*150)

##########################################################################################################################################################################################################################################################
#region -  Vertical rebar
##########################################################################################################################################################################################################################################################

for line in rebarAnchorLine:
    bottom_point = get_bottom_point_of_line(line)
    vertical_vector = line.Direction.Normalize()
    horizontal_vector = vertical_vector.CrossProduct(front_face.ComputeNormal(UV(0,0))).Normalize()
    rebar = Structure.Rebar.CreateFromRebarShape(doc, sc_38, barType16, element, bottom_point, vertical_vector, horizontal_vector)
    rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(Spacing * 0.00328084, 5,True , True, True) #SetLayoutAsMaximumSpacing(double spacing,double arrayLength,bool barsOnNormalSide,bool includeFirstBar,bool includeLastBar)
    shape_accessor = rebar.GetShapeDrivenAccessor()
    shape_accessor.UseRebarConstraintsToProduceVaryingBars = True
# to make sure the barSegment1 and barSegment3 connect to the correct element(slab or botslab) we check which handle is on top 
    for f in face_candidates:
        if is_line_in_plane(line, f.GetSurface()) and f.Id != front_face.Id:
            reinforcement_face = f.GetSurface()

# to corect the incorect sign on the cover offset of the outermost planes (abutments not piers) we get the lengt of the line and identify the line as an outher line if it is longer that the rest of the lines
    min_line_length = min(c.Length for c in rebarAnchorLine)
    if line.Length > min_line_length*1.01: # if the line is more than 1% longer than the shortest line, we consider it as an outer line
        cover_dir_corector = -1
        botPointZ_corrector = botSlabHeight + Cul_Haunch_Bot_D
    else:
        cover_dir_corector = 1
        botPointZ_corrector = 0

#set rebar constraint handles to the curve loop
    handleList = rebar.GetRebarConstraintsManager().GetAllHandles()
    conman = rebar.GetRebarConstraintsManager()
    conSurface = reinforcement_face

#colect all lines that go through bottom_point
    rebarVec = None
    for c in curve_candidates:
            if (c.GetEndPoint(0).DistanceTo(bottom_point) < 1e-6 or c.GetEndPoint(1).DistanceTo(bottom_point) < 1e-6):
                if not is_line_in_plane(c, front_face.GetSurface()):
                    rebarVec = c.Direction.Normalize()
                    rebarVecZ = abs(c.GetEndPoint(0).Z - c.GetEndPoint(1).Z)
                    rebarVecPlanLength = (XYZ(c.GetEndPoint(0).X - c.GetEndPoint(1).X, c.GetEndPoint(0).Y - c.GetEndPoint(1).Y, 0)).GetLength()

    horVec = XYZ(horizontal_vector.X, horizontal_vector.Y, 0).Normalize()
    norVec = rebarVec.CrossProduct(horVec).Normalize()

    # line_start = bottom_point
    # line_end = bottom_point + norVec * 100
    # sketch_plane = SketchPlane.Create(doc, Plane.CreateByThreePoints(line_start, line_end, XYZ(0,0,0)))
    # model_line = doc.Create.NewModelCurve(Line.CreateBound(line_start, line_end), sketch_plane)
#get surface paralel to the conSurface with offset of cover distance
    conSurface = Plane.CreateByNormalAndOrigin(conSurface.Normal, conSurface.Origin + conSurface.Normal * cover * cover_dir_corector)
    conSurface3 = Plane.CreateByNormalAndOrigin(conSurface.Normal, conSurface.Origin + conSurface.Normal * (-600*0.00328084  - cover))

#get the top handle between bar segment 1 and 3
    for handle in handleList:
        if handle.GetHandleName() == "Bar Segment 1":
            handle1_Z = handle.GetHandleSurface().Origin.Z
            for handle in handleList:
                if handle.GetHandleName() == "Bar Segment 3":
                    handle3_Z = handle.GetHandleSurface().Origin.Z
                    if handle1_Z > handle3_Z:
                        handle1_ontop = True
                    else:            
                        handle1_ontop = False


    for handle in handleList:
        print("Starting handle:" + str(handle.GetHandleName()))
        if handle.GetHandleName() == "Bar Segment 2":
            # constraint = conman.GetCurrentConstraintOnHandle(handle)
            constraint = RebarConstraint.CreateConstraintToSurface(handle, conSurface)
            # constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)[0]
            # print("Current constraint type:", constraint)
            # print("Is ToCover constraint:",  constraint.IsToCover())
            # print("Is ToHostFace constraint:", constraint.IsToHostFaceOrCover())
            if constraint and constraint.IsToCover():
                new_offset = -cover  
                constraint.SetDistanceToTargetCover(new_offset)
                # cons = RebarConstraint.CreateConstraintToSurface(handle, plane)
                conman.SetPreferredConstraint(constraint)
                doc.Regenerate()
                # print("Updated:", constraint.GetDistanceToTargetCover())
            elif constraint and constraint.IsToHostFaceOrCover(): 
                new_offset = -cover  
                constraint.SetDistanceToTargetHostFace(new_offset)
                # cons = RebarConstraint.CreateConstraintToSurface(handle, plane)
                constraint.SetDistanceToTargetHostFace(new_offset)
                conman.SetPreferredConstraint(constraint)
                doc.Regenerate()
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
            print("##### Bar Segment 2 - Handle processed #####")   

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

        if handle.GetHandleName() == "Bar Segment 1":
            # constraint = conman.GetCurrentConstraintOnHandle(handle)
            # constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)[0]
            top_constraint_plane = Plane.CreateByNormalAndOrigin(norVec, bottom_point + XYZ(0,0,1) * (cuvertHeight  - botSlabHeight -Cul_Haunch_Bot_D - cover + botPointZ_corrector))
            bot_constraint_plane = Plane.CreateByNormalAndOrigin(norVec, bottom_point + XYZ(0,0,1) * (- botSlabHeight -Cul_Haunch_Bot_D + cover + botPointZ_corrector))

            if handle1_ontop:
                faceTemp1 = top_constraint_plane
            else:
                faceTemp1 = bot_constraint_plane
            
            constraint = RebarConstraint.CreateConstraintToSurface(handle, faceTemp1)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
            print("##### Bar Segment 1 - Handle processed #####") 

        if handle.GetHandleName() == "Bar Segment 3":
            top_constraint_plane = Plane.CreateByNormalAndOrigin(norVec, bottom_point + XYZ(0,0,1) * (cuvertHeight - botSlabHeight -Cul_Haunch_Bot_D - cover + botPointZ_corrector))
            bot_constraint_plane = Plane.CreateByNormalAndOrigin(norVec, bottom_point + XYZ(0,0,1) * (- botSlabHeight -Cul_Haunch_Bot_D + cover + botPointZ_corrector))

            if handle1_ontop:
                faceTemp3 = bot_constraint_plane
            else:
                faceTemp3 = top_constraint_plane

            constraint = RebarConstraint.CreateConstraintToSurface(handle, faceTemp3)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()

            if constraint and constraint.IsToCover():
                new_offsetBot = -cover
                constraint.SetDistanceToTargetCover(new_offsetBot)
                # cons = RebarConstraint.CreateConstraintToSurface(handle, plane)
                conman.SetPreferredConstraint(constraint)
                doc.Regenerate()
            elif constraint and constraint.IsToHostFaceOrCover(): 
                new_offsetBot = -cover
                constraint.SetDistanceToTargetHostFace(new_offsetBot)
                # cons = RebarConstraint.CreateConstraintToSurface(handle, plane)
                constraint.SetDistanceToTargetHostFace(new_offsetBot)
                conman.SetPreferredConstraint(constraint)
                doc.Regenerate()
            print("##### Bar Segment 3 - Handle processed #####")

        if handle.GetHandleName() == "Start of Bar":
            constraint = RebarConstraint.CreateConstraintToSurface(handle, conSurface3)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
            print("##### Bar Segment 2 - Handle processed #####")   

        if handle.GetHandleName() == "End of Bar":
            constraint = RebarConstraint.CreateConstraintToSurface(handle, conSurface3)
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
            print("##### Out of Plane Extent - Handle processed #####") 

    rebar.LookupParameter("Mark").Set(Wall_Vertical)
    print("#"*100)

##########################################################################################################################################################################################################################################################
#endregion

##########################################################################################################################################################################################################################################################
#region -  Transverse rebar
##########################################################################################################################################################################################################################################################
print("#"*100)
print("Starting transverse rebar creation")
print("#"*100)

for line in rebarAnchorLine:
    bottom_point = get_bottom_point_of_line(line)
    rebar_vecLength = -(get_bottom_point_of_line(leftLineFront) - get_bottom_point_of_line(leftLineBack)).GetLength()
    rebar_vector = -(get_bottom_point_of_line(leftLineFront) - get_bottom_point_of_line(leftLineBack)).Normalize()
    horizontal_vector = rebar_vector.CrossProduct(XYZ(0,0,1))
    vertical_vector = horizontal_vector.CrossProduct(rebar_vector).Normalize()
    if line.Direction.Z > 0:
        horizontal_vector = -horizontal_vector
    min_line_length = min(c.Length for c in rebarAnchorLine)
    if line.Length > min_line_length*1.01:
        horizontal_vector = -horizontal_vector
        External = True
    else:
        External = False

    externalBotPoints = []
    for li in rebarAnchorLine:
        if External:
            externalBotPoints.append(get_bottom_point_of_line(li))

    rebar = Structure.Rebar.CreateFromRebarShape(doc, sc_20, barType16, element, bottom_point, rebar_vector, horizontal_vector)
    rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(Spacing * 0.00328084, 5,False , True, True) #SetLayoutAsMaximumSpacing(double spacing,double arrayLength,bool barsOnNormalSide,bool includeFirstBar,bool includeLastBar)
    shape_accessor = rebar.GetShapeDrivenAccessor()
    shape_accessor.UseRebarConstraintsToProduceVaryingBars = True

    for f in face_candidates:
        if is_line_in_plane(line, f.GetSurface()) and f.Id != front_face.Id:
            reinforcement_face = f.GetSurface()
            reinforcement_face_Id = f.Id

#set rebar constraint handles to the curve loop
    handleList = rebar.GetRebarConstraintsManager().GetAllHandles()
    conman = rebar.GetRebarConstraintsManager()
    conSurface = reinforcement_face

    for handle in handleList:
        print("Starting handle:" + str(handle.GetHandleName()))
        if handle.GetHandleName() == "Bar Segment 1":
            constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
            for const in constraint:
                conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
                if conSur.Id == reinforcement_face_Id :
                        constraint = const
                        conman.SetPreferredConstraint(constraint)
                        doc.Regenerate()
            if constraint and constraint.IsToCover():
                new_offset = -cover - barSize 
                constraint.SetDistanceToTargetCover(new_offset)
                conman.SetPreferredConstraint(constraint)
                doc.Regenerate()
            elif constraint and constraint.IsToHostFaceOrCover(): 
                new_offset = -cover  - barSize 
                constraint.SetDistanceToTargetHostFace(new_offset)
                conman.SetPreferredConstraint(constraint)
                doc.Regenerate()
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()

            print("##### Bar Segment 1 - Handle processed #####")   

        if handle.GetHandleName() == "Start of Bar": 
            constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
            for const in constraint:
                conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
                if conSur.Id == back_face.Id:
                    constraint = const
                    conman.SetPreferredConstraint(constraint)
                    doc.Regenerate()
                    break   

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
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
            print("##### Start of Bar - Handle processed #####")   

        if handle.GetHandleName() == "End of Bar": 
            constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
            for const in constraint:
                conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
                if conSur.Id == front_face.Id:
                    constraint = const
                    conman.SetPreferredConstraint(constraint)
                    doc.Regenerate()

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
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
            print("##### End of Bar - Handle processed #####")   

        if handle.GetHandleName() == "Out of Plane Extent":
            constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
            for const in constraint:
                conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
                if conSur.Id == bot_face.Id:
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
                constraint.SetDistanceToTargetHostFace(new_offset)
                conman.SetPreferredConstraint(constraint)
                doc.Regenerate()
            print("##### Out of Plane Extent - Handle processed #####")

        # if handle.GetHandleName() == "Bar Segment 1":
        #     rebarLine = Line.CreateBound(bottom_point, bottom_point + rebar_vector * 1000)
        #     rebarVecProjection = XYZ(rebar_vector.X, rebar_vector.Y, 0)
        #     skewCorrect = 0
        #     if line.Direction.Z < 0:
        #         print(line.Direction.Z)
        #         if not External:
        #             skewCorrect = Cell_Wall_T_Int*math.tan(skew)
        #     if External:
        #         skewCorrect = Cell_Wall_T_Ext*math.tan(skew)
        #     rebarBackSurface = Plane.CreateByNormalAndOrigin(rebarVecProjection, bottom_point + rebar_vector.Normalize() * (cover + skewCorrect))
        #     constraint = RebarConstraint.CreateConstraintToSurface(handle, rebarBackSurface)
        #     conman.SetPreferredConstraint(constraint)
        #     doc.Regenerate()

        #     print("##### Bar Segment 1 - Handle processed #####")

        # if handle.GetHandleName() == "Bar Segment 3":
        #     rebarLine = Line.CreateBound(bottom_point, bottom_point + rebar_vector * 1000)
        #     #create plane parallel to the back_face and going through the rebar line with offset of cover distance
        #     skewCorrect = 0
        #     if line.Direction.Z > 0:
        #         print(line.Direction.Z)
        #         if not External:
        #             skewCorrect = Cell_Wall_T_Int*math.tan(skew)
        #     if External:
        #         skewCorrect = Cell_Wall_T_Ext*math.tan(skew)
        #     rebarBackFace = Plane.CreateByNormalAndOrigin(back_face.ComputeNormal(UV(0,0)), back_face.GetSurface().Origin + back_face.GetSurface().Normal *( cover + skewCorrect))
        #     rebarBackSurface = Plane.CreateByNormalAndOrigin(rebarVecProjection, line_plane_intersection(rebarLine, rebarBackFace))
        #     constraint = RebarConstraint.CreateConstraintToSurface(handle, rebarBackSurface)
        #     conman.SetPreferredConstraint(constraint)
        #     doc.Regenerate()

        #     print("##### Bar Segment 3 - Handle processed #####")

        if handle.GetHandleName() == "Bar Plane":
            constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
            for const in constraint:
                conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
                # target_face = (top_face if line.Direction.Z > 0 else bot_face)
                if conSur.Id == top_face.Id:
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


    rebar.LookupParameter("Mark").Set(Wall_Horizontal)
    
    # val = max(rebar.LookupParameter("A").AsDouble(), rebar.LookupParameter("C").AsDouble())
    # print("A value in mm:", val * 304.8)
    # rebar.LookupParameter("A").Set(val)
    # rebar.LookupParameter("C").Set(val)
    print("#"*50)
##########################################################################################################################################################################################################################################################
# Rebar Transverse end bars

    if line.Direction.Z > 0 and not External or (line.Direction.Z < 0 and External):
        ##################################################################################################################################################################################################################################################
        #Front face
        rebar = Structure.Rebar.CreateFromRebarShape(doc, sc_38, barType16, element, bottom_point, horizontal_vector , rebar_vector)
        rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(Spacing * 0.00328084, 5,False , True, True) #SetLayoutAsMaximumSpacing(double spacing,double arrayLength,bool barsOnNormalSide,bool includeFirstBar,bool includeLastBar)
        shape_accessor = rebar.GetShapeDrivenAccessor()
        shape_accessor.UseRebarConstraintsToProduceVaryingBars = True

        
        rebarSec2Vector = horizontal_vector.CrossProduct(XYZ(0,0,1)).Normalize()
        rebarSec2Plane = Plane.CreateByNormalAndOrigin(rebarSec2Vector, bottom_point)
        #offset plane to the cover distance
        rebarSec2Plane = Plane.CreateByNormalAndOrigin(rebarSec2Plane.Normal, rebarSec2Plane.Origin + rebarSec2Plane.Normal * cover)

        handleList = rebar.GetRebarConstraintsManager().GetAllHandles()
        conman = rebar.GetRebarConstraintsManager()
        doc.Regenerate()

        for f in face_candidates:
            if is_line_in_plane(line, f.GetSurface()) and f.Id != front_face.Id:
                r_face = f.GetSurface()
                r_face_Id = f.Id
                r_faceSur = f

        if External:
            offset = Cell_Wall_T_Ext -cover -barSize
            mark = Wall_Horizontal_U_bar_Ext
        else:
            offset = -(Cell_Wall_T_Int -cover -barSize)
            mark = Wall_Horizontal_U_bar_Int
        
        rebar.LookupParameter("Mark").Set(mark)

        if point_plane_distance(bottom_point, left_face.GetSurface()) < 4:
            offset = -(Cell_Wall_T_Ext - cover - barSize)
            rebar.LookupParameter("Mark").Set(Wall_Horizontal_U_bar_Ext)

        conSur1 = Plane.CreateByNormalAndOrigin(reinforcement_face.Normal.Normalize(), reinforcement_face.Origin + reinforcement_face.Normal.Normalize() * offset *-1)
        for handle in handleList:
            print("Starting handle:" + str(handle.GetHandleName()))
            if handle.GetHandleName() == "Bar Segment 3":
                # constraint = conman.GetCurrentConstraintOnHandle(handle)
                constraint = RebarConstraint.CreateConstraintToSurface(handle, conSur1)
                conman.SetPreferredConstraint(constraint)
                doc.Regenerate()
                print("##### Bar Segment 3 - Handle processed #####")  

            if handle.GetHandleName() == "Start of Bar": 
                print("Starting handle:" + str(handle.GetHandleName()))
                #get plane parallel to the rebar segment 2 offset with cover distance
                rebarSec2Plane2 = Plane.CreateByNormalAndOrigin(rebarSec2Plane.Normal, rebarSec2Plane.Origin + rebarSec2Plane.Normal * (cover + barSize*50))
                constraint = RebarConstraint.CreateConstraintToSurface(handle, rebarSec2Plane2)
                conman.SetPreferredConstraint(constraint)
                doc.Regenerate()
                print("##### Start of Bar - Handle processed #####")   

            if handle.GetHandleName() == "End of Bar": 
                print("Starting handle:" + str(handle.GetHandleName()))
                #get plane parallel to the rebar segment 2 offset with cover distance
                rebarSec2Plane2 = Plane.CreateByNormalAndOrigin(rebarSec2Plane.Normal, rebarSec2Plane.Origin + rebarSec2Plane.Normal * (cover + barSize*50))
                constraint = RebarConstraint.CreateConstraintToSurface(handle, rebarSec2Plane2)
                conman.SetPreferredConstraint(constraint)
                doc.Regenerate()
                print("##### End of Bar - Handle processed #####")   

            if handle.GetHandleName() == "Out of Plane Extent":
                constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
                for const in constraint:
                    conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
                    if conSur.Id == top_face.Id:
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

            if handle.GetHandleName() == "Bar Segment 2":
                print("Starting handle:" + str(handle.GetHandleName()))
                constraint = RebarConstraint.CreateConstraintToSurface(handle, rebarSec2Plane)
                conman.SetPreferredConstraint(constraint)
                doc.Regenerate()
                print("##### Bar Segment 2 - Handle processed #####")

            if handle.GetHandleName() == "Bar Segment 1":
                constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
                for const in constraint:
                    conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
                    if conSur.Id == r_face_Id:
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
                print("##### Bar Segment 1 - Handle processed #####")

            if handle.GetHandleName() == "Bar Plane":
                constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
                for const in constraint:
                    conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
                    if conSur.Id == bot_face.Id:
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

                    ##################################################################################################################################################################################################################################################
        
        ##################################################################################################################################################################################################################################################
        #Back face
        bottom_point = bottom_point + -rebar_vector * rebar_vecLength
        rebar = Structure.Rebar.CreateFromRebarShape(doc, sc_38, barType16, element, bottom_point, horizontal_vector , -rebar_vector)
        rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(Spacing * 0.00328084, 5,False , True, True) #SetLayoutAsMaximumSpacing(double spacing,double arrayLength,bool barsOnNormalSide,bool includeFirstBar,bool includeLastBar)
        shape_accessor = rebar.GetShapeDrivenAccessor()
        shape_accessor.UseRebarConstraintsToProduceVaryingBars = True

        rebarSec2Vector = horizontal_vector.CrossProduct(XYZ(0,0,1)).Normalize()
        rebarSec2Plane = Plane.CreateByNormalAndOrigin(rebarSec2Vector, bottom_point)
        #offset plane to the cover distance
        rebarSec2Plane = Plane.CreateByNormalAndOrigin(rebarSec2Plane.Normal, rebarSec2Plane.Origin + rebarSec2Plane.Normal * -2*cover)

        handleList = rebar.GetRebarConstraintsManager().GetAllHandles()
        conman = rebar.GetRebarConstraintsManager()
        doc.Regenerate()

        for f in face_candidates:
            if is_line_in_plane(line, f.GetSurface()) and f.Id != front_face.Id:
                r_face = f.GetSurface()
                r_face_Id = f.Id
                r_faceSur = f

        if External:
            offset = Cell_Wall_T_Ext -cover -barSize
            mark = Wall_Horizontal_U_bar_Ext
        else:
            offset = -(Cell_Wall_T_Int -cover -barSize)
            mark = Wall_Horizontal_U_bar_Int

        rebar.LookupParameter("Mark").Set(mark) 

        if point_plane_distance(bottom_point, left_face.GetSurface()) < 4:
            offset = -(Cell_Wall_T_Ext - cover - barSize)
            rebar.LookupParameter("Mark").Set(Wall_Horizontal_U_bar_Ext)

        conSur1 = Plane.CreateByNormalAndOrigin(reinforcement_face.Normal.Normalize(), reinforcement_face.Origin + reinforcement_face.Normal.Normalize() * offset *-1)
        for handle in handleList:
            print("Starting handle:" + str(handle.GetHandleName()))
            if handle.GetHandleName() == "Bar Segment 3":
                # constraint = conman.GetCurrentConstraintOnHandle(handle)
                constraint = RebarConstraint.CreateConstraintToSurface(handle, conSur1)
                conman.SetPreferredConstraint(constraint)
                doc.Regenerate()
                print("##### Bar Segment 3 - Handle processed #####")  

            if handle.GetHandleName() == "Bar Segment 2":
                print("Starting handle:" + str(handle.GetHandleName()))
                constraint = RebarConstraint.CreateConstraintToSurface(handle, rebarSec2Plane)
                conman.SetPreferredConstraint(constraint)
                doc.Regenerate()
                print("##### Bar Segment 2 - Handle processed #####")

            if handle.GetHandleName() == "Start of Bar": 
                print("Starting handle:" + str(handle.GetHandleName()))
                #get plane parallel to the rebar segment 2 offset with cover distance
                rebarSec2Plane2 = Plane.CreateByNormalAndOrigin(rebarSec2Plane.Normal, rebarSec2Plane.Origin + rebarSec2Plane.Normal * -(cover + barSize*50))
                constraint = RebarConstraint.CreateConstraintToSurface(handle, rebarSec2Plane2)
                conman.SetPreferredConstraint(constraint)
                doc.Regenerate()
                print("##### Start of Bar - Handle processed #####")   

            if handle.GetHandleName() == "End of Bar": 
                print("Starting handle:" + str(handle.GetHandleName()))
                #get plane parallel to the rebar segment 2 offset with cover distance
                rebarSec2Plane2 = Plane.CreateByNormalAndOrigin(rebarSec2Plane.Normal, rebarSec2Plane.Origin + rebarSec2Plane.Normal * -(cover + barSize*50))
                constraint = RebarConstraint.CreateConstraintToSurface(handle, rebarSec2Plane2)
                conman.SetPreferredConstraint(constraint)
                doc.Regenerate()
                print("##### End of Bar - Handle processed #####")   

            if handle.GetHandleName() == "Out of Plane Extent":
                constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
                for const in constraint:
                    conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
                    if conSur.Id == top_face.Id:
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


            if handle.GetHandleName() == "Bar Segment 1":
                constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
                for const in constraint:
                    conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
                    if conSur.Id == r_face_Id:
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
                print("##### Bar Segment 1 - Handle processed #####")

            if handle.GetHandleName() == "Bar Plane":
                constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)
                for const in constraint:
                    conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
                    if conSur.Id == bot_face.Id:
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

        ##################################################################################################################################################################################################################################################
        #Wall Clips  
        clipcount = 0
        clipHeight = 0

        while clipHeight < cuvertHeight - ClipSpacing_Y*0.00328084 - botSlabHeight - Cul_Haunch_Bot_D - 2*cover:
            clipHeight =  ClipSpacing_Y*clipcount*0.00328084
            clipcount += 1
            print("Clip height:", clipHeight)
            print("$$$"*20)
            bottom_point = get_bottom_point_of_line(line)
            if line.Direction.Z > 0:
                vertical_vector = rebar_vector.CrossProduct(horizontal_vector).Normalize()
            else:
                vertical_vector = -rebar_vector.CrossProduct(horizontal_vector).Normalize()
            
            if point_plane_distance(bottom_point, right_face.GetSurface()) < 4 and External:
                bottom_point = bottom_point + XYZ(0,0,1) * (botSlabHeight + Cul_Haunch_Bot_D)
                vertical_vector = -vertical_vector

            rebar = Structure.Rebar.CreateFromRebarShape(doc, sc_85, barType10, element, bottom_point+ XYZ(0,0,1)*clipHeight,  horizontal_vector, vertical_vector)
            rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(ClipSpacing_X * 0.00328084, 5,True , False, False) #SetLayoutAsMaximumSpacing(double spacing,double arrayLength,bool barsOnNormalSide,bool includeFirstBar,bool includeLastBar)

        # to make sure the barSegment1 and barSegment3 connect to the correct element(slab or botslab) we check which handle is on top 
            for f in face_candidates:
                if is_line_in_plane(line, f.GetSurface()) and f.Id != front_face.Id:
                    reinforcement_face = f.GetSurface()
                    reinforcement_Sur = f

        # to corect the incorect sign on the cover offset of the outermost planes (abutments not piers) we get the lengt of the line and identify the line as an outher line if it is longer that the rest of the lines
            min_line_length = min(c.Length for c in rebarAnchorLine)
            print("Line length:", line.Length, "Min line length:", min_line_length)
            if line.Length > min_line_length*1.01: # if the line is more than 1% longer than the shortest line, we consider it as an outer line
                cover_dir_corector = -1
                botPointZ_corrector = botSlabHeight + Cul_Haunch_Bot_D
                External = True
                print("external")
            else:
                cover_dir_corector = 1
                botPointZ_corrector = 0
                External = False
                print("internal")
        #set rebar constraint handles to the curve loop
            handleList = rebar.GetRebarConstraintsManager().GetAllHandles()
            conman = rebar.GetRebarConstraintsManager()
            conSurface = reinforcement_face

        #colect all lines that go through bottom_point
            rebarVec = None
            for c in curve_candidates:
                    if (c.GetEndPoint(0).DistanceTo(bottom_point) < 1e-6 or c.GetEndPoint(1).DistanceTo(bottom_point) < 1e-6):
                        if not is_line_in_plane(c, front_face.GetSurface()):
                            rebarVec = c.Direction.Normalize()
                            rebarVecZ = abs(c.GetEndPoint(0).Z - c.GetEndPoint(1).Z)
                            rebarVecPlanLength = (XYZ(c.GetEndPoint(0).X - c.GetEndPoint(1).X, c.GetEndPoint(0).Y - c.GetEndPoint(1).Y, 0)).GetLength()


            for handle in handleList:
                print("Starting handle:" + str(handle.GetHandleName()))
                if handle.GetHandleName() == "Bar Segment 2":

                    if External:
                        offset = -Cell_Wall_T_Ext + cover
                    else:   
                        if point_plane_distance(bottom_point, left_face.GetSurface()) < 4:
                            offset = Cell_Wall_T_Ext - cover
                        else:
                            offset =  Cell_Wall_T_Int - cover
                    
                    constraint_plane = Plane.CreateByNormalAndOrigin(reinforcement_Sur.GetSurface().Normal.Normalize(), reinforcement_Sur.GetSurface().Origin + reinforcement_Sur.GetSurface().Normal.Normalize() * offset)
                    constraint = RebarConstraint.CreateConstraintToSurface(handle, constraint_plane)
                    conman.SetPreferredConstraint(constraint)
                    doc.Regenerate()
                    # constraint = rebar.GetRebarConstraintsManager().GetConstraintCandidatesForHandle(handle,element.Id)
                    # for const in constraint:
                    #     conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
                    #     if conSur.Id == reinforcement_Sur.Id:
                    #         constraint = const
                    # if External:
                    #     offset = Cell_Wall_T_Ext 
                    # else:   
                    #     offset = Cell_Wall_T_Int 
                    # doc.Regenerate()
                    # if constraint and constraint.IsToCover():
                    #     constraint.SetDistanceToTargetCover(-offset)
                    #     conman.SetPreferredConstraint(constraint)
                    #     doc.Regenerate()
                    # elif constraint and constraint.IsToHostFaceOrCover(): 
                    #     print("Host face to surface")
                    #     print(-offset + cover)
                    #     constraint.SetDistanceToTargetHostFace(-offset +cover)
                    #     conman.SetPreferredConstraint(constraint)
                    #     doc.Regenerate()
                    print("##### Bar Segment 2 - Handle processed #####")   
                    
                if handle.GetHandleName() == "Out of Plane Extent":
                    back_constraint_plane = Plane.CreateByNormalAndOrigin(rebar_vector, bottom_point - rebar_vector * rebar_vecLength)
                    constraint = RebarConstraint.CreateConstraintToSurface(handle, back_constraint_plane)
                    conman.SetPreferredConstraint(constraint)
                    doc.Regenerate()
                    print("##### Out of Plane Extent - Handle processed #####")

                if handle.GetHandleName() == "Bar Segment 4":
                    offset =  0
                    constraint_plane = Plane.CreateByNormalAndOrigin(reinforcement_Sur.GetSurface().Normal.Normalize(), reinforcement_Sur.GetSurface().Origin + reinforcement_Sur.GetSurface().Normal.Normalize() * offset)
                    constraint = RebarConstraint.CreateConstraintToSurface(handle, constraint_plane)
                    conman.SetPreferredConstraint(constraint)
                    doc.Regenerate()

                    print("##### Bar Segment 4 - Handle processed #####")

                if handle.GetHandleName() == "Start of Bar":
                    constraint = rebar.GetRebarConstraintsManager().GetConstraintCandidatesForHandle(handle,element.Id)
                    for const in constraint:
                        conSur = element.GetGeometryObjectFromReference(const.GetTargetHostFaceReference())
                        if conSur.Id == reinforcement_Sur.Id:
                            constraint = const
                    if constraint and constraint.IsToCover():
                        constraint.SetDistanceToTargetCover(Cell_Wall_T_Int/2 )
                        conman.SetPreferredConstraint(constraint)
                        doc.Regenerate()
                    elif constraint and constraint.IsToHostFaceOrCover(): 
                        new_offset = - Cell_Wall_T_Int/2
                        constraint.SetDistanceToTargetHostFace(new_offset)
                        conman.SetPreferredConstraint(constraint)
                        doc.Regenerate()
                    print("##### Start of Bar - Handle processed #####")   

            if External or point_plane_distance(bottom_point, left_face.GetSurface()) < 4:
                rebar.LookupParameter("B").Set(Cell_Wall_T_Ext-cover - barSize*2)
                rebar.LookupParameter("Mark").Set(Wall_Clips_external)
            else:
                rebar.LookupParameter("B").Set(Cell_Wall_T_Int-cover - barSize*2)
                rebar.LookupParameter("Mark").Set(Wall_Clips_internal)
            rebar.LookupParameter("C").Set(13*wallClipBarSize)
            doc.Regenerate()
        print("#"*100)

##########################################################################################################################################################################################################################################################
#endregion


# supress warnings ####################################################################
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SupressWarnings())
t.SetFailureHandlingOptions(failHandler)

t.Commit()

print("#"*100)
print("DONE")
print("#"*100)