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
element = doc.GetElement(ElementId(Int64(552179)))
# botslab = doc.GetElement(ElementId(Int64(508812)))
# slab = doc.GetElement(ElementId(Int64(502467)))
front_face_Id = 147
top_face_Id = 138
bot_face_Id = 127
back_face_Id = 181

topSlabHeight = element.LookupParameter("Roof_T").AsDouble()
print("Top slab height:", topSlabHeight*304.8, "ft")
botSlabHeight = element.LookupParameter("Flr_T").AsDouble()
print("Bottom slab height:", botSlabHeight*304.8, "ft")

def get_type_name(elem_type):
    name_param = elem_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
    return name_param.AsString() if name_param else None

def get_level_name(level_elem):
    name_param = level_elem.get_Parameter(BuiltInParameter.DATUM_TEXT)
    return name_param.AsString() if name_param else None

# Get a floor type with name "Generic 300mm"
floor_type = None
floor_type_FEC = FilteredElementCollector(doc).OfClass(FloorType)
for ftype in floor_type_FEC:
    ftype_name = get_type_name(ftype)
    print("Checking floor type:", ftype_name)
    if ftype_name == "Generic 300mm":
        floor_type = ftype
        print("Floor type found:", ftype_name)
        break

# Get hold of level 1 to place the dummy floor on it
level_FEC = FilteredElementCollector(doc).OfClass(Level)
for lvl in level_FEC:
    if lvl.Name == "Level 1":
        level = lvl
        print("Level found:", get_level_name(level))
        break


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
barType16 = barTypeBySize("Y16") # 10mm rebar
barSize = 10 * 0.00328084 # 10mm rebar

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

def filter_rebar_anchor_lines_on_face(curve_candidates, front_face):
    rebarAnchorLine = []
    for c in curve_candidates:
        if is_line_in_plane(c, front_face.GetSurface()):
            if abs(c.Direction.Normalize().Z) > 0.9:
                rebarAnchorLine.append(c)

    print("Number of rebar anchor lines:", len(rebarAnchorLine))
    print("#"*150)

    # remove lines that are not on the top of the element, but head wall lines in the front face plane
    for _ in range(2):  # Remove top 2 lines
        if not rebarAnchorLine:
            break
        maxZ = max(max(line.GetEndPoint(0).Z, line.GetEndPoint(1).Z) for line in rebarAnchorLine)
        print(maxZ)
        rebarAnchorLine = [line for line in rebarAnchorLine 
                          if maxZ not in (line.GetEndPoint(0).Z, line.GetEndPoint(1).Z)]
    
    return rebarAnchorLine

def is_face_parallel_to_face(face1, face2, tol=1e-9):
    n1 = face1.GetSurface().Normal.Normalize()
    n2 = face2.GetSurface().Normal.Normalize()
    return abs(n1.DotProduct(n2)) > tol

def draw_lines_around_face(face, doc):
    """Draw model lines around the perimeter of a face."""
    face_loops = face.GetEdgesAsCurveLoops()
    for loop in face_loops:
        for curve in loop:
            sketch_plane = SketchPlane.Create(doc, Plane.CreateByNormalAndOrigin(
                face.GetSurface().Normal, curve.GetEndPoint(0)))
            doc.Create.NewModelCurve(curve, sketch_plane)

def is_curves_planar(curve_list, tol=1e-9):
    if len(curve_list) < 3:
        return True
    p0 = curve_list[0].GetEndPoint(0)
    p1 = curve_list[0].GetEndPoint(1)
    v1 = (p1 - p0).Normalize()
    normal = None
    for curve in curve_list[1:]:
        q0 = curve.GetEndPoint(0)
        q1 = curve.GetEndPoint(1)
        v2 = (q1 - q0).Normalize()
        cross = v1.CrossProduct(v2)
        if cross.GetLength() > tol:
            normal = cross.Normalize()
            break
    if normal is None:
        return True
    for curve in curve_list:
        for t in [0, 0.5, 1]:
            pt = curve.Evaluate(t, True)
            if abs(normal.DotProduct(pt - p0)) > tol:
                return False
    return True

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

print("Front face Id:", front_face.Id)
print("Top face Id:", top_face.Id)  
print("Bottom face Id:", bot_face.Id)
print("Back face Id:", back_face.Id)
####################################################################################################################################################################################################

t = Transaction(doc, "Create Dummy Floors")
t.Start()  

Fcorner_points = []
Bcorner_points = []
frontLines = filter_rebar_anchor_lines_on_face(curve_candidates, front_face)
backLines = filter_rebar_anchor_lines_on_face(curve_candidates, back_face)

#get the four line end points with the highest Z value as corner points for the dummy floor
for line in frontLines:
    p1 = line.GetEndPoint(0)
    p2 = line.GetEndPoint(1)
    if p1.Z > p2.Z:
        Fcorner_points.append(p1)
    else:
        Fcorner_points.append(p2)

for line in backLines:
    p1 = line.GetEndPoint(0)
    p2 = line.GetEndPoint(1)
    if p1.Z > p2.Z:
        Bcorner_points.append(p1)
    else:
        Bcorner_points.append(p2)
print("Corner points for dummy floor:")
print(len(Fcorner_points), len(Bcorner_points))
corner_points = []
corner_points.extend(sorted(Fcorner_points, key=lambda p: p.Z, reverse=True)[:2])
corner_points.extend(sorted(Bcorner_points, key=lambda p: p.Z, reverse=True)[:2])
print("Corner points for dummy floor:")
print(len(corner_points))
# curve_loop = CurveLoop()
# get  IList CurveLoop from the corner points and ensure the loop is closed and the points are in the correct order
frontPoints = []
backPoints = []
for point in corner_points:
    if is_point_on_face(point, front_face):
        frontPoints.append(point)
        print("Point on front face:")
    elif is_point_on_face(point, back_face):
        backPoints.append(point)
        print("Point on back face:")

if len(frontPoints) != 2 or len(backPoints) != 2:
    raise Exception("Expected exactly 2 front points and 2 back points to build dummy floor loop.")

# Pair front points to back points using shortest total matching
f0, f1 = frontPoints[0], frontPoints[1]
b0, b1 = backPoints[0], backPoints[1]
pair_a = f0.DistanceTo(b0) + f1.DistanceTo(b1)
pair_b = f0.DistanceTo(b1) + f1.DistanceTo(b0)

if pair_a <= pair_b:
    b_for_f0, b_for_f1 = b0, b1
else:
    b_for_f0, b_for_f1 = b1, b0

# Build contiguous loop: F0 -> F1 -> B1 -> B0 -> F0


frontLine = Line.CreateBound(XYZ(f0.X, f0.Y, 0), XYZ(f1.X, f1.Y, 0))
rightLine = Line.CreateBound(XYZ(f1.X, f1.Y, 0), XYZ(b_for_f1.X, b_for_f1.Y, 0))
backLine = Line.CreateBound(XYZ(b_for_f1.X, b_for_f1.Y, 0), XYZ(b_for_f0.X, b_for_f0.Y, 0))
leftLine = Line.CreateBound(XYZ(b_for_f0.X, b_for_f0.Y, 0), XYZ(f0.X, f0.Y, 0))

crvList = List[Curve]()
crvList.Add(frontLine)
crvList.Add(rightLine)
crvList.Add(backLine)
crvList.Add(leftLine)
curves = crvList

curve_loop = CurveLoop.Create(curves)
IlistCurveLoop = List[CurveLoop]()
IlistCurveLoop.Add(curve_loop)

dummy_floor_top = Floor.Create(doc, IlistCurveLoop, floor_type.Id, level.Id)#,True, slope_line, slope)

frontLine = Line.CreateBound(f0, f1)
rightLine = Line.CreateBound(f1, b_for_f1)
backLine = Line.CreateBound(b_for_f1, b_for_f0)
leftLine = Line.CreateBound(b_for_f0, f0)

crvList = List[Curve]()
crvList.Add(frontLine)
crvList.Add(rightLine)
crvList.Add(backLine)
crvList.Add(leftLine)
curves = crvList

curve_loop = CurveLoop.Create(curves)
IlistCurveLoop = List[CurveLoop]()
IlistCurveLoop.Add(curve_loop)
# draw line to check the position of the curve loop
# for line in curves:
#     modelLine = doc.Create.NewModelCurve(line, SketchPlane.Create(doc, Plane.CreateByThreePoints(line.GetEndPoint(0), line.GetEndPoint(1), XYZ.BasisX)))

# t.Commit()



structural_param = dummy_floor_top.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL)
if structural_param is None:
    structural_param = dummy_floor_top.LookupParameter("Structural")
structural_param.Set(1)

# try slab-shape editing 
slab_shape_editor = None
if hasattr(dummy_floor_top, "GetSlabShapeEditor"):
    slab_shape_editor = dummy_floor_top.GetSlabShapeEditor()
elif hasattr(dummy_floor_top, "SlabShapeEditor"):
    slab_shape_editor = dummy_floor_top.SlabShapeEditor

if slab_shape_editor is not None:
    slab_shape_editor.Enable()
    print("SlabShapeEditor enabled for dummy floor.")
else:
    print("SlabShapeEditor not available on this Floor API version; skipping slab-shape edits.")


subelements_top = dummy_floor_top.GetSubelements() 
print("Sub elements of the dummy floor on top face:")
print(subelements_top.Count)

doc.Regenerate()

if slab_shape_editor is not None:
    for curve in IlistCurveLoop[0]:
        print("Original curve start point:", curve.GetEndPoint(0))
        print("Original curve end point:", curve.GetEndPoint(1))
        point = curve.GetEndPoint(0)
        slab_shape_editor.AddPoint(point)
    

translation_vector = top_face.Origin.Z - bot_face.Origin.Z -botSlabHeight   
print("Translation vector for slab shape editor:", translation_vector*0.3048,)
#copy dummy_floor_top down

dummy_floor_botid = ElementTransformUtils.CopyElement(doc, dummy_floor_top.Id, XYZ(0,0,-translation_vector))
dummy_floor_bot   = None
for floor_id in dummy_floor_botid:
    dummy_floor_bot = doc.GetElement(floor_id)
    break

if dummy_floor_bot is None:
    raise Exception("Failed to copy top dummy floor.")

# duplicate floor type, rename, and set target thickness
def duplicate_floor_type_with_thickness(base_floor_type, base_name, target_thickness):
    if base_floor_type is None:
        raise Exception("Base floor type is None.")

    all_floor_types = FilteredElementCollector(doc).OfClass(FloorType).ToElements()
    existing_type_names = []
    for f_type in all_floor_types:
        name_param = f_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        type_name = name_param.AsString() if name_param else None
        if type_name:
            existing_type_names.append(type_name)

    unique_type_name = base_name
    suffix = 1
    while unique_type_name in existing_type_names:
        unique_type_name = "{0} {1}".format(base_name, suffix)
        suffix += 1

    new_floor_type = base_floor_type.Duplicate(unique_type_name)
    compound_structure = new_floor_type.GetCompoundStructure()
    if compound_structure is None:
        raise Exception("Cannot edit thickness: duplicated floor type has no compound structure.")

    layers = compound_structure.GetLayers()
    if layers is None or layers.Count == 0:
        raise Exception("Cannot edit thickness: duplicated floor type has no compound layers.")

    current_total_thickness = 0.0
    for layer in layers:
        current_total_thickness += layer.Width

    delta_thickness = target_thickness - current_total_thickness
    new_first_layer_width = layers[0].Width + delta_thickness
    if new_first_layer_width <= 0:
        raise Exception("Requested thickness results in non-positive first layer width.")

    compound_structure.SetLayerWidth(0, new_first_layer_width)
    new_floor_type.SetCompoundStructure(compound_structure)
    return new_floor_type, unique_type_name

# apply dedicated type to top dummy floor
top_floor_type, top_type_name = duplicate_floor_type_with_thickness(floor_type, "Dummy Top Floor Type", topSlabHeight)
dummy_floor_top.ChangeTypeId(top_floor_type.Id)
print("Applied floor type:", top_type_name)
print("Top floor thickness (mm):", topSlabHeight * 304.8)

# apply dedicated type to bottom dummy floor
bot_floor_type, bot_type_name = duplicate_floor_type_with_thickness(floor_type, "Dummy Bottom Floor Type", botSlabHeight)
dummy_floor_bot.ChangeTypeId(bot_floor_type.Id)
print("Applied floor type:", bot_type_name)
print("Bottom floor thickness (mm):", botSlabHeight * 304.8)


####TOP SLAB FACES################################################################################################

Top_slab_face_candidates  = []
top_geom = dummy_floor_top.get_Geometry(Options())
for geo_obj in top_geom:
    if isinstance(geo_obj, Face):
        Top_slab_face_candidates.append(geo_obj)
    elif isinstance(geo_obj, Solid) and geo_obj.Volume > 0:
        for face in geo_obj.Faces:
            Top_slab_face_candidates.append(face)
    elif isinstance(geo_obj, GeometryInstance):
        for inst_obj in geo_obj.GetInstanceGeometry():
            if isinstance(inst_obj, Face):
                Top_slab_face_candidates.append(inst_obj)
            elif isinstance(inst_obj, Solid) and inst_obj.Volume > 0:
                for face in inst_obj.Faces:
                    Top_slab_face_candidates.append(face)

#get top face of a top slab by getting the two faces with the largest area and after that checking which one is on top by comparing the Z value of the face origin 
max_area_list = sorted(Top_slab_face_candidates, key=lambda f: f.Area, reverse=True)[:2]
if max_area_list[0].Origin.Z > max_area_list[1].Origin.Z:
    top_slab_top_face = max_area_list[0]
    top_slab_bot_face = max_area_list[1]
else:
    top_slab_top_face = max_area_list[1]
    top_slab_bot_face = max_area_list[0]

topSlabFaceListTemp = list(set(Top_slab_face_candidates) - set(max_area_list))

for face in topSlabFaceListTemp:
    if is_line_in_plane(frontLine, face.GetSurface()):
        top_slab_front_face = face
    elif is_line_in_plane(rightLine, face.GetSurface()):
        top_slab_right_face = face
    elif is_line_in_plane(backLine, face.GetSurface()):
        top_slab_back_face = face
    elif is_line_in_plane(leftLine, face.GetSurface()):
        top_slab_left_face = face

####BOTTOM SLAB FACES################################################################################################

Bot_slab_face_candidates  = []
bot_geom = dummy_floor_bot.get_Geometry(Options())
for geo_obj in bot_geom:
    if isinstance(geo_obj, Face):
        Bot_slab_face_candidates.append(geo_obj)
    elif isinstance(geo_obj, Solid) and geo_obj.Volume > 0:
        for face in geo_obj.Faces:
            Bot_slab_face_candidates.append(face)
    elif isinstance(geo_obj, GeometryInstance):
        for inst_obj in geo_obj.GetInstanceGeometry():
            if isinstance(inst_obj, Face):
                Bot_slab_face_candidates.append(inst_obj)
            elif isinstance(inst_obj, Solid) and inst_obj.Volume > 0:
                for face in inst_obj.Faces:
                    Bot_slab_face_candidates.append(face)

#get top face of a top slab by getting the two faces with the largest area and after that checking which one is on top by comparing the Z value of the face origin 
max_area_list = sorted(Bot_slab_face_candidates, key=lambda f: f.Area, reverse=True)[:2]
if max_area_list[0].Origin.Z > max_area_list[1].Origin.Z:
    bot_slab_top_face = max_area_list[0]
    bot_slab_bot_face = max_area_list[1]
else:
    bot_slab_top_face = max_area_list[1]
    bot_slab_bot_face = max_area_list[0]

botSlabFaceListTemp = list(set(Bot_slab_face_candidates) - set(max_area_list))

for face in botSlabFaceListTemp:
    if is_line_in_plane(frontLine, face.GetSurface()):
        bot_slab_front_face = face
    elif is_line_in_plane(rightLine, face.GetSurface()):
        bot_slab_right_face = face
    elif is_line_in_plane(backLine, face.GetSurface()):
        bot_slab_back_face = face
    elif is_line_in_plane(leftLine, face.GetSurface()):
        bot_slab_left_face = face

######################################################################################################################

# supress warnings ####################################
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SupressWarnings())
t.SetFailureHandlingOptions(failHandler)

t.Commit()


t = Transaction(doc, "Create Free Form Rebar")
t.Start()


rebarAnchorLine = filter_rebar_anchor_lines_on_face(curve_candidates, front_face)

print("Number of rebar anchor lines after removing head wall lines:", len(rebarAnchorLine))
print("#"*150)


for line in rebarAnchorLine:
    bottom_point = get_bottom_point_of_line(line)
    vertical_vector = line.Direction.Normalize()
    horizontal_vector = vertical_vector.CrossProduct(front_face.ComputeNormal(UV(0,0))).Normalize()
    rebar = Structure.Rebar.CreateFromRebarShape(doc, sc_38, barType16, element, bottom_point, vertical_vector, horizontal_vector)
    rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(300 * 0.00328084, 5,True , True, True) #SetLayoutAsMaximumSpacing(double spacing,double arrayLength,bool barsOnNormalSide,bool includeFirstBar,bool includeLastBar)
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
    else:
        cover_dir_corector = 1

#set rebar constraint handles to the curve loop
    handleList = rebar.GetRebarConstraintsManager().GetAllHandles()
    conman = rebar.GetRebarConstraintsManager()
    conSurface = reinforcement_face

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
        def get_constraint_target_face(constr):
            if constr is None:
                return None
            face_ref = constr.GetTargetHostFaceReference()
            if face_ref is None:
                return None
            target_elem = constr.GetTargetElement()
            if target_elem is None:
                target_elem = doc.GetElement(face_ref.ElementId)
            if target_elem is None:
                return None
            try:
                return target_elem.GetGeometryObjectFromReference(face_ref)
            except:
                return None

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
            candidate_constraints = conman.GetConstraintCandidatesForHandle(handle, element.Id)
            if candidate_constraints.Count == 0:
                continue
            idx = 2 if candidate_constraints.Count > 2 else 0
            constraint = candidate_constraints[idx]
            conSurface = get_constraint_target_face(constraint)
            if conSurface is None:
                continue
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
            # constraint = RebarConstraint.CreateConstraintToSurface(handle, conSurface2)
            # constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)[0]
            if handle1_ontop:
                tempSlab = dummy_floor_top
                topSlab_face_Id = top_slab_top_face.Id
            else:
                tempSlab = dummy_floor_bot
                topSlab_face_Id = bot_slab_bot_face.Id
            constraint = None
            constraints = conman.GetConstraintCandidatesForHandle(handle,tempSlab.Id)
            for const in constraints:
                conSurface_temp = get_constraint_target_face(const)
                if conSurface_temp is not None and conSurface_temp.Id == topSlab_face_Id:
                    constraint = const
                    break

            if constraint and constraint.IsToCover():
                new_offset = -cover
                constraint.SetDistanceToTargetCover(new_offset)
                # cons = RebarConstraint.CreateConstraintToSurface(handle, plane)
                conman.SetPreferredConstraint(constraint)
                doc.Regenerate()
            elif constraint and constraint.IsToHostFaceOrCover(): 
                new_offset = -cover
                constraint.SetDistanceToTargetHostFace(new_offset)
                # cons = RebarConstraint.CreateConstraintToSurface(handle, plane)
                constraint.SetDistanceToTargetHostFace(new_offset)
                conman.SetPreferredConstraint(constraint)
                doc.Regenerate()
            print("##### Bar Segment 1 - Handle processed #####") 

        if handle.GetHandleName() == "Bar Segment 3":
            # constraint = conman.GetCurrentConstraintOnHandle(handle)
            # constraint = RebarConstraint.CreateConstraintToSurface(handle, conSurface2)
            # constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)[0]
            if handle1_ontop:
                tempSlab3 = dummy_floor_bot
                tempSlab3_face_Id = bot_slab_bot_face.Id
                print(tempSlab3_face_Id)
                print(tempSlab3.Id)
            else:
                tempSlab3 = dummy_floor_top
                tempSlab3_face_Id = top_slab_top_face.Id
            constraint = None
            constraints = conman.GetConstraintCandidatesForHandle(handle,tempSlab3.Id)
            for const in constraints:
                conSurface_temp = get_constraint_target_face(const)
                if conSurface_temp is not None and conSurface_temp.Id == tempSlab3_face_Id:
                    constraint = const
                    break

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
            candidate_constraints = conman.GetConstraintCandidatesForHandle(handle,element.Id)
            constraint = None
            for const in candidate_constraints:
                conSur = get_constraint_target_face(const)
                if conSur is not None and conSur.Id == front_face.Id:
                    constraint = const
                    break

            if constraint is None:
                continue

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

    print("#"*100)

# supress warnings ####################################################################
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SupressWarnings())
t.SetFailureHandlingOptions(failHandler)

t.Commit()

print("#"*100)
print("DONE")
print("#"*100)