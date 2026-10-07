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
Spacing = 300 * 0.00328084

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

Cul_Haunch_Top_D = element.LookupParameter("Cul_Haunch_Top_D").AsDouble()
Cul_Haunch_Top_W = element.LookupParameter("Cul_Haunch_Top_W").AsDouble()

Cul_Haunch_Bot_D = element.LookupParameter("Cul_Haunch_Bot_D").AsDouble()
Cul_Haunch_Bot_W = element.LookupParameter("Cul_Haunch_Bot_W").AsDouble()

Cell_Wall_T_Ext = element.LookupParameter("Cell_Wall_T_Ext").AsDouble()
Cell_Wall_T_Int = element.LookupParameter("Cell_Wall_T_Int").AsDouble()

###### MARKS ##################################################################################

Top_Haunch_bars_Abutment    = "Top Haunch Bars Abutment"
Bottom_Haunch_bars_Abutment = "Bottom Haunch Bars Abutment"
Top_Haunch_bars_Pier        = "Top Haunch Bars Pier"
Bottom_Haunch_bars_Pier     = "Bottom Haunch Bars Pier"


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
sc_48 = getRebarShapeByName("48")
barType16 = barTypeBySize("Y16") # 16mm rebar
barSize = 16 * 0.00328084 # 16mm rebar

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

def get_top_point_of_line(line):
    if line.IsBound:
        p1 = line.GetEndPoint(0)
        p2 = line.GetEndPoint(1)
        return p1 if p1.Z > p2.Z else p2
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

def draw_line_in_view(start , end):
    if start is None or end is None:
        return
    doc.Create.NewModelCurve(Line.CreateBound(start, end), SketchPlane.Create(doc, Plane.CreateByThreePoints(start, end, XYZ(0,0,0))))

def sc48_by4points(rebar_p1, rebar_p2, rebar_p3,rebar_p4,Plane, element, barType =barType16, document = doc, sc_48 = sc_48, ModelLine = False):

        #place curves
        curve1 = Line.CreateBound(rebar_p1, rebar_p2)
        curve2 = Line.CreateBound(rebar_p2, rebar_p3)
        curve3 = Line.CreateBound(rebar_p3, rebar_p4)

        geomPlane = Plane.CreateByThreePoints(rebar_p1, rebar_p2, rebar_p3)
        sketch = SketchPlane.Create(doc, geomPlane)

        if ModelLine == True:
            model_line = doc.Create.NewModelCurve(curve1, sketch)
            model_line = doc.Create.NewModelCurve(curve2, sketch)
            model_line = doc.Create.NewModelCurve(curve3, sketch)
        else:
           #### Cast the list to IList<Curve>
            curve_list48 = List[Curve]([curve1, curve2, curve3])
            
            #### Bluid ####################################################################        
            rebar = Structure.Rebar.CreateFromCurvesAndShape(doc, 
                                                sc_48, 
                                                barType, 
                                                element,
                                                Plane.Normal,  
                                                curve_list48, 
                                                BarTerminationsData(doc))
            
        return rebar
    
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

def project_point_to_plane(point, plane):
    n = plane.Normal.Normalize()
    d = n.DotProduct(point - plane.Origin)   # signed distance
    return point - d * n                     # projected XYZ on infinite plane

def line_plane_intersection(line, plane, tol=1e-9):
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

def set_rebar_dim_to_avg(rebar_mark):
    FEC = FilteredElementCollector(doc).OfClass(Rebar).WhereElementIsNotElementType()
    rebar_instances = []
    for r in FEC:
        mark_param = r.LookupParameter("Mark")
        mark_value = mark_param.AsString() if mark_param else None
        if mark_value == rebar_mark:
            rebar_instances.append(r)

    if not rebar_instances:
        print("No rebars found for mark: " + str(rebar_mark))
        return

    paramList = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O"]
    paramDict = {param: [] for param in paramList}

    for rebar in rebar_instances:
        for param in paramList:
            # Not every shape has every dimension (A..O), so skip missing params safely.
            rebar_param = rebar.LookupParameter(param)
            if rebar_param:
                paramDict[param].append(rebar_param.AsDouble())

    for param in paramList:
        avgValue = sum(paramDict[param]) / len(paramDict[param]) if paramDict[param] else 0
        maxValue = max(paramDict[param]) if paramDict[param] else 0
        minValue = min(paramDict[param]) if paramDict[param] else 0
        for rebar in rebar_instances:
            rebar_param = rebar.LookupParameter(param)
            if rebar_param is None:
                continue
            if max(maxValue - avgValue, avgValue - minValue) > 0.1:
                print("Setting " + str(param) +  " MAX: " + str(max(maxValue - avgValue, avgValue - minValue) * 304.8) +"mm   to average value: " + str(avgValue*304.8))
            rebar_param.Set(avgValue)



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

leftLineFront, rightLineFront, leftLineBack, rightLineBack = collect_corner_lines(curve_candidates, front_face, back_face, left_face, right_face)

topPoint_front_left = get_top_point_of_line(leftLineFront)
topPoint_front_right = get_top_point_of_line(rightLineFront)
topPoint_back_left = get_top_point_of_line(leftLineBack)
topPoint_back_right = get_top_point_of_line(rightLineBack)

botPoint_front_left = get_bottom_point_of_line(leftLineFront)
botPoint_front_right = get_bottom_point_of_line(rightLineFront)
botPoint_back_left = get_bottom_point_of_line(leftLineBack)
botPoint_back_right = get_bottom_point_of_line(rightLineBack)
 
horVec =  botPoint_front_right - botPoint_front_left

print("Front face Id:", front_face.Id)
print("Top face Id:", top_face.Id)  
print("Bottom face Id:", bot_face.Id)

t = Transaction(doc, "Create Haunch Rebar")
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

for line in rebarAnchorLine:
    if line.GetEndPoint(0).Z== maxZ or line.GetEndPoint(1).Z == maxZ:
        rebarAnchorLine.remove(line)
#remove second line
maxZ = max(max(line.GetEndPoint(0).Z, line.GetEndPoint(1).Z) for line in rebarAnchorLine)
cuvertHeight = max(line.Length for line in rebarAnchorLine)
# print("Cuvet height:", cuvertHeight)

for line in rebarAnchorLine:
    if line.GetEndPoint(0).Z== maxZ or line.GetEndPoint(1).Z == maxZ:
        rebarAnchorLine.remove(line)

print("Number of rebar anchor lines after removing head wall lines:", len(rebarAnchorLine))
print("#"*150)
#remove the two outer lines that are more than 1% longer than the shortest line
min_line_length = min(c.Length for c in rebarAnchorLine)
for line in rebarAnchorLine:
    if abs(line.Length) > abs(min_line_length*1.01): # if the line is more than 1% longer than the shortest line, it is an outer line
        rebarAnchorLine.remove(line)

for line in rebarAnchorLine:
    if abs(line.Length) > abs(min_line_length*1.01): # if the line is more than 1% longer than the shortest line, it is an outer line
        rebarAnchorLine.remove(line)



print("Number of rebar anchor lines after removing head wall lines:", len(rebarAnchorLine))

for line in rebarAnchorLine:
    bottom_point = get_bottom_point_of_line(line)

    distToLeft = (bottom_point - project_point_to_plane(bottom_point, left_face.GetSurface())).GetLength()
    distToRight = (bottom_point - project_point_to_plane(bottom_point, right_face.GetSurface())).GetLength()
    if distToLeft < 4 or distToRight < 4:
        Exterior_face = True
    else:        
        Exterior_face = False

    V = -line.Direction.Normalize()
    U = horVec.Normalize()

    for c in curve_candidates:
        if (c.GetEndPoint(0).DistanceTo(bottom_point) < 1e-6 or c.GetEndPoint(1).DistanceTo(bottom_point) < 1e-6):
            if  is_line_in_plane(c, front_face.GetSurface()) and abs(c.Direction.Normalize().Z) < 0.9:
                haunch_vector = c.Direction.Normalize()
                haunch_curve  = c
       
    topHaunchPoint = haunch_curve.GetEndPoint(0) if haunch_curve.GetEndPoint(0).DistanceTo(bottom_point) < 1e-6 else haunch_curve.GetEndPoint(1)
    botHaunchPoint = haunch_curve.GetEndPoint(0) if haunch_curve.GetEndPoint(0).DistanceTo(bottom_point) > 1e-6 else haunch_curve.GetEndPoint(1)

    if line.Direction.Z < 0:
        botHaunchSlope = Cul_Haunch_Bot_D / Cul_Haunch_Bot_W
        V = XYZ(0,0,1)
        if Exterior_face:
            side_offset = Cell_Wall_T_Ext -cover
            mark = Bottom_Haunch_bars_Abutment
        else:
            side_offset = Cell_Wall_T_Int -cover
            mark = Bottom_Haunch_bars_Pier
        rebarP2 = botHaunchPoint - botSlabHeight*V - (botSlabHeight/botHaunchSlope)*U  + cover*U*2 + cover*V 
        rebarP1 = rebarP2 - 50*barSize*U  
        rebarP3 = topHaunchPoint + side_offset*U + (side_offset * botHaunchSlope )*V   - cover*V*2
        rebarP4 = rebarP3 + 50*barSize*V  

    elif line.Direction.Z > 0:
        botHaunchSlope = Cul_Haunch_Bot_D / Cul_Haunch_Bot_W
        V = XYZ(0,0,1)
        U = -U
        if Exterior_face:
            side_offset = Cell_Wall_T_Ext -cover
            mark = Bottom_Haunch_bars_Abutment
        else:
            side_offset = Cell_Wall_T_Int -cover
            mark = Bottom_Haunch_bars_Pier

        rebarP2 = botHaunchPoint - botSlabHeight*V - (botSlabHeight/botHaunchSlope)*U   + cover*U*2  + cover*V
        rebarP1 = rebarP2 - 50*barSize*U  
        rebarP3 = topHaunchPoint + side_offset*U + (side_offset * botHaunchSlope )*V  - cover*V*2
        rebarP4 = rebarP3 + 50*barSize*V  



    Plane = front_face.GetSurface()
    rebar = sc48_by4points(rebarP1, rebarP2, rebarP3, rebarP4, Plane, element, barType16, doc, sc_48, False)
    rebar.LookupParameter("Mark").Set(mark)
    # rebar = Structure.Rebar.CreateFromRebarShape(doc, sc_48, barType16, element, bottom_point, vertical_vector, horizontal_vector)
    rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(Spacing , 5,True , True, True) #SetLayoutAsMaximumSpacing(double spacing,double arrayLength,bool barsOnNormalSide,bool includeFirstBar,bool includeLastBar)
    shape_accessor = rebar.GetShapeDrivenAccessor()
    # shape_accessor.UseRebarConstraintsToProduceVaryingBars = True
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
                    rvecline = c
                    rebarVecZ = abs(c.GetEndPoint(0).Z - c.GetEndPoint(1).Z)
                    rebarVecPlanLength = (XYZ(c.GetEndPoint(0).X - c.GetEndPoint(1).X, c.GetEndPoint(0).Y - c.GetEndPoint(1).Y, 0)).GetLength()


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

    # draw_line_in_view(rvecline.GetEndPoint(0), rvecline.GetEndPoint(1))
    #get normal of Plane  going through bottom_point
    faceVec = Plane.Normal.Normalize()
    facePt =  bottom_point + faceVec*5000
    PlaneNormalLine = Line.CreateBound(bottom_point, facePt)
    faceVecPoint = line_plane_intersection(PlaneNormalLine, back_face.GetSurface())
    faceVec = (faceVecPoint - bottom_point).Normalize()


    # draw_line_in_view(transStart, transEnd)
    
    for bar in range(rebar.NumberOfBarPositions):
        transStart = bottom_point + faceVec * (bar) * Spacing 
        transEnd = bottom_point + rebarVec * (bar) * Spacing
        bar_transform = Transform.CreateTranslation( transEnd - transStart)
        rebar.MoveBarInSet(bar, bar_transform)
        # draw_line_in_view(transStart, transEnd)

##################################################################################################################################################################################################################################################################
##################################################################################################################################################################################################################################################################

# top haunch bars placed top of rebarAnchorLine

for line in rebarAnchorLine:
    top_point = get_top_point_of_line(line)

    distToLeft = (top_point - project_point_to_plane(top_point, left_face.GetSurface())).GetLength()
    distToRight = (top_point - project_point_to_plane(top_point, right_face.GetSurface())).GetLength()
    if distToLeft < 4 or distToRight < 4:
        Exterior_face = True
    else:        
        Exterior_face = False

    V = -line.Direction.Normalize()
    U = horVec.Normalize()

    for c in curve_candidates:
        if (c.GetEndPoint(0).DistanceTo(top_point) < 1e-6 or c.GetEndPoint(1).DistanceTo(top_point) < 1e-6):
            if  is_line_in_plane(c, front_face.GetSurface()) and abs(c.Direction.Normalize().Z) < 0.9:
                haunch_vector = c.Direction.Normalize()
                haunch_curve  = c
       
    topHaunchPoint = haunch_curve.GetEndPoint(0) if haunch_curve.GetEndPoint(0).DistanceTo(top_point) > 1e-6 else haunch_curve.GetEndPoint(1)
    botHaunchPoint = haunch_curve.GetEndPoint(0) if haunch_curve.GetEndPoint(0).DistanceTo(top_point) < 1e-6 else haunch_curve.GetEndPoint(1) 
                        
    if line.Direction.Z < 0:
        topHaunchSlope = Cul_Haunch_Bot_D / Cul_Haunch_Bot_W
        V = XYZ(0,0,1)
        if Exterior_face:
            side_offset = Cell_Wall_T_Ext -cover
            mark = Top_Haunch_bars_Abutment
        else:
            side_offset = Cell_Wall_T_Int -cover
            mark = Top_Haunch_bars_Pier
            
        rebarP2 = topHaunchPoint + topSlabHeight*V    - cover*V  - (topSlabHeight/topHaunchSlope)*U  - cover*U
        rebarP1 = rebarP2 - 50*barSize*U  
        rebarP3 = botHaunchPoint + side_offset*U - (side_offset * topHaunchSlope )*V   + cover*V*4
        rebarP4 = rebarP3 - 50*barSize*V  

        #sketch line to test the points
        # sketch_plane = SketchPlane.Create(doc, Plane.CreateByThreePoints(rebarP1,rebarP2, rebarP3))
        # doc.Create.NewModelCurve(Line.CreateBound(rebarP1, rebarP2), sketch_plane)    
        # doc.Create.NewModelCurve(Line.CreateBound(rebarP2, rebarP3), sketch_plane)    
        # doc.Create.NewModelCurve(Line.CreateBound(rebarP3, rebarP4), sketch_plane)    

    elif line.Direction.Z > 0:
        topHaunchSlope = Cul_Haunch_Bot_D / Cul_Haunch_Bot_W
        V = XYZ(0,0,1)
        U = -U
        if Exterior_face:
            side_offset = Cell_Wall_T_Ext -cover
            mark = Top_Haunch_bars_Abutment
        else:
            side_offset = Cell_Wall_T_Int -cover
            mark = Top_Haunch_bars_Pier

        rebarP2 = topHaunchPoint + topSlabHeight*V    - cover*V  - (topSlabHeight/topHaunchSlope)*U  - cover*U
        rebarP1 = rebarP2 - 50*barSize*U  
        rebarP3 = botHaunchPoint + side_offset*U - (side_offset * topHaunchSlope )*V   + cover*V*4
        rebarP4 = rebarP3 - 50*barSize*V  

        #sketch line to test the points
        # sketch_plane = SketchPlane.Create(doc, Plane.CreateByThreePoints(rebarP1,rebarP2, rebarP3))
        # doc.Create.NewModelCurve(Line.CreateBound(rebarP3, rebarP2), sketch_plane)    
        # doc.Create.NewModelCurve(Line.CreateBound(rebarP2, rebarP1), sketch_plane)    
        # doc.Create.NewModelCurve(Line.CreateBound(rebarP3, rebarP4), sketch_plane)    


    Plane = front_face.GetSurface()
    rebar = sc48_by4points(rebarP1, rebarP2, rebarP3, rebarP4, Plane, element, barType16, doc, sc_48, False)
    rebar.LookupParameter("Mark").Set(mark)

    # rebar = Structure.Rebar.CreateFromRebarShape(doc, sc_48, barType16, element, bottom_point, vertical_vector, horizontal_vector)
    rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(Spacing , 5,True , True, True) #SetLayoutAsMaximumSpacing(double spacing,double arrayLength,bool barsOnNormalSide,bool includeFirstBar,bool includeLastBar)
    shape_accessor = rebar.GetShapeDrivenAccessor()
    # shape_accessor.UseRebarConstraintsToProduceVaryingBars = True
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
                    rvecline = c
                    rebarVecZ = abs(c.GetEndPoint(0).Z - c.GetEndPoint(1).Z)
                    rebarVecPlanLength = (XYZ(c.GetEndPoint(0).X - c.GetEndPoint(1).X, c.GetEndPoint(0).Y - c.GetEndPoint(1).Y, 0)).GetLength()


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

    # draw_line_in_view(rvecline.GetEndPoint(0), rvecline.GetEndPoint(1))
    #get normal of Plane  going through bottom_point
    faceVec = Plane.Normal.Normalize()
    facePt =  bottom_point + faceVec*500
    PlaneNormalLine = Line.CreateBound(bottom_point, facePt)
    faceVecPoint = line_plane_intersection(PlaneNormalLine, back_face.GetSurface())
    faceVec = (faceVecPoint - bottom_point).Normalize()


    # draw_line_in_view(transStart, transEnd)
    
    for bar in range(rebar.NumberOfBarPositions):
        transStart = bottom_point + faceVec * (bar) * Spacing 
        transEnd = bottom_point + rebarVec * (bar) * Spacing
        bar_transform = Transform.CreateTranslation( transEnd - transStart)
        rebar.MoveBarInSet(bar, bar_transform)
        # draw_line_in_view(transStart, transEnd)


##################################################################################################################################################################################################################################################################
##################################################################################################################################################################################################################################################################

set_rebar_dim_to_avg(Top_Haunch_bars_Abutment)
set_rebar_dim_to_avg(Bottom_Haunch_bars_Abutment)
set_rebar_dim_to_avg(Top_Haunch_bars_Pier)
set_rebar_dim_to_avg(Bottom_Haunch_bars_Pier)

print("#"*100)

# supress warnings ####################################################################
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SupressWarnings())
t.SetFailureHandlingOptions(failHandler)

t.Commit()

print("#"*100)
print("DONE")
print("#"*100)