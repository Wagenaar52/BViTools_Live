# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import * 
import Functions as func

from pyrevit import revit, forms, script

import clr
clr.AddReference("System")
from System import Int64
from System.Collections.Generic import List

doc = revit.doc
uidoc = revit.uidoc




# doc = __revit__.ActiveUIDocument.Document
# uidoc = __revit__.ActiveUIDocument
# view = doc.ActiveView

cover = 40 * 0.00328084
hSlab = 200 * 0.00328084

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
sc_62 = getRebarShapeByName("62")
barType16 = barTypeBySize("Y16") # 10mm rebar
barSize = 10 * 0.00328084 # 10mm rebar

ww = doc.GetElement(ElementId(Int64(503255)))
################################################################################################
#   Create free from rebar


options = Options()
geometry = ww.get_Geometry(options)
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


for gobj in geometry:
    collect_curves_from_geo(gobj)
    if isinstance(gobj, GeometryInstance):
        for iobj in gobj.GetInstanceGeometry():
            #collect_curves_from_geo(iobj)
            collect_faces_from_geo(iobj)

# print(len(curve_candidates))
print(len(face_candidates))



# face = max(face_candidates, key=lambda c: c.Area)
# for f in face_candidates:
#     print(f.Area)
#     if f.Area == 11.727589897441327:
#         print("Found face with area 1")
#         face = f
#         break
# for curve in face.GetEdgesAsCurveLoops()[0]:
#     curve_candidates.append(curve)
# # print(len(curve_candidates))
# curve = max(curve_candidates, key=lambda c: c.Length)
# curve_loop = []

# for crv in face.GetEdgesAsCurveLoops()[0]:
#     curve_loop.append(crv)

for f in face_candidates:
    if f.Id == 174:
        face = f

t = Transaction(doc, "Create Free Form Rebar")
t.Start()

# Define origin point
origin = XYZ(0, 0, 0)

#draw model lines for visualisation
# for curve in curve_loop:
#     geomPlane = Plane.CreateByThreePoints(curve.GetEndPoint(0), curve.GetEndPoint(1), origin)
#     doc.Create.NewModelCurve(curve, SketchPlane.Create(doc, geomPlane))

# rebar = doc.GetElement(ElementId(Int64(497442)))
rebar = doc.GetElement(ElementId(Int64(503265)))


#create surface to constrain the rebar to
topPlane = Plane.CreateByNormalAndOrigin(XYZ(-0.231387604, -0.061820405, 0.970895470), XYZ(-23.670230325, 520.177325410, 131.595109247))  # Normal vector and origin point



plane = face.GetSurface()


ref = Reference(ww)

#Ilist of references in ww
# ref = List[Reference]()

#set rebar constraint handles to the curve loop
handleList = rebar.GetRebarConstraintsManager().GetAllHandles()
conman = rebar.GetRebarConstraintsManager()
for handle in handleList:
    print(handle.GetHandleName())
    # if handle.GetHandleName() == "End of Bar":
    #     print("$$$$")
    #     print(handle.GetHandleSurface())
    #     print(conman.GetCurrentConstraintOnHandle(handle))
    #     constraint = conman.GetCurrentConstraintOnHandle(handle)
    #     if constraint and constraint.IsToCover():
    #         new_offset = -40 * 0.00328084  # -40 mm
    #         constraint.SetDistanceToTargetCover(new_offset)
    #         conman.SetPreferredConstraint(constraint)
    #         doc.Regenerate()
    #         print("Updated:", constraint.GetDistanceToTargetCover())
    #     else:
    #         print("Current constraint is not ToCover")
    # if handle.GetHandleName() == "Bar Segment 2":
    constraint = conman.GetCurrentConstraintOnHandle(handle)
    cons = RebarConstraint.CreateConstraintToSurface(handle, plane)
    if handle.GetHandleName() == "Bar Segment 2":
        if constraint and constraint.IsToCover():
            new_offset = -40 * 0.00328084  # -40 mm
            # constraint.SetDistanceToTargetCover(new_offset)
            cons = RebarConstraint.CreateConstraintToSurface(handle, plane)
            conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
            print("Updated:", constraint.GetDistanceToTargetCover())
        else: 
            new_offset = -40 * 0.00328084  
            cons = RebarConstraint.CreateConstraintToSurface(handle, plane)
            # constraint.SetDistanceToTargetHostFace(new_offset)
            # conman.SetPreferredConstraint(constraint)
            conman.SetPreferredConstraint(cons)
            print("^&^&^")
            doc.Regenerate()
            # else:
            #     new_offset = -40 * 0.00328084  # -40 mm
            #     constraint.SetDistanceToTargetHostFace(new_offset)
            #     conman.SetPreferredConstraint(constraint)
            #     print("*&*&*")
            #     doc.Regenerate()

    print("#"*10)   



    # constraint = conman.GetCurrentConstraintOnHandle(handle)
    # cons = RebarConstraint.CreateConstraintToSurface(handle, face.GetSurface())
    if handle.GetHandleName() == "Bar Segment 2":
        constraint = conman.GetCurrentConstraintOnHandle(handle)
        constraint = conman.GetConstraintCandidatesForHandle(handle,element.Id)[0]
        print("Current constraint type:", constraint)
        print("Is ToCover constraint:",  constraint.IsToCover())
        print("Is ToHostFace constraint:", constraint.IsToHostFaceOrCover())
        if constraint and constraint.IsToCover():
            new_offset = -40 * 0.00328084  # -40 mm
            constraint.SetDistanceToTargetCover(new_offset)
            # cons = RebarConstraint.CreateConstraintToSurface(handle, plane)
            # conman.SetPreferredConstraint(constraint)
            doc.Regenerate()
            print("Updated:", constraint.GetDistanceToTargetCover())
        elif constraint and constraint.IsToHostFaceOrCover(): 
            new_offset = -40 * 0.00328084  
            constraint.SetDistanceToTargetHostFace(new_offset)
            # cons = RebarConstraint.CreateConstraintToSurface(handle, plane)
            # constraint.SetDistanceToTargetHostFace(new_offset)
            # conman.SetPreferredConstraint(constraint)
            # conman.SetPreferredConstraint(cons)
            doc.Regenerate()
            print("^&^&^")
            doc.Regenerate()


# rebar = Rebar.CreateFreeForm(doc,  barType16, ww, curve_loops, RebarStyle.Standard)


# supress warnings ####################################################################
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SupressWarnings())
t.SetFailureHandlingOptions(failHandler)

t.Commit()



print("#"*100)
print("DONE")
print("#"*100)