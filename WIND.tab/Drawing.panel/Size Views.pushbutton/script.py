from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import * 
import Functions as func
from pyrevit import revit, forms, script
import clr
import math
clr.AddReference("System")
from System import Int64
from System.Collections.Generic import List
uidoc   = __revit__.ActiveUIDocument
app     = __revit__.Application
doc     = __revit__.ActiveUIDocument.Document


def polar_to_xyz(radius, angle_degrees, z=0.0):
    angle_radians = math.radians(angle_degrees)
    x = radius * math.cos(angle_radians)
    y = radius * math.sin(angle_radians)
    return XYZ(x, y, z)

def try_set_parameter(parameter, value=None, value_string=None):
    if parameter is None or parameter.IsReadOnly:
        return False

    if value_string is not None:
        parameter.SetValueString(value_string)
    else:
        parameter.Set(value)

    return True

def drag_curve_end(curve_elem, new_point, move_end_index=1, is_section=False):
    """Move one endpoint of a line-based curve element (0=start, 1=end).
    is_section=False (plan): moves X,Y and preserves Z.
    is_section=True (section): moves X,Z and preserves Y.
    """
    loc = curve_elem.Location
    if not isinstance(loc, LocationCurve):
        raise Exception("Element {} does not have a LocationCurve.".format(curve_elem.Id))

    line = loc.Curve
    if not isinstance(line, Line):
        raise Exception("Element {} is not line-based.".format(curve_elem.Id))

    p0 = line.GetEndPoint(0)
    p1 = line.GetEndPoint(1)

    fixed = p1 if move_end_index == 1 else p0
    if is_section:
        target = XYZ(new_point.X, fixed.Y, new_point.Z)
    else:
        target = XYZ(new_point.X, new_point.Y, fixed.Z)

    if move_end_index == 1:
        loc.Curve = Line.CreateBound(p0, target)
    else:
        loc.Curve = Line.CreateBound(target, p1)

def draw_line_in_view(doc, start, end):
    """Draw a model line in the given view between start and end points."""
    line = Line.CreateBound(start, end)
    sketch_plane = SketchPlane.Create(doc, Plane.CreateByThreePoints(start, end, XYZ(0,0,1)))
    return doc.Create.NewModelCurve(line, sketch_plane)

def get_location_y(element):
    """Return element Y from LocationPoint or LocationCurve."""
    loc = element.Location
    if isinstance(loc, LocationPoint):
        return loc.Point.Y
    if isinstance(loc, LocationCurve):
        return loc.Curve.Evaluate(0.5, True).Y
    raise Exception("Element {} has unsupported Location type.".format(element.Id))

def setup_Crop_box(cropView , boxMin_X, boxMin_Y, boxMax_X, boxMax_Y, secDepth):

    #initialize a new bounding box for the crop box
    box1 = BoundingBoxXYZ()

    view_dir = cropView.ViewDirection
    view_origin = cropView.Origin   

    x_min = min(boxMin_X, boxMax_X)
    x_max = max(boxMin_X, boxMax_X)
    y_min = min(boxMin_Y, boxMax_Y)
    y_max = max(boxMin_Y, boxMax_Y)
    safe_depth = max(secDepth, 1e-6)

    box1.Min = XYZ(x_min - view_origin.X, y_min - view_origin.Z, 0 - view_origin.Y)
    box1.Max = XYZ(x_max - view_origin.X, y_max - view_origin.Z, safe_depth - view_origin.Y)

    Trans = Transform.Identity
    forward = view_dir.Normalize()
    right = XYZ(0,0,1).CrossProduct(forward).Normalize()
    up = forward.CrossProduct(right).Normalize()

    Trans.BasisX = right  
    Trans.BasisY = up  
    Trans.BasisZ = forward 
    box1.Transform = Trans
    cropView.CropBox = box1

def setup_Section_box(view, boxMin_X, boxMax_X, boxMin_Y, boxMax_Y, boxMin_Z, boxMax_Z):
    """Set up a section box for a 3D view."""
    if isinstance(view, View3D):
        newBB = BoundingBoxXYZ()
        newBB.Transform = Transform.Identity   # Min/Max are now in world coords
        newBB.Min = XYZ(boxMin_X, boxMin_Y, boxMin_Z)
        newBB.Max = XYZ(boxMax_X, boxMax_Y, boxMax_Z)
        view.SetSectionBox(newBB)

def set_cropbox_from_project_coords(view, model_min, model_max):
    """Set CropBox using project (global/internal) coordinates."""
    crop = view.CropBox
    inv = crop.Transform.Inverse

    p0 = inv.OfPoint(model_min)
    p1 = inv.OfPoint(model_max)

    crop.Min = XYZ(min(p0.X, p1.X), min(p0.Y, p1.Y), min(p0.Z, p1.Z))
    crop.Max = XYZ(max(p0.X, p1.X), max(p0.Y, p1.Y), max(p0.Z, p1.Z))
    view.CropBox = crop

sheets = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).WhereElementIsNotElementType()
views = FilteredElementCollector(doc).OfClass(View).WhereElementIsNotElementType()

#prompt user to select wheter Radial plan views should be moved to align breaklines or not the difine and set the variable MoveViews to True or False accordingly
try:
    if hasattr(forms, "ask_for_bool"):
        MoveViews = forms.ask_for_bool("Move radial plan views on DWG1 to align breaklines?", default=False)
    else:
        MoveViews = forms.alert(
            "Move radial plan views on DWG1 to align breaklines?",
            title="Align Radial Plan Views",
            yes=True,
            no=True,
            ok=False
        )
except Exception:
    MoveViews = False



for s in sheets:
    if s.Id == ElementId(426778):
        GA_sheet = s

for vId in GA_sheet.GetAllViewports():
    vp = doc.GetElement(vId)
    v = doc.GetElement(vp.ViewId)
    if v.Id == ElementId(137083):
        GA_planView = v
        GA_planViewport = vp
    elif v.Id == ElementId(453188):
        GA_secAview = v
        GA_secAviewport = vp
    elif v.Id == ElementId(5041804):
        GA_secBview = v
        GA_secBviewport = vp
    elif v.Id == ElementId(5226711):
        GA_secCview = v
        GA_secCviewport = vp
    elif v.Id == ElementId(3042702):
        GA_coverview = v
        GA_coverviewport = vp
    elif v.Id == ElementId(3296029):
        GA_tfTopview = v
        GA_tfTopviewport = vp
    elif v.Id == ElementId(3296191):
        GA_tfBotview = v
        GA_tfBotviewport = vp
    elif v.Id == ElementId(4509296):
        GA_ISOleftview = v
        GA_ISOleftviewport = vp
    elif v.Id == ElementId(2500439):
        GA_ISOrightview = v
        GA_ISOrightviewport = vp
    elif v.Id == ElementId(6009381):
        GA_radBaseview = v
        GA_radBaseviewport = vp


FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType()
for e in FEC:
    if e.Name == "1PA_WTF_Blinding":
        WTF_blinding = e
    elif e.Name == "1PA_WTF_SteelTower":
        WTF_foundation = e
    elif e.Name == "1PA_WTF_Blinging":
        WTF_blinging = e
    elif e.Name == "1PA_AnchorCage_Assembly":
        WTF_ac = e
    elif e.Name == "1PA_WTF_BottomTowerSectionWithDoor":
        WTF_Tower = e
    elif e.Name == "1PA_WTF_Backfill":
        WTG_Backfill = e

rBase               = WTF_blinding.LookupParameter("rBase").AsDouble()
rBlindingExtension  = WTF_blinding.LookupParameter("rBlindingExtension").AsDouble()
hPlinth             = WTF_foundation.LookupParameter("hPlinth").AsDouble()
rTower              = WTF_foundation.LookupParameter("rTower").AsDouble()
rPlinth             = WTF_foundation.LookupParameter("rPlinth").AsDouble()
wGroutTop           = WTF_foundation.LookupParameter("wGroutTop").AsDouble()
hBottomVoid         = WTF_foundation.LookupParameter("hBottomVoid").AsDouble()
rVoidOuter         = WTF_foundation.LookupParameter("rVoidOuter").AsDouble()
hBlinding           = WTF_blinding.LookupParameter("hBlinding").AsDouble()
wFlangeBot          = WTF_ac.LookupParameter("wFlangeBot").AsDouble()   
TOflange_TOtower    = WTF_ac.LookupParameter("ovTOCtoBOTtopFlange").AsDouble() + WTF_ac.LookupParameter("tBearingPlate").AsDouble()  + WTF_ac.LookupParameter("tFlangeTop").AsDouble() + WTF_ac.LookupParameter("hShell").AsDouble()
tShell              = WTF_ac.LookupParameter("tShell").AsDouble()   
htower              = WTF_Tower.LookupParameter("hDoorCentre").AsDouble() + WTF_Tower.LookupParameter("hDoorOut").AsDouble()/2 +1
hPit                = WTF_foundation.LookupParameter("hBottomVoid").AsDouble()
rGrout              = WTF_foundation.LookupParameter("wGroutTop").AsDouble()/2 + WTF_foundation.LookupParameter("rTower").AsDouble()
hGrout              = WTF_foundation.LookupParameter("dGroutMiddle").AsDouble() 
hCone               = WTF_foundation.LookupParameter("hCone").AsDouble()

#region Resize crop box and move breaklines to match new crop box - PLAN VIEW

rad = rBase + rBlindingExtension*2
topBreakline = doc.GetElement(ElementId(8171664))
botBreakline = doc.GetElement(ElementId(5016982))

top_y = get_location_y(topBreakline)
bot_y = get_location_y(botBreakline)

newCropBox = BoundingBoxXYZ()
newCropBox.Min = XYZ(-rad, bot_y - rBlindingExtension, -rad)
newCropBox.Max = XYZ(rad, top_y + rBlindingExtension, rad)

t = Transaction(doc, "Update 01-232 Crop - Plan")
t.Start()

GA_planView.CropBox = newCropBox

final_point0 = XYZ(-rad*math.sin(math.radians(45)), rad*math.cos(math.radians(45)), 0)
final_point1 = XYZ(rad*math.sin(math.radians(45)) , rad*math.cos(math.radians(45)), 0)

drag_curve_end(topBreakline, final_point0, move_end_index=0)
drag_curve_end(topBreakline, final_point1, move_end_index=1)

t.Commit()

#endregion

#region Resize crop box and move breaklines to match new crop box - SEC A VIEW

boxMin_X = -(rBase + rBlindingExtension*2)
boxMax_X = rBase + rBlindingExtension*2
boxMin_Y = -3000/304.8
boxMax_Y = hPlinth + 3000/304.8
secDepth = 0.01

t = Transaction(doc, "Edit Sheet Viewports")
t.Start()

setup_Crop_box(GA_secAview, boxMin_X, boxMin_Y, boxMax_X, boxMax_Y, secDepth)

t.Commit()

#endregion

#region Resize crop box and move breaklines to match new crop box - 01-232 TOWER FLANGE DETAIL BOTTOM
p1 = XYZ(rTower -2 ,  0 , -hBottomVoid)
p2 = XYZ(rTower +2 ,  0.5 ,  hBottomVoid)

min_pt = XYZ(min(p1.X, p2.X), min(p1.Y, p2.Y), min(p1.Z, p2.Z))
max_pt = XYZ(max(p1.X, p2.X), max(p1.Y, p2.Y), max(p1.Z, p2.Z))

t = Transaction(doc, "Update 01-232 Crop")
t.Start()

boxMin_X = rTower -wFlangeBot/2 - 0.5
boxMin_Y = -hBottomVoid - hBlinding/2
boxMax_X = rTower +wFlangeBot/2 +0.5
boxMax_Y = hBottomVoid + 0.5
secDepth = 1

setup_Crop_box(GA_tfBotview, boxMin_X, boxMin_Y, boxMax_X, boxMax_Y, secDepth)

#place breaklines at the edge of the crop box
topBL = doc.GetElement(ElementId(3296216))
leftBL = doc.GetElement(ElementId(3296213))
rightBL = doc.GetElement(ElementId(3296215))
botBL = doc.GetElement(ElementId(4267364))

L0 = XYZ(boxMin_X +0.15, 0, boxMin_Y +0.15)   
R0 = XYZ(boxMax_X -0.15, 0, boxMin_Y +0.15)
R1 = XYZ(boxMax_X -0.15, 0, boxMax_Y -0.15) 
L1 = XYZ(boxMin_X +0.15, 0, boxMax_Y -0.15)

drag_curve_end(topBL, L1, move_end_index=0, is_section=True)
drag_curve_end(topBL, R1, move_end_index=1, is_section=True)
drag_curve_end(topBL, L1, move_end_index=0)
drag_curve_end(topBL, R1, move_end_index=1)

drag_curve_end(botBL, L0, move_end_index=0, is_section=True)
drag_curve_end(botBL, R0, move_end_index=1, is_section=True)

drag_curve_end(rightBL, R0, move_end_index=0)
drag_curve_end(rightBL, R1, move_end_index=1)

drag_curve_end(leftBL, L0, move_end_index=0)
drag_curve_end(leftBL, L1, move_end_index=1)

t.Commit()
#endregion

#region Resize crop box and move breaklines to match new crop box - 01-232 TOWER FLANGE DETAIL TOP

t = Transaction(doc, "Update 01-232 Crop - Top")
t.Start()

boxMin_X = rTower -wGroutTop/2 - 0.5
boxMax_X = rPlinth + 1
boxMin_Y = hPlinth - 4
boxMax_Y = hPlinth + TOflange_TOtower - 0.1
secDepth = 0.01

setup_Crop_box(GA_tfTopview, boxMin_X, boxMin_Y, boxMax_X, boxMax_Y, secDepth)

#place breaklines at the edge of the crop box
topBL = doc.GetElement(ElementId(3297481))
leftBL = doc.GetElement(ElementId(3296048))
rightBL = doc.GetElement(ElementId(3296046))
botBL = doc.GetElement(ElementId(3296038))

L0 = XYZ(boxMin_X +0.15, 0, boxMin_Y +0.15)   
R0 = XYZ(boxMax_X -0.15, 0, boxMin_Y +0.15)
R1 = XYZ(boxMax_X -0.15, 0, boxMax_Y -0.15) 
L1 = XYZ(boxMin_X +0.15, 0, boxMax_Y -0.15)

drag_curve_end(topBL, XYZ(rTower - tShell, L1.Y, L1.Z), move_end_index=1, is_section=True)
drag_curve_end(topBL, XYZ(rTower + tShell, L1.Y, L1.Z), move_end_index=0, is_section=True)
drag_curve_end(topBL, XYZ(rTower - tShell, L1.Y, L1.Z), move_end_index=1)
drag_curve_end(topBL, XYZ(rTower + tShell, L1.Y, L1.Z), move_end_index=0)


drag_curve_end(botBL, L0, move_end_index=0, is_section=True)
drag_curve_end(botBL, R0, move_end_index=1, is_section=True)
drag_curve_end(botBL, L0, move_end_index=0)
drag_curve_end(botBL, R0, move_end_index=1)

drag_curve_end(rightBL, R0, move_end_index=0, is_section=True)
drag_curve_end(rightBL, XYZ(R1.X, R1.Y, hPlinth), move_end_index=1, is_section=True)

drag_curve_end(leftBL, L0, move_end_index=0, is_section=True)
drag_curve_end(leftBL, XYZ(L1.X, L1.Y, hPlinth+0.1), move_end_index=1, is_section=True)

t.Commit()
#endregion

#region Resize section box acording to WTF_Foundation- GA_ISOleftview

_bb = WTG_Backfill.get_BoundingBox(None)
boxMin_X = _bb.Min.X - 1
boxMax_X = _bb.Max.X + 1
boxMin_Y = _bb.Min.Y - 1
boxMax_Y = _bb.Max.Y + 1
boxMin_Z = -hPit - 1
boxMax_Z = hPlinth + htower + 1


t = Transaction(doc, "Edit Sheet Viewports")
t.Start()

isoLeftView = doc.GetElement(ElementId(4509296))
if isinstance(isoLeftView, View3D):
    newSecBB = BoundingBoxXYZ()
    newSecBB.Transform = Transform.Identity   # Min/Max are now in world coords
    newSecBB.Min = XYZ(boxMin_X, boxMin_Y, boxMin_Z)
    newSecBB.Max = XYZ(0, boxMax_Y, boxMax_Z)
    isoLeftView.SetSectionBox(newSecBB)

t.Commit()

#endregion

#region Resize section box acording to WTF_Foundation- GA_ISOrightview

_bb = WTG_Backfill.get_BoundingBox(None)
boxMin_X = _bb.Min.X - 1
boxMax_X = _bb.Max.X + 1
boxMin_Y = _bb.Min.Y - 1
boxMax_Y = _bb.Max.Y + 1
boxMin_Z = -hPit - 1
boxMax_Z = hPlinth + htower + 1


t = Transaction(doc, "Edit Sheet Viewports")
t.Start()

isoRightView = doc.GetElement(ElementId(2500439))
if isinstance(isoRightView, View3D):
    newSecBB = BoundingBoxXYZ()
    newSecBB.Transform = Transform.Identity   # Min/Max are now in world coords
    newSecBB.Min = XYZ(0, 0, boxMin_Z)
    newSecBB.Max = XYZ(boxMax_X, boxMax_Y, boxMax_Z)
    isoRightView.SetSectionBox(newSecBB)

t.Commit()

#endregion

#region Resize crop box and move breaklines to match new crop box - SECTION C
# t = Transaction(doc, "Update 01-232 Crop - Section C")
# t.Start()

# boxMin_X = -rPlinth 
# boxMax_X = rPlinth*0.75
# boxMin_Y = hPlinth - 4
# boxMax_Y = hPlinth + htower
# secDepth = 800/304.8

# setup_Crop_box(GA_secCview, boxMin_X, boxMin_Y, boxMax_X, boxMax_Y, secDepth)

# #place breaklines at the edge of the crop box
# topBL = doc.GetElement(ElementId(5231705))
# leftBL = doc.GetElement(ElementId(5234988))
# rightBL = doc.GetElement(ElementId(5235080))
# botBL = doc.GetElement(ElementId(5235045))

# L0 = XYZ(boxMin_X +0.15, 0, boxMin_Y +0.15)   
# R0 = XYZ(boxMax_X -0.15, 0, boxMin_Y +0.15)
# R1 = XYZ(boxMax_X -0.15, 0, boxMax_Y -0.15) 
# L1 = XYZ(boxMin_X +0.15, 0, boxMax_Y -0.15)

# drag_curve_end(topBL, XYZ(-rTower -1, L1.Y, L1.Z), move_end_index=0, is_section=True)
# drag_curve_end(topBL, R1, move_end_index=1, is_section=True)
# drag_curve_end(topBL, XYZ(-rTower -1, L1.Y, L1.Z), move_end_index=0)
# drag_curve_end(topBL, R1, move_end_index=1)

# drag_curve_end(botBL, L0, move_end_index=0, is_section=True)
# drag_curve_end(botBL, R0, move_end_index=1, is_section=True)
# drag_curve_end(botBL, L0, move_end_index=0)
# drag_curve_end(botBL, R0, move_end_index=1)

# drag_curve_end(rightBL, R0, move_end_index=0, is_section=True)
# drag_curve_end(rightBL, R1, move_end_index=1, is_section=True)

# drag_curve_end(leftBL, L0, move_end_index=1, is_section=True)
# drag_curve_end(leftBL, XYZ(L1.X, L1.Y, hPlinth+0.1), move_end_index=0, is_section=True)

# t.Commit()
#endregion

################################################################################################################################################################################################################################################################################################################################################################################################################################################################
#region #ANCHOR CAGE################################################################################################################################################################################################################################################################################################################################################################################################################################################################
################################################################################################################################################################################################################################################################################################################################################################################################################################################################

#Resize crop box and move breaklines to match new crop box - ANCHOR CAGE

boxMin_X = -rGrout - 1
boxMax_X = +rGrout + 1
boxMin_Y = -hPit - 1
boxMax_Y = hPlinth + TOflange_TOtower + 1
secDepth = 0.01

t = Transaction(doc, "Edit Sheet Viewports")
t.Start()

################################################################################################################################################################################################################################################################################################################################################################################################################################################################
ISO_VIEW_OF_ANCHOR_CAGE = doc.GetElement(ElementId(545480))
boxMin_X = -rGrout - 1
boxMax_X = +rGrout + 1
boxMin_Y = -rGrout - 1
boxMax_Y = 0
boxMin_Z = -hPit - 1
boxMax_Z = hPlinth + TOflange_TOtower + 1
if isinstance(ISO_VIEW_OF_ANCHOR_CAGE, View3D):
    newBB = BoundingBoxXYZ()
    newBB.Transform = Transform.Identity   # Min/Max are now in world coords
    newBB.Min = XYZ(boxMin_X, boxMin_Y, boxMin_Z)
    newBB.Max = XYZ(boxMax_X, boxMax_Y, boxMax_Z)
    ISO_VIEW_OF_ANCHOR_CAGE.SetSectionBox(newBB)

################################################################################################################################################################################################################################################################################################################################################################################################################################################################
ISO_VIEW_OF_ANCHOR_CAGE_SUPPORT_STUD = doc.GetElement(ElementId(792595))
boxMin_X = -rTower -wGroutTop/2 - 0.5
boxMax_X = -rTower +wGroutTop/2 + 0.5
boxMin_Y = -1.5
boxMax_Y = 1.5
boxMin_Z = -hPit -0.1
boxMax_Z = 1
if isinstance(ISO_VIEW_OF_ANCHOR_CAGE_SUPPORT_STUD, View3D):
    newBB = BoundingBoxXYZ()
    newBB.Transform = Transform.Identity   # Min/Max are now in world coords
    newBB.Min = XYZ(boxMin_X, boxMin_Y, boxMin_Z)
    newBB.Max = XYZ(boxMax_X, boxMax_Y, boxMax_Z)
    ISO_VIEW_OF_ANCHOR_CAGE_SUPPORT_STUD.SetSectionBox(newBB)

################################################################################################################################################################################################################################################################################################################################################################################################################################################################
SECTION_OF_ANCHOR_CAGE = doc.GetElement(ElementId(793179))
boxMin_X = -rGrout - 1
boxMax_X = +rGrout + 1
boxMin_Y = -hPit - 1
boxMax_Y = hPlinth + TOflange_TOtower + 1
secDepth = 0.01
setup_Crop_box(SECTION_OF_ANCHOR_CAGE, boxMin_X, boxMin_Y, boxMax_X, boxMax_Y, secDepth)

################################################################################################################################################################################################################################################################################################################################################################################################################################################################
PLAN_VIEW_OF_BOTTOM_ANCHOR_PLATE = doc.GetElement(ElementId(794699))
newCropBox = BoundingBoxXYZ()
newCropBox.Min = XYZ(-rGrout - 1, -rGrout - 1, 0)
newCropBox.Max = XYZ( rGrout + 1 , rGrout + 1,  hPlinth/2)
PLAN_VIEW_OF_BOTTOM_ANCHOR_PLATE.CropBox = newCropBox

################################################################################################################################################################################################################################################################################################################################################################################################################################################################
PLAN_VIEW_OF_ANCHOR_CAGE = doc.GetElement(ElementId(5700956))
newCropBox = BoundingBoxXYZ()
newCropBox.Min = XYZ(-rGrout - 1, -rGrout - 1, hPlinth/2)
newCropBox.Max = XYZ( rGrout + 1 , rGrout + 1,  hPlinth + TOflange_TOtower + 1)
PLAN_VIEW_OF_ANCHOR_CAGE.CropBox = newCropBox

################################################################################################################################################################################################################################################################################################################################################################################################################################################################
SECTION_OF_ANCHOR_CAGE_CALL1 = doc.GetElement(ElementId(3597517))
boxMin_X = rTower -wGroutTop/2 - 0.5
boxMax_X = rTower +wGroutTop/2 + 0.5
boxMin_Y = hPlinth - hGrout - 1
boxMax_Y = hPlinth + TOflange_TOtower - 0.1
secDepth = 0.01
setup_Crop_box(SECTION_OF_ANCHOR_CAGE_CALL1, boxMin_X, boxMin_Y, boxMax_X, boxMax_Y, secDepth)

################################################################################################################################################################################################################################################################################################################################################################################################################################################################
SECTION_OF_ANCHOR_CAGE_CALL2 = doc.GetElement(ElementId(3597532))
SECTION_OF_ANCHOR_CAGE_CALL2_sec = doc.GetElement(ElementId(3597531))
boxMin_X = rTower -wFlangeBot/2 - 0.5
boxMin_Y = -hBottomVoid - hBlinding/2
boxMax_X = rTower +wFlangeBot/2 + 0.5
boxMax_Y = 1
farClip = SECTION_OF_ANCHOR_CAGE_CALL2.get_Parameter(BuiltInParameter.VIEWER_BOUND_ACTIVE_FAR)

farClipSetting = SECTION_OF_ANCHOR_CAGE_CALL2.LookupParameter("Far Clip Settings")
if try_set_parameter(farClipSetting, value_string="Same as parent view"):
    doc.Regenerate() # regenerate to update the far clip distance after changing the setting

    farClipSetting = SECTION_OF_ANCHOR_CAGE_CALL2.LookupParameter("Far Clip Settings")
    if try_set_parameter(farClipSetting, value_string="Independent"):
        doc.Regenerate() # regenerate to update the far clip distance after changing the setting

try_set_parameter(farClip, value=10/304.8)
secDepth = 10/304.8
setup_Crop_box(SECTION_OF_ANCHOR_CAGE_CALL2, boxMin_X, boxMin_Y, boxMax_X, boxMax_Y, secDepth)

################################################################################################################################################################################################################################################################################################################################################################################################################################################################
DETAIL_3 = doc.GetElement(ElementId(8646061))#8646061))
boxMin_X = -rTower -wGroutTop/2 - 0.5
boxMax_X = -rTower +wGroutTop/2 + 0.5
boxMin_Y = -hPit - 0.5
boxMax_Y = 1
secDepth = 300/304.8
setup_Crop_box(DETAIL_3, boxMin_X, boxMin_Y, boxMax_X, boxMax_Y, secDepth)

################################################################################################################################################################################################################################################################################################################################################################################################################################################################
TOWER_FLANGE_DETAIL_TOP_PRE_GROUT = doc.GetElement(ElementId(7103215))
boxMin_X = rTower -wGroutTop/2 - 0.5
boxMax_X = rTower +wGroutTop/2 + 0.5
boxMin_Y = hPlinth - hGrout - 1
boxMax_Y = hPlinth + TOflange_TOtower - 0.1
secDepth = 0.01
setup_Crop_box(TOWER_FLANGE_DETAIL_TOP_PRE_GROUT, boxMin_X, boxMin_Y, boxMax_X, boxMax_Y, secDepth)

#endregion ################################################################################################################################################################################################################################################################################################################################################################################################################################################################

#region REBAR DWG 1 ################################################################################################################################################################################################################################################################################################################################################################################################################################################################

#Plan Views
###########################################################################################################################################################
TOP_CONCENTRIC_REBAR_PLAN_VIEW = doc.GetElement(ElementId(446567))
newCropBox = BoundingBoxXYZ()
newCropBox.Min = XYZ(0,0,0)
newCropBox.Max = XYZ( rBase + 1 ,  rBase + 1,  hCone)
#create a curve loop for the crop shape
curvLoop = CurveLoop()
L1 =Line.CreateBound(XYZ(rBase + 1, rBase + 1, 0), XYZ(0, rBase + 1, 0))
L2 =Line.CreateBound(XYZ(0, rBase + 1, 0), XYZ(0, 0, 0))
L3 =Line.CreateBound(XYZ(0, 0, 0), XYZ(rBase + 1, rBase + 1, 0))
curvLoop.Append(L1)
curvLoop.Append(L2)
curvLoop.Append(L3)
TOP_CONCENTRIC_REBAR_PLAN_VIEW.GetCropRegionShapeManager().SetCropShape(curvLoop) 
###########################################################################################################################################################
FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType()
GR100_max = 0
GR100_h = 0
GR200_max = 0
GR200_h = 0
GR300_max = 0
GR300_h = 0
GR400_max = 0
GR400_h = 0

for bar in FEC:
    if "GR100" in bar.LookupParameter("Schedule Mark").AsValueString():
        if bar.LookupParameter("A").AsDouble() > GR100_max:
            GR100_max = bar.LookupParameter("A").AsDouble()
            GR100_h = bar.GetCenterlineCurves(adjustForSelfIntersection=False,suppressHooks=False,suppressBendRadius=False,multiplanarOption=0,barPositionIndex=0)
            GR100_h = GR100_h[0].GetEndPoint(1).Z
    elif "GR200" in bar.LookupParameter("Schedule Mark").AsValueString():
        if bar.LookupParameter("A").AsDouble() > GR200_max:
            GR200_max = bar.LookupParameter("A").AsDouble()
            GR200_h = bar.GetCenterlineCurves(adjustForSelfIntersection=False,suppressHooks=False,suppressBendRadius=False,multiplanarOption=0,barPositionIndex=0)
            GR200_h = GR200_h[0].GetEndPoint(1).Z
    elif "GR300" in bar.LookupParameter("Schedule Mark").AsValueString():
        if bar.LookupParameter("A").AsDouble() > GR300_max:
            GR300_max = bar.LookupParameter("A").AsDouble()
            GR300_h = bar.GetCenterlineCurves(adjustForSelfIntersection=False,suppressHooks=False,suppressBendRadius=False,multiplanarOption=0,barPositionIndex=0)
            GR300_h = GR300_h[0].GetEndPoint(0).Z
    elif "GR400" in bar.LookupParameter("Schedule Mark").AsValueString():
        if bar.LookupParameter("A").AsDouble() > GR400_max:
            GR400_max = bar.LookupParameter("A").AsDouble()
            GR400_h = bar.GetCenterlineCurves(adjustForSelfIntersection=False,suppressHooks=False,suppressBendRadius=False,multiplanarOption=0,barPositionIndex=0)
            GR400_h = GR400_h[0].GetEndPoint(1).Z

GRID_1_PLAN_VIEW = doc.GetElement(ElementId(453468))
if GRID_1_PLAN_VIEW is not None and GR100_max > 0:
    model_min = XYZ(-(GR100_max/2)-1, -(GR100_max/2)-1, GR100_h - 1)
    model_max = XYZ((GR100_max/2)+1, (GR100_max/2)+1, GR100_h + 1)
    set_cropbox_from_project_coords(GRID_1_PLAN_VIEW, model_min, model_max)

    pvr = GRID_1_PLAN_VIEW.GetViewRange()
    pvr.SetOffset(PlanViewPlane.TopClipPlane,     GR100_h + 100/304.8)
    pvr.SetOffset(PlanViewPlane.CutPlane,         GR100_h + 100/304.8)
    pvr.SetOffset(PlanViewPlane.BottomClipPlane,  GR100_h - 100/304.8)
    pvr.SetOffset(PlanViewPlane.ViewDepthPlane,   GR100_h - 100/304.8)
    GRID_1_PLAN_VIEW.SetViewRange(pvr)

###########################################################################################################################################################
GRID_2_PLAN_VIEW = doc.GetElement(ElementId(453478))
if GRID_2_PLAN_VIEW is not None and GR200_max > 0:
    model_min = XYZ(-(GR200_max/2)-1, -(GR200_max/2)-1, GR200_h - 1)
    model_max = XYZ((GR200_max/2)+1, (GR200_max/2)+1, GR200_h + 1)
    set_cropbox_from_project_coords(GRID_2_PLAN_VIEW, model_min, model_max)

    pvr = GRID_2_PLAN_VIEW.GetViewRange()
    pvr.SetOffset(PlanViewPlane.TopClipPlane,     GR200_h + 100/304.8)
    pvr.SetOffset(PlanViewPlane.CutPlane,         GR200_h + 100/304.8)
    pvr.SetOffset(PlanViewPlane.BottomClipPlane,  GR200_h - 100/304.8)
    pvr.SetOffset(PlanViewPlane.ViewDepthPlane,   GR200_h - 100/304.8)
    GRID_2_PLAN_VIEW.SetViewRange(pvr)

###########################################################################################################################################################
GRID_3_PLAN_VIEW = doc.GetElement(ElementId(9137627))
if GRID_3_PLAN_VIEW is not None and GR300_max > 0:
    model_min = XYZ(-(GR300_max/2)-1, -(GR300_max/2)-1, GR300_h - 1)
    model_max = XYZ((GR300_max/2)+1, (GR300_max/2)+1, GR300_h + 1)
    set_cropbox_from_project_coords(GRID_3_PLAN_VIEW, model_min, model_max)

    pvr = GRID_3_PLAN_VIEW.GetViewRange()
    pvr.SetOffset(PlanViewPlane.TopClipPlane,     GR300_h + 100/304.8)
    pvr.SetOffset(PlanViewPlane.CutPlane,         GR300_h + 100/304.8)
    pvr.SetOffset(PlanViewPlane.BottomClipPlane,  GR300_h - 100/304.8)
    pvr.SetOffset(PlanViewPlane.ViewDepthPlane,   GR300_h - 100/304.8)
    GRID_3_PLAN_VIEW.SetViewRange(pvr)

###########################################################################################################################################################
GRID_4_PLAN_VIEW = doc.GetElement(ElementId(4593885))
if GRID_4_PLAN_VIEW is not None and GR400_max > 0:
    model_min = XYZ(-(GR400_max/2)-1, -(GR400_max/2)-1, GR400_h - 1)
    model_max = XYZ((GR400_max/2)+1, (GR400_max/2)+1, GR400_h + 1)
    set_cropbox_from_project_coords(GRID_4_PLAN_VIEW, model_min, model_max)

    pvr = GRID_4_PLAN_VIEW.GetViewRange()
    pvr.SetOffset(PlanViewPlane.TopClipPlane,     GR400_h + 100/304.8)
    pvr.SetOffset(PlanViewPlane.CutPlane,         GR400_h + 100/304.8)
    pvr.SetOffset(PlanViewPlane.BottomClipPlane,  GR400_h - 100/304.8)
    pvr.SetOffset(PlanViewPlane.ViewDepthPlane,   GR400_h - 100/304.8)
    GRID_4_PLAN_VIEW.SetViewRange(pvr)

###########################################################################################################################################################
BOTTOM_1_CONCENTRIC_REBAR_PLAN_VIEW = doc.GetElement(ElementId(5905497))
newCropBox = BoundingBoxXYZ()
newCropBox.Min = XYZ(0,0,0)
newCropBox.Max = XYZ( rBase + 1 ,  rBase + 1,  hCone)
#create a curve loop for the crop shape
curvLoop = CurveLoop()
L1 =Line.CreateBound(XYZ(0, 0, 0), XYZ(rBase +1, 0, 0))
L2 =Line.CreateBound(XYZ(rBase +1, 0, 0), XYZ(rBase + 1, rBase + 1, 0))
L3 =Line.CreateBound(XYZ(rBase + 1, rBase + 1, 0), XYZ(0, 0, 0))
curvLoop.Append(L1)
curvLoop.Append(L2)
curvLoop.Append(L3)
BOTTOM_1_CONCENTRIC_REBAR_PLAN_VIEW.GetCropRegionShapeManager().SetCropShape(curvLoop) 

###########################################################################################################################################################
BOTTOM_3_RADIAL_REBAR_PLAN_VIEW = doc.GetElement(ElementId(5915879))
newCropBox = BoundingBoxXYZ()
newCropBox.Min = XYZ(0,0,0)
newCropBox.Max = XYZ( rBase + 1 ,  rBase + 1,  hCone)
#create a curve loop for the crop shape
curvLoop = CurveLoop()
L1 =Line.CreateBound(XYZ(0, 0, 0), polar_to_xyz(rBase +5, 22.5))
L2 =Line.CreateBound(polar_to_xyz(rBase +5, 22.5), polar_to_xyz(rBase + 5, 45))
L3 =Line.CreateBound(polar_to_xyz(rBase + 5, 45), XYZ(0, 0, 0))
curvLoop.Append(L1)
curvLoop.Append(L2)
curvLoop.Append(L3)
BOTTOM_3_RADIAL_REBAR_PLAN_VIEW.GetCropRegionShapeManager().SetCropShape(curvLoop) 

###########################################################################################################################################################
BOTTOM_1_RADIAL_REBAR_PLAN_VIEW = doc.GetElement(ElementId(453671))
newCropBox = BoundingBoxXYZ()
newCropBox.Min = XYZ(0,0,0)
newCropBox.Max = XYZ( rBase + 1 ,  rBase + 1,  hCone)
#create a curve loop for the crop shape
curvLoop = CurveLoop()
L1 =Line.CreateBound(XYZ(0, 0, 0), polar_to_xyz(rBase +5, 0))
L2 =Line.CreateBound(polar_to_xyz(rBase +5, 0), polar_to_xyz(rBase + 5, 22.5))
L3 =Line.CreateBound(polar_to_xyz(rBase + 5, 22.5), XYZ(0, 0, 0))
curvLoop.Append(L1)
curvLoop.Append(L2)
curvLoop.Append(L3)
BOTTOM_1_RADIAL_REBAR_PLAN_VIEW.GetCropRegionShapeManager().SetCropShape(curvLoop) 

###########################################################################################################################################################
TOP_2_RADIAL_REBAR_PLAN_VIEW = doc.GetElement(ElementId(2511780))
newCropBox = BoundingBoxXYZ()
newCropBox.Min = XYZ(0,0,0)
newCropBox.Max = XYZ( rBase + 1 ,  rBase + 1,  hCone)
#create a curve loop for the crop shape
curvLoop = CurveLoop()
L1 =Line.CreateBound(XYZ(0, 0, 0), polar_to_xyz(rBase +5, 45))
L2 =Line.CreateBound(polar_to_xyz(rBase +5, 45), polar_to_xyz(rBase + 5, 67.5))
L3 =Line.CreateBound(polar_to_xyz(rBase + 5, 67.5), XYZ(0, 0, 0))
curvLoop.Append(L1)
curvLoop.Append(L2)
curvLoop.Append(L3)
TOP_2_RADIAL_REBAR_PLAN_VIEW.GetCropRegionShapeManager().SetCropShape(curvLoop) 

###########################################################################################################################################################
TOP_1_RADIAL_REBAR_PLAN_VIEW = doc.GetElement(ElementId(2511791))
newCropBox = BoundingBoxXYZ()
newCropBox.Min = XYZ(0,0,0)
newCropBox.Max = XYZ( rBase + 1 ,  rBase + 1,  hCone)
#create a curve loop for the crop shape
curvLoop = CurveLoop()
L1 =Line.CreateBound(XYZ(0, 0, 0), polar_to_xyz(rBase +5, 67.5))
L2 =Line.CreateBound(polar_to_xyz(rBase +5, 67.5), polar_to_xyz(rBase + 5, 90))
L3 =Line.CreateBound(polar_to_xyz(rBase + 5, 90), XYZ(0, 0, 0))
curvLoop.Append(L1)
curvLoop.Append(L2)
curvLoop.Append(L3)
TOP_1_RADIAL_REBAR_PLAN_VIEW.GetCropRegionShapeManager().SetCropShape(curvLoop) 


#Isometric Views
###########################################################################################################################################################
BOTTOM_ISOMETRIC_VIEW = doc.GetElement(ElementId(446525))
setup_Section_box(BOTTOM_ISOMETRIC_VIEW, 0, +rBase + 1, 0, +rBase + 1, -hPit - 1, hCone)

TOP_ISOMETRIC_VIEW = doc.GetElement(ElementId(446535))
#get max Z of bars with "TR" in schedule mark to set the section box height accordingly
FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
for bar in FEC:
    if "TR" in bar.LookupParameter("Schedule Mark").AsValueString():
        bb = bar.get_BoundingBox(None)
        if bb is not None:
            if bb.Max.Z > hCone:
                hCone = bb.Max.Z
setup_Section_box(TOP_ISOMETRIC_VIEW, 0, +rBase + 1, 0, +rBase + 1, -hPit - 1, hCone)
hCone = WTF_foundation.LookupParameter("hCone").AsDouble()

SHEAR_REINFORCEMENT_ISOMETRIC_VIEW = doc.GetElement(ElementId(2503522))
setup_Section_box(SHEAR_REINFORCEMENT_ISOMETRIC_VIEW, 0, +rBase + 1, 0, +rBase + 1, -hPit - 1, hCone)

STOOL_SETUP_ISOMETRIC_VIEW = doc.GetElement(ElementId(2514065))
# setup_Section_box(STOOL_SETUP_ISOMETRIC_VIEW, 0, +rBase + 1, 0, +rBase + 1, -hPit - 1, hCone)

ST103_SETUP_ISOMETRIC_VIEW = doc.GetElement(ElementId(5998757))


if MoveViews == True:
    # Align by model origin so crop region size differences do not affect placement.
    radial_view_ids = {
        'TOP_1':    TOP_1_RADIAL_REBAR_PLAN_VIEW.Id,
        'BOTTOM_1': BOTTOM_1_RADIAL_REBAR_PLAN_VIEW.Id,
        'BOTTOM_3': BOTTOM_3_RADIAL_REBAR_PLAN_VIEW.Id,
        'TOP_2':    TOP_2_RADIAL_REBAR_PLAN_VIEW.Id,
    }

    radial_views = {
        'TOP_1':    TOP_1_RADIAL_REBAR_PLAN_VIEW,
        'BOTTOM_1': BOTTOM_1_RADIAL_REBAR_PLAN_VIEW,
        'BOTTOM_3': BOTTOM_3_RADIAL_REBAR_PLAN_VIEW,
        'TOP_2':    TOP_2_RADIAL_REBAR_PLAN_VIEW,
    }

    def _sheet_point_for_model_point(view, viewport, model_point):
        if hasattr(viewport, "GetProjectionToSheetTransform"):
            tf = viewport.GetProjectionToSheetTransform()
            if tf is not None:
                return tf.OfPoint(model_point)

        delta_model = model_point - view.Origin
        u = delta_model.DotProduct(view.RightDirection) / float(view.Scale)
        v = delta_model.DotProduct(view.UpDirection) / float(view.Scale)

        if viewport.Rotation == ViewportRotation.Clockwise:
            u, v = v, -u
        elif viewport.Rotation == ViewportRotation.Counterclockwise:
            u, v = -v, u
        elif viewport.Rotation == ViewportRotation.Rotate180:
            u, v = -u, -v

        c = viewport.GetBoxCenter()
        return XYZ(c.X + u, c.Y + v, c.Z)

    radial_vp = {}
    for vp in FilteredElementCollector(doc).OfClass(Viewport).WhereElementIsNotElementType():
        for key, vid in radial_view_ids.items():
            if vp.ViewId == vid:
                radial_vp[key] = vp

    if all(k in radial_vp for k in radial_view_ids):
        doc.Regenerate()
        model_origin = XYZ(0, 0, 0)

        ref_vp = radial_vp['TOP_1']
        ref_rotation = ref_vp.Rotation
        ref_origin_on_sheet = _sheet_point_for_model_point(radial_views['TOP_1'], ref_vp, model_origin)

        for key in ['BOTTOM_1', 'BOTTOM_3', 'TOP_2']:
            vp = radial_vp[key]
            if vp.Rotation != ref_rotation:
                vp.Rotation = ref_rotation
                doc.Regenerate()

            current_origin_on_sheet = _sheet_point_for_model_point(radial_views[key], vp, model_origin)
            delta = XYZ(
                ref_origin_on_sheet.X - current_origin_on_sheet.X,
                ref_origin_on_sheet.Y - current_origin_on_sheet.Y,
                0
            )
            c = vp.GetBoxCenter()
            vp.SetBoxCenter(XYZ(c.X + delta.X, c.Y + delta.Y, c.Z))


#endregion ################################################################################################################################################################################################################################################################################################################################################################################################################################################################

#region REBAR DWG 2 ################################################################################################################################################################################################################################################################################################################################################################################################################################################################

#Plan Views
RADIATOR_BASE_PLAN_VIEW = doc.GetElement(ElementId(5051956))

#Section Views
REINFORCEMENT_SECTION_VIEW = doc.GetElement(ElementId(4355490))
boxMin_X = -rBase - 2
boxMax_X = rBase + 2
boxMin_Y = -hPit - 2
boxMax_Y = hPlinth + TOflange_TOtower + 2 
secDepth = 0.01
setup_Crop_box(REINFORCEMENT_SECTION_VIEW, boxMin_X, boxMin_Y, boxMax_X, boxMax_Y, secDepth)

RADIATOR_BASE_SECTION_A_A_VIEW = doc.GetElement(ElementId(5084746))
RADIATOR_BASE_SECTION_B_B_VIEW = doc.GetElement(ElementId(6008895))

#Isometric Views

SURFACE_REINFORCEMENT_ISOMETRIC_VIEW = doc.GetElement(ElementId(603339))
setup_Section_box(SURFACE_REINFORCEMENT_ISOMETRIC_VIEW, -32/304.8, +rVoidOuter + 1, -32/304.8, +rVoidOuter + 1, -hPit - 1, 0)

BOTTOM_BURSTING_REBAR_ISOMETRIC_VIEW = doc.GetElement(ElementId(603329))
setup_Section_box(BOTTOM_BURSTING_REBAR_ISOMETRIC_VIEW, -32/304.8, +rVoidOuter + 1, -32/304.8, +rVoidOuter + 1, 0, 2)

PLINTH_REINFORCEMENT_ISOMETRIC_VIEW = doc.GetElement(ElementId(966797))
setup_Section_box(PLINTH_REINFORCEMENT_ISOMETRIC_VIEW, -max(rPlinth,rVoidOuter) - 1,32/304.8 , -32/304.8, +max(rPlinth,rVoidOuter) + 1, -hPit - 1, hPlinth + TOflange_TOtower + 1)

TOP_BURSTING_REBAR_1_ISOMETRIC_VIEW = doc.GetElement(ElementId(3095802))
setup_Section_box(TOP_BURSTING_REBAR_1_ISOMETRIC_VIEW, -32/304.8, +rVoidOuter + 1, -32/304.8, +rVoidOuter + 1, hCone+3, hPlinth)

TOP_BURSTING_REBAR_2_ISOMETRIC_VIEW = doc.GetElement(ElementId(3096005))
setup_Section_box(TOP_BURSTING_REBAR_2_ISOMETRIC_VIEW, -32/304.8, +rVoidOuter + 1, -32/304.8, +rVoidOuter + 1, 2, hCone+1.5)
#endregion################################################################################################################################################################################################################################################################################################################################################################################################################################################################

t.Commit()