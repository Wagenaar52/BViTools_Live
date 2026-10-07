from Autodesk.Revit.DB.Structure import * 
from Autodesk.Revit.DB import CurveByPoints, ReferencePointArray, ReferencePoint, CurveArray, PolyLine, Plane, SketchPlane, ElementTransformUtils
from System.Collections.Generic import List
from Autodesk.Revit.DB import Curve, Line, XYZ
from Autodesk.Revit.DB.Structure import RebarShape
import math, clr
from Autodesk.Revit.DB import Transaction, Structure, FilteredElementCollector, RadialArray, ArrayAnchorMember
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, Line, XYZ, FailureSeverity, FailureProcessingResult,IFailuresPreprocessor
from pyrevit import forms
import Functions as func

clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
import atexit

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
view = uidoc.ActiveView

DEBUG = True
PARAM_MARK = "Mark"
PARAM_SCHEDULE_MARK = "Schedule Mark"

def debug_print(*args):
    if DEBUG:
        import sys
        sys.stdout.write(" ".join([str(a) for a in args]) + "\n")

__doc__ = """
Version = 2.0
Date = 04.05.2026
_____________________________________________________________________
How-to:
-> be in a 3D view
-> click the button
-> select the excel file with the rebar data
_____________________________________________________________________
Excel file format (columns):
  A -> bar mark        e.g. BV100
  B -> radius (mm)     e.g. 12000
  D -> bar type        e.g. Y20
  E -> rot. switch     e.g. 1
  G -> no. bars factor e.g. 4

- Bar mark: must be BV100
- Outer radius: the distance from the center of the tower to the center of the rebar (in mm)
- Bar type: the size of rebar (e.g., Y20)
- Rotation switch: a value to control rebar rotation. 1 means the first bar is rotated by 360/no_bars_factor degrees; 2 means 2*360/no_bars_factor degrees; etc.
- No bars factor: a factor to determine the number of bars
"""

#### Constants ####################################################################

MM_TO_FT        = 1.0 / 304.8
TOP_COVER_MM    = 40.0
BOT_COVER_MM    = 50.0
FULL_CIRCLE_RAD = 2 * math.pi


class SuppressWarnings(IFailuresPreprocessor):

    def PreprocessFailures(self, failuresAccessor):
        try:
            failures = failuresAccessor.GetFailureMessages()
            for failure in failures:
                severity = failure.GetSeverity()
                if severity == FailureSeverity.Warning:
                    failuresAccessor.DeleteWarning(failure)
        except:
            import traceback
            debug_print(traceback.format_exc())

        return FailureProcessingResult.Continue


def set_rebar_params(rebar_elem, bar_mark, no_bars):
    rebar_elem.LookupParameter(PARAM_MARK).Set("BASE VERTICALS")
    rebar_elem.LookupParameter("Rebar r Custom").Set(0)
    rebar_elem.LookupParameter(PARAM_SCHEDULE_MARK).Set(bar_mark)
    rebar_elem.LookupParameter("Rebar Quantity").Set(no_bars)


def create_rebar(doc, bar_type, wtf_element, curve_list):
    return Structure.Rebar.CreateFromCurves(
        doc,
        RebarStyle.Standard,
        bar_type,
        None,
        None,
        wtf_element,
        XYZ.BasisY,
        curve_list,
        RebarHookOrientation.Left,
        RebarHookOrientation.Left, 1, 0)


#### Input from excel sheet ####################################################################

FPath = func.getFilePath()

excel = Excel.ApplicationClass()
excel.Visible = False
workbook = excel.Workbooks.Open(FPath)
xl = workbook.Worksheets['A']

def _cleanup_excel():
    try:
        workbook.Close(False)
    except:
        pass
    try:
        excel.Quit()
    except:
        pass

atexit.register(_cleanup_excel)

AnchorBolts = func.getAnchorBolts(xl)

#### Rebar Shape ####################################################################

rebar_shape = FilteredElementCollector(doc).OfClass(RebarShape).WhereElementIsElementType().ToElements()
target_shapes = {'97': None, '20': None, '37': None}
for r_shape in rebar_shape:
    name = r_shape.LookupParameter("Type Name").AsString()
    if name in target_shapes:
        target_shapes[name] = r_shape
sc_97 = target_shapes['97']

#### Rebar Types (collected once) ####################################################################

all_rebar_types = FilteredElementCollector(doc) \
    .OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsElementType() \
    .ToElements()

#### Element Host (collected once) ####################################################################

wtf_collection = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
wtf_element = None
for element in wtf_collection:
    if element.Name == "1PA_WTF_SteelTower":
        wtf_element = element
        break

r_base    = wtf_element.LookupParameter("rBase").AsDouble()
h_base    = wtf_element.LookupParameter("hBase").AsDouble()
r_plinth  = wtf_element.LookupParameter("rPlinth").AsDouble()
h_cone    = wtf_element.LookupParameter("hCone").AsDouble()
h_plinth  = wtf_element.LookupParameter("hPlinth").AsDouble()
slabSlope = (h_cone - h_base) / (r_base - r_plinth)

top_cover = TOP_COVER_MM * MM_TO_FT
bot_cover = BOT_COVER_MM * MM_TO_FT
locPoint  = wtf_element.Location.Point

SlabFace_dia = 0

t = Transaction(doc, 'Reinforce')
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SuppressWarnings())
t.SetFailureHandlingOptions(failHandler)
t.Start()

for i in range(5, 200):
    if "SF100" in str(xl.Cells(i, 1).Value2):
        SlabFace_dia = int(str(xl.Cells(i, 4).Value2)[1:]) * MM_TO_FT


for i in range(5, 200):
    if "BV100" in str(xl.Cells(i, 1).Value2):
        bar_mark       = str(xl.Cells(i, 1).Value2)
        no_bars_factor = float(xl.Cells(i, 7).Value2)
        size           = "Y" + str(xl.Cells(i, 4).Value2)[1:]
        RotSwitch      = float(xl.Cells(i, 5).Value2)

        debug_print("  ----   " + "BAR MARK" + "  ----   " + "NO BARS FACTOR" + "  ---  " + "SIZE")
        debug_print("  ----   " + bar_mark + "  ----   " + str(no_bars_factor) + "  ---  " + size)
        debug_print("*" * 30)

        no_bars = no_bars_factor * AnchorBolts

        ##### Rebar type ####################################################################

        bar_type = None
        for rebar_type in all_rebar_types:
            rebar_name = rebar_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            if rebar_name == size:
                bar_type = rebar_type
                break

        barDia = bar_type.LookupParameter("Bar Diameter").AsDouble()

        #### Rebar Shape Properties ####################################################################

        A = 800 * MM_TO_FT
        y_Rangetop = h_base + slabSlope * bot_cover - top_cover - barDia * 1.5
        y_Rangebot = bot_cover + barDia
        B = y_Rangetop - y_Rangebot
        C = A * slabSlope
        D = 13.5 * barDia
        J = B + C

        debug_print("A: " + str(A / MM_TO_FT))
        debug_print("B: " + str(B / MM_TO_FT))
        debug_print("C: " + str(C / MM_TO_FT))
        debug_print("D: " + str(D / MM_TO_FT))
        debug_print("J: " + str(J / MM_TO_FT))

        #### Build ####################################################################

        rebar_p1 = locPoint + XYZ(r_base - bot_cover - A - SlabFace_dia,          0, bot_cover + J)
        rebar_p2 = locPoint + XYZ(r_base - bot_cover - SlabFace_dia - (barDia/2), 0, bot_cover + B + barDia)
        rebar_p3 = locPoint + XYZ(r_base - bot_cover - SlabFace_dia - (barDia/2), 0, bot_cover + barDia/2)
        rebar_p4 = locPoint + XYZ(r_base - bot_cover - SlabFace_dia - A,          0, bot_cover + barDia/2)
        rebar_p5 = locPoint + XYZ(r_base - bot_cover - A - SlabFace_dia + D,      0, bot_cover + barDia/2)
        rebar_p6 = locPoint + XYZ(r_base - bot_cover - A - SlabFace_dia,          0, bot_cover + D + barDia/2)

        curve1 = Line.CreateBound(rebar_p5, rebar_p4)
        curve2 = Line.CreateBound(rebar_p4, rebar_p1)
        curve3 = Line.CreateBound(rebar_p1, rebar_p2)
        curve4 = Line.CreateBound(rebar_p2, rebar_p3)
        curve5 = Line.CreateBound(rebar_p3, rebar_p4)
        curve6 = Line.CreateBound(rebar_p4, rebar_p6)

        curve_list = List[Curve]([curve1, curve2, curve3, curve4, curve5, curve6])

        geomPlane = Plane.CreateByThreePoints(rebar_p1, rebar_p2, rebar_p3)
        sketch = SketchPlane.Create(doc, geomPlane)

        rebar = create_rebar(doc, bar_type, wtf_element, curve_list)

        #### Build Section Bar ####################################################################
        #copy rebar to left and right of the section in the plane Y=0
        section_bar_right_id = ElementTransformUtils.CopyElement(doc, rebar.Id, XYZ(0, 0, 0))[0]
        section_bar_right = doc.GetElement(section_bar_right_id)
        section_bar_right.LookupParameter(PARAM_MARK).Set("SECTION BASE VERTICAL")
        section_bar_right.LookupParameter(PARAM_SCHEDULE_MARK).Set(bar_mark)
        section_bar_right.LookupParameter("Rebar r Custom").Set(0)
        section_bar_right.LookupParameter("Rebar Spacing").Set((360.0 / no_bars) / 304.8)
        section_bar_right.LookupParameter("Rebar Quantity").Set(no_bars)

        # rotate the second copy 180 degrees around the vertical axis to mirror it on the other side of the section
        section_bar_left_id = ElementTransformUtils.CopyElement(doc, rebar.Id, XYZ(0, 0, 0))[0]
        section_bar_left = doc.GetElement(section_bar_left_id)
        section_bar_left.Location.Rotate(Line.CreateBound(locPoint, locPoint + XYZ.BasisZ), math.pi)
        section_bar_left.LookupParameter(PARAM_MARK).Set("SECTION BASE VERTICAL")
        section_bar_left.LookupParameter(PARAM_SCHEDULE_MARK).Set(bar_mark)
        section_bar_left.LookupParameter("Rebar r Custom").Set(0)
        section_bar_left.LookupParameter("Rebar Spacing").Set((360.0 / no_bars) / 304.8)
        section_bar_left.LookupParameter("Rebar Quantity").Set(no_bars)

        #### Radial Array ####################################################################

        axis = Line.CreateBound(locPoint, locPoint + XYZ.BasisZ)

        if no_bars <= 200:
            elem = RadialArray.ArrayElementWithoutAssociation(doc, view, rebar.Id, no_bars, axis, FULL_CIRCLE_RAD, ArrayAnchorMember.Last)
            for elm in elem:
                doc.GetElement(elm).Location.Rotate(Line.CreateBound(locPoint, locPoint + XYZ.BasisZ), RotSwitch * FULL_CIRCLE_RAD / (no_bars * 4))
                set_rebar_params(doc.GetElement(elm), bar_mark, no_bars)
        else:
            elem  = RadialArray.ArrayElementWithoutAssociation(doc, view, rebar.Id, no_bars/4, axis, FULL_CIRCLE_RAD, ArrayAnchorMember.Last)
            elem2 = RadialArray.ArrayElementWithoutAssociation(doc, view, rebar.Id, no_bars/4, axis, FULL_CIRCLE_RAD, ArrayAnchorMember.Last)
            elem3 = RadialArray.ArrayElementWithoutAssociation(doc, view, rebar.Id, no_bars/4, axis, FULL_CIRCLE_RAD, ArrayAnchorMember.Last)
            elem4 = RadialArray.ArrayElementWithoutAssociation(doc, view, rebar.Id, no_bars/4, axis, FULL_CIRCLE_RAD, ArrayAnchorMember.Last)

            ###########
            rebar2 = create_rebar(doc, bar_type, wtf_element, curve_list)
            set_rebar_params(rebar2, bar_mark, no_bars)
            doc.GetElement(rebar2.Id).Location.Rotate(Line.CreateBound(locPoint, locPoint + XYZ.BasisZ), FULL_CIRCLE_RAD / no_bars)
            ###########
            rebar2 = create_rebar(doc, bar_type, wtf_element, curve_list)
            set_rebar_params(rebar2, bar_mark, no_bars)
            doc.GetElement(rebar2.Id).Location.Rotate(Line.CreateBound(locPoint, locPoint + XYZ.BasisZ), FULL_CIRCLE_RAD / no_bars)
            ###########
            rebar3 = create_rebar(doc, bar_type, wtf_element, curve_list)
            set_rebar_params(rebar3, bar_mark, no_bars)
            doc.GetElement(rebar3.Id).Location.Rotate(Line.CreateBound(locPoint, locPoint + XYZ.BasisZ), 2 * FULL_CIRCLE_RAD / no_bars)
            ###########
            rebar4 = create_rebar(doc, bar_type, wtf_element, curve_list)
            set_rebar_params(rebar4, bar_mark, no_bars)
            doc.GetElement(rebar4.Id).Location.Rotate(Line.CreateBound(locPoint, locPoint + XYZ.BasisZ), 3 * FULL_CIRCLE_RAD / no_bars)

            for elm in elem2:
                doc.GetElement(elm).Location.Rotate(Line.CreateBound(locPoint, locPoint + XYZ.BasisZ), FULL_CIRCLE_RAD / no_bars)
                set_rebar_params(doc.GetElement(elm), bar_mark, no_bars)
            for elm in elem3:
                doc.GetElement(elm).Location.Rotate(Line.CreateBound(locPoint, locPoint + XYZ.BasisZ), 2 * FULL_CIRCLE_RAD / no_bars)
                set_rebar_params(doc.GetElement(elm), bar_mark, no_bars)
            for elm in elem4:
                doc.GetElement(elm).Location.Rotate(Line.CreateBound(locPoint, locPoint + XYZ.BasisZ), 3 * FULL_CIRCLE_RAD / no_bars)
                set_rebar_params(doc.GetElement(elm), bar_mark, no_bars)
            for elm in elem:
                set_rebar_params(doc.GetElement(elm), bar_mark, no_bars)

            doc.Delete(rebar.Id)
            doc.Delete(rebar2.Id)
            rebar = create_rebar(doc, bar_type, wtf_element, curve_list)
            set_rebar_params(rebar, bar_mark, no_bars)

t.Commit()

workbook.Close(False)
excel.Quit()









