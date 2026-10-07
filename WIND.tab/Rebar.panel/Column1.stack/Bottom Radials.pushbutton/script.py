import math
import clr
import traceback
from Autodesk.Revit.DB import (
    Curve, Line, XYZ,
    Transaction, FilteredElementCollector,
    RadialArray, ArrayAnchorMember,
    BuiltInCategory, BuiltInParameter,
    ElementTransformUtils,
    FailureSeverity, FailureProcessingResult, IFailuresPreprocessor,
    Structure,
)
from Autodesk.Revit.DB.Structure import RebarShape, RebarHookOrientation, RebarStyle
from System.Collections.Generic import List
import Functions as func

clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
import atexit

# Conversion factor: millimetres to Revit internal feet
MM = 1.0 / 304.8

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


class SuppressWarnings(IFailuresPreprocessor):

    def PreprocessFailures(self, failuresAccessor):
        try:
            failures = failuresAccessor.GetFailureMessages()
            for failure in failures:
                if failure.GetSeverity() == FailureSeverity.Warning:
                    failuresAccessor.DeleteWarning(failure)
        except Exception:
            debug_print(traceback.format_exc())
        return FailureProcessingResult.Continue

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
  A -> bar mark        e.g. BR100
    B -> start radius    e.g. 500
    C -> end radius      e.g. 1500
    D -> size            e.g. Y32
    E -> rotation switch e.g. 1 (IF = 1, bar goes between anchors; IF = 0, bar goes on anchors)
    F -> (not used)
    G -> no bars factor  e.g. 2 (number of bars will be this factor multiplied by the number of anchor bolts)
    H -> (not used)
    I -> (not used)
    J -> level           e.g. 1 or 2 (1 = lower level, 2 = upper level)
    K -> Start Hook Length e.g. 150 (if > 100mm, a hook will be created)
    L -> (not used)
    M -> Make 20mm rebar IF = 1 (IF = 1, bars will be made with 20mm shape regardless of their size, to achieve the required hook shape)
_____________________________________________________________________

"""

#### Input from excel sheet ####################################################################
FPath = func.getFilePath()

excel = Excel.ApplicationClass()
excel.Visible = False
workbook = excel.Workbooks.Open(FPath)
try:
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

    ####  Get number of Anchor Bolts#################################################################
    AnchorBolts = func.getAnchorBolts(xl)

    BRMaxDia = 0
    for i in range(5, 200):
        if "BR" in str(xl.Cells(i, 1).Value2):
            try:
                dia = int(str(xl.Cells(i, 4).Value2)[1:])
                if BRMaxDia < dia:
                    BRMaxDia = dia
            except Exception:
                BRMaxDia = 32
    slabFaceConDia = 20 * MM
    for i in range(5, 200):
        if "SF" in str(xl.Cells(i, 1).Value2):
            try:
                slabFaceConDia = int(str(xl.Cells(i, 4).Value2)[1:]) * MM
            except Exception:
                slabFaceConDia = 20 * MM
            break

    #### Collect Revit data once before the transaction loop ####################################

    rebar_shapes = FilteredElementCollector(doc).OfClass(RebarShape).WhereElementIsElementType().ToElements()
    sc_41 = sc_20 = sc_37 = None
    for r_shape in rebar_shapes:
        name = r_shape.LookupParameter("Type Name").AsString()
        if name == '41':
            sc_41 = r_shape
        elif name == '20':
            sc_20 = r_shape
        elif name == '37':
            sc_37 = r_shape

    all_rebar_types = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_Rebar)
        .WhereElementIsElementType()
        .ToElements()
    )

    wtf_elements = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
    wtf_host = None
    for element in wtf_elements:
        if element.Name == "1PA_WTF_SteelTower":
            wtf_host = element
            break

    r_base    = wtf_host.LookupParameter("rBase").AsDouble()
    rPitOuter = wtf_host.LookupParameter("rVoidOuter").AsDouble()
    rPitInner = wtf_host.LookupParameter("rVoidInner").AsDouble()
    hPit      = wtf_host.LookupParameter("hBottomVoid").AsDouble()
    # Assumes tower family is placed at the project origin (0,0,0)
    locPoint  = XYZ(0, 0, 0)

    # constant cover and grid bar diameter
    bot_cover = 50 * MM
    grid1dia  = 32 * MM

    #### Transaction ################################################################################
    t = Transaction(doc, 'Reinforce')
    fail_opts = t.GetFailureHandlingOptions()
    fail_opts.SetFailuresPreprocessor(SuppressWarnings())
    t.SetFailureHandlingOptions(fail_opts)
    t.Start()

    for i in range(5, 200):
        if "BR" not in str(xl.Cells(i, 1).Value2):
            continue

        bar_mark        = str(xl.Cells(i, 1).Value2)
        no_bars_factor  = float(xl.Cells(i, 7).Value2)
        size            = "Y" + str(xl.Cells(i, 4).Value2)[1:]
        rot_switch      = float(xl.Cells(i, 5).Value2)
        StartRad        = int(xl.Cells(i, 2).Value2) * MM
        EndRad          = int(xl.Cells(i, 3).Value2) * MM
        StartHookLength = float(xl.Cells(i, 11).Value2) * MM
        Level           = int(xl.Cells(i, 10).Value2)
        Make_20         = int(xl.Cells(i, 13).Value2)

        debug_print("#" * 50)
        debug_print("  MARK: {0}  |  FACTOR: {1}  |  SIZE: {2}  |  ROT: {3}"
              "  |  StartR: {4:.0f}mm  |  EndR: {5:.0f}mm"
              "  |  Hook: {6:.0f}mm  |  Level: {7}".format(
                  bar_mark, no_bars_factor, size, rot_switch,
                  StartRad / MM, EndRad / MM, StartHookLength / MM, Level))
        debug_print("  Make_20: {}".format(Make_20))

        no_bars = int(no_bars_factor * AnchorBolts)

        ##### Rebar type ####################################################################

        bar_type = None
        for rebar_type in all_rebar_types:
            rebar_name = rebar_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            if rebar_name == size:
                bar_type = rebar_type
                break

        barDia = bar_type.LookupParameter("Bar Diameter").AsDouble()

        if StartRad < rPitOuter and StartRad > rPitInner:
            debug_print("check bar {} - Starting Radius is on pit slope".format(bar_mark))

        if EndRad > (r_base - bot_cover - slabFaceConDia):
            EndRad = r_base - bot_cover - slabFaceConDia

        #### Rebar Shape Properties ####################################################################

        A = StartHookLength

        debug_print("  A: {:.1f}mm".format(A / MM))

        # BRMaxDia is raw mm; barDia is Revit feet - both terms convert to feet
        level2Offset = (64.0 + BRMaxDia / 2.0) * MM + barDia / 2.0

        z_bot = bot_cover + barDia / 2
        if Level == 1:
            rebar_p1 = locPoint + XYZ(StartRad, 0, bot_cover + A)
            rebar_p2 = locPoint + XYZ(StartRad, 0, z_bot)
            rebar_p3 = locPoint + XYZ(EndRad,   0, z_bot)
        elif Level == 2:
            rebar_p1 = locPoint + XYZ(StartRad, 0, bot_cover + A + level2Offset)
            rebar_p2 = locPoint + XYZ(StartRad, 0, z_bot + level2Offset)
            rebar_p3 = locPoint + XYZ(EndRad,   0, z_bot + level2Offset)
        else:
            debug_print("Level not valid")
            continue

        theta     = math.atan(hPit / (rPitOuter - rPitInner))
        xdelta    = (barDia / 2 + bot_cover) * math.cos(math.pi / 2 - theta)
        gridDelta = (2 * grid1dia) / math.tan(theta)

        rebar_p4 = locPoint + XYZ(rPitOuter - xdelta,                   0, z_bot)
        rebar_p5 = locPoint + XYZ(rPitInner - xdelta + gridDelta,        0, z_bot - hPit + grid1dia * 2)
        rebar_p6 = locPoint + XYZ(StartRad,                              0, z_bot - hPit + grid1dia * 2)

        curve1 = Line.CreateBound(rebar_p1, rebar_p2)
        curve2 = Line.CreateBound(rebar_p2, rebar_p3)
        curve3 = Line.CreateBound(rebar_p6, rebar_p5)
        curve4 = Line.CreateBound(rebar_p5, rebar_p4)
        curve5 = Line.CreateBound(rebar_p4, rebar_p3)

        curve_list20 = List[Curve]([curve2])
        curve_list37 = List[Curve]([curve1, curve2])
        curve_list41 = List[Curve]([curve3, curve4, curve5])

        #### Build rebar ####################################################################
        hook_threshold = 100 * MM
        if Make_20 == 1:
            rebar = Structure.Rebar.CreateFromCurvesAndShape(
                doc, sc_20, bar_type, None, None,
                wtf_host, XYZ.BasisY, curve_list20,
                RebarHookOrientation.Left, RebarHookOrientation.Left)
        elif StartHookLength < hook_threshold and Level == 2:
            rebar = Structure.Rebar.CreateFromCurvesAndShape(
                            doc, sc_20, bar_type, None, None,
                            wtf_host, XYZ.BasisY, curve_list20,
                            RebarHookOrientation.Left, RebarHookOrientation.Left)
        elif StartRad > rPitOuter + 0.1:
            if StartHookLength < hook_threshold:
                rebar = Structure.Rebar.CreateFromCurvesAndShape(
                    doc, sc_20, bar_type, None, None,
                    wtf_host, XYZ.BasisY, curve_list20,
                    RebarHookOrientation.Left, RebarHookOrientation.Left)
            else:
                rebar = Structure.Rebar.CreateFromCurvesAndShape(
                    doc, RebarStyle.Standard, bar_type, None, None,
                    wtf_host, XYZ.BasisY, curve_list37,
                    RebarHookOrientation.Left, RebarHookOrientation.Left, 1, 0)
        else:
            if StartHookLength > hook_threshold:
                rebar = Structure.Rebar.CreateFromCurvesAndShape(
                    doc, sc_37, bar_type, None, None,
                    wtf_host, XYZ.BasisY, curve_list37,
                    RebarHookOrientation.Left, RebarHookOrientation.Left)
            else:
                rebar = Structure.Rebar.CreateFromCurvesAndShape(
                    doc, sc_41, bar_type, None, None,
                    wtf_host, XYZ.BasisY, curve_list41,
                    RebarHookOrientation.Left, RebarHookOrientation.Left)

        rebar.LookupParameter(PARAM_MARK).Set("BOTTOM RADIAL")
        rebar.LookupParameter(PARAM_SCHEDULE_MARK).Set(bar_mark)
        rebar.LookupParameter("Rebar r Custom").Set(0)
        rebar.LookupParameter("Rebar Spacing").Set((360.0 / no_bars) / 304.8)
        rebar.LookupParameter("Rebar Quantity").Set(no_bars)

        debug_print("  Construction bar created")
        #### Build Section Bar ####################################################################
        #copy rebar to left and right of the section in the plane Y=0
        section_bar_right_id = ElementTransformUtils.CopyElement(doc, rebar.Id, XYZ(0, 0, 0))[0]
        section_bar_right = doc.GetElement(section_bar_right_id)
        section_bar_right.LookupParameter(PARAM_MARK).Set("SECTION BOTTOM RADIAL")
        section_bar_right.LookupParameter(PARAM_SCHEDULE_MARK).Set(bar_mark)
        section_bar_right.LookupParameter("Rebar r Custom").Set(0)
        section_bar_right.LookupParameter("Rebar Spacing").Set((360.0 / no_bars) / 304.8)
        section_bar_right.LookupParameter("Rebar Quantity").Set(no_bars)

        # rotate the second copy 180 degrees around the vertical axis to mirror it on the other side of the section
        section_bar_left_id = ElementTransformUtils.CopyElement(doc, rebar.Id, XYZ(0, 0, 0))[0]
        section_bar_left = doc.GetElement(section_bar_left_id)
        section_bar_left.Location.Rotate(Line.CreateBound(locPoint, locPoint + XYZ.BasisZ), math.pi)
        section_bar_left.LookupParameter(PARAM_MARK).Set("SECTION BOTTOM RADIAL")
        section_bar_left.LookupParameter(PARAM_SCHEDULE_MARK).Set(bar_mark)
        section_bar_left.LookupParameter("Rebar r Custom").Set(0)
        section_bar_left.LookupParameter("Rebar Spacing").Set((360.0 / no_bars) / 304.8)
        section_bar_left.LookupParameter("Rebar Quantity").Set(no_bars)
        #### Radial array ####################################################################
        RotAngle = 2 * math.pi
        axis_line = Line.CreateBound(locPoint, locPoint + XYZ.BasisZ)
        elem = RadialArray.ArrayElementWithoutAssociation(
            doc, view, rebar.Id, no_bars, axis_line, RotAngle, ArrayAnchorMember.Last)

        rot_step = rot_switch * RotAngle / (no_bars * 2)
        for elm in elem:
            element = doc.GetElement(elm)
            element.Location.Rotate(Line.CreateBound(locPoint, locPoint + XYZ.BasisZ), rot_step)
            element.LookupParameter(PARAM_MARK).Set("BOTTOM RADIAL")
            element.LookupParameter(PARAM_SCHEDULE_MARK).Set(bar_mark)
            element.LookupParameter("Rebar r Custom").Set(0)
            element.LookupParameter("Rebar Spacing").Set((360.0 / no_bars)/304.8)
            element.LookupParameter("Rebar Quantity").Set(no_bars)
        debug_print("  Radial array created")
        debug_print("#" * 50)

    t.Commit()

finally:
    workbook.Close(False)
    excel.Quit()









