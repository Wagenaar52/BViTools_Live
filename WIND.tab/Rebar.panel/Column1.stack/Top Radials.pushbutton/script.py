from Autodesk.Revit.DB.Structure import *
from System.Collections.Generic import List
from Autodesk.Revit.DB import (
    Curve, Line, XYZ,
    Transaction, Structure, FilteredElementCollector,
    RadialArray, ArrayAnchorMember,
    ElementTransformUtils,
    BuiltInCategory, BuiltInParameter,
    FailureSeverity, FailureProcessingResult, IFailuresPreprocessor,
)
from Autodesk.Revit.DB.Structure import RebarShape
import math
import clr
from pyrevit import forms
import Functions as func

clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
import atexit

MM_TO_FT = 1.0 / 304.8

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
view = uidoc.ActiveView

DEBUG = False
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
    A -> bar mark        e.g. TR100
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
_____________________________________________________________________

"""

class SuppressWarnings(IFailuresPreprocessor):
    
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
            debug_print(traceback.format_exc())
        
        return FailureProcessingResult.Continue


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

try:

    ##### Element Host ####################################################################

    WTF = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
    for element in WTF:
        if element.Name == "1PA_WTF_SteelTower":
            WTF = element
            break

    wtf_type = doc.GetElement(WTF.GetTypeId())
    r_base = WTF.LookupParameter("rBase").AsDouble()
    h_base = WTF.LookupParameter("hBase").AsDouble()
    r_plinth = WTF.LookupParameter("rPlinth").AsDouble()
    h_cone = WTF.LookupParameter("hCone").AsDouble()
    h_plinth = WTF.LookupParameter("hPlinth").AsDouble()
    slabSlope = (h_cone - h_base) / (r_base - r_plinth)
    rPitOuter = WTF.LookupParameter("rVoidOuter").AsDouble()
    rPitInner = WTF.LookupParameter("rVoidInner").AsDouble()
    hPit = WTF.LookupParameter("hBottomVoid").AsDouble()
    locPoint = WTF.Location.Point
    TopRadBendLocRad = float(xl.Cells(3, 7).Value2) * MM_TO_FT

    #####Rebar Shape ####################################################################

    sc_62 = sc_99g = sc_20 = sc_99i = None
    shape_names = {'62': None, '99g': None, '20': None, '99i': None}
    for r_shape in FilteredElementCollector(doc).OfClass(RebarShape).WhereElementIsElementType().ToElements():
        name = r_shape.LookupParameter("Type Name").AsString()
        if name in shape_names:
            shape_names[name] = r_shape
    sc_62   = shape_names['62']
    sc_99g  = shape_names['99g']
    sc_20   = shape_names['20']
    sc_99i  = shape_names['99i']

    #### Get number of Anchor Bolts ####################################################################
    AnchorBolts = func.getAnchorBolts(xl)

    TRMaxDia = 0
    for i in range(5, 200):
        if "TR" in str(xl.Cells(i, 1).Value2):
            dia = int(str(xl.Cells(i, 4).Value2)[1:])
            if TRMaxDia < dia:
                TRMaxDia = dia

    for i in range(5, 200):
        if "SF" in str(xl.Cells(i, 1).Value2):
            slabFaceConDia = int(str(xl.Cells(i, 4).Value2)[1:]) * MM_TO_FT

    #### Rebar Types (collected once) ####################################################################
    all_rebar_types = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_Rebar)
        .WhereElementIsElementType()
        .ToElements()
    )

    #### Transaction ####################################################################
    t = Transaction(doc, 'Reinforce')
    fail_options = t.GetFailureHandlingOptions()
    fail_options.SetFailuresPreprocessor(SuppressWarnings())
    t.SetFailureHandlingOptions(fail_options)
    t.Start()

    for i in range(5, 200):
        if "TR" in str(xl.Cells(i, 1).Value2):
            bar_mark         = str(xl.Cells(i, 1).Value2)
            no_bars_factor   = float(xl.Cells(i, 7).Value2)
            size             = "Y" + str(xl.Cells(i, 4).Value2)[1:]
            RotSwitch        = float(xl.Cells(i, 5).Value2)
            StartRad         = int(xl.Cells(i, 2).Value2) * MM_TO_FT
            EndRad           = int(xl.Cells(i, 3).Value2) * MM_TO_FT
            StartHookLength  = float(xl.Cells(i, 11).Value2) * MM_TO_FT
            Level            = int(xl.Cells(i, 10).Value2)

            debug_print("#" * 50)
            debug_print("  BAR MARK | NO BARS FACTOR | SIZE | ROT SWITCH | START RAD | END RAD | HOOK | LEVEL")
            debug_print("  {0} | {1} | {2} | {3} | {4:.1f} | {5:.1f} | {6:.1f} | {7}".format(
                bar_mark, no_bars_factor, size, RotSwitch,
                StartRad / MM_TO_FT, EndRad / MM_TO_FT,
                StartHookLength / MM_TO_FT, Level))
            debug_print("*" * 10)

            no_bars = no_bars_factor * AnchorBolts

            ##### Rebar type ####################################################################
            bar_type = None
            for rebar_type in all_rebar_types:
                rebar_name = rebar_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                if rebar_name == size:
                    bar_type = rebar_type
                    break

            barDia = bar_type.LookupParameter("Bar Diameter").AsDouble()

            # cover
            top_cover = (40 + 32) * MM_TO_FT
            bot_cover = 50 * MM_TO_FT

            if EndRad > (r_base - bot_cover - slabFaceConDia):
                EndRad = r_base - bot_cover - slabFaceConDia
                debug_print("End Radius adjusted to: {0:.1f}".format(EndRad / MM_TO_FT))
                debug_print("Check End Radius")

            #### Rebar Shape Properties ####################################################################

            level2Offset = (32 + TRMaxDia / 2) * MM_TO_FT + barDia / 2

            theta  = math.atan(slabSlope)
            ydelta = top_cover / math.cos(theta)

            if Level == 1:
                rebar_p1 = locPoint + XYZ(EndRad,            0, h_base + slabSlope * ((r_base - r_plinth) - (EndRad   - r_plinth)) - barDia / 2 - ydelta)
                rebar_p2 = locPoint + XYZ(StartRad,          0, h_base + slabSlope * ((r_base - r_plinth) - (StartRad - r_plinth)) - barDia / 2 - ydelta)
                rebar_p3 = locPoint + XYZ(TopRadBendLocRad,  0, h_cone - barDia / 2 - ydelta + slabSlope * (r_plinth - TopRadBendLocRad))
                rebar_p4 = locPoint + XYZ(StartRad,          0, h_cone - barDia / 2 - ydelta + slabSlope * (r_plinth - TopRadBendLocRad))
                rebar_p5 = locPoint + XYZ(StartRad,          0, h_cone - barDia / 2 - ydelta + slabSlope * (r_plinth - TopRadBendLocRad) - StartHookLength)
            elif Level == 2:
                rebar_p1 = locPoint + XYZ(EndRad,            0, h_base + slabSlope * ((r_base - r_plinth) - (EndRad   - r_plinth)) - barDia / 2 - ydelta - level2Offset)
                rebar_p2 = locPoint + XYZ(StartRad,          0, h_base + slabSlope * ((r_base - r_plinth) - (StartRad - r_plinth)) - barDia / 2 - ydelta - level2Offset)
                rebar_p3 = locPoint + XYZ(TopRadBendLocRad,  0, h_cone - barDia / 2 - ydelta - level2Offset + slabSlope * (r_plinth - TopRadBendLocRad))
                rebar_p4 = locPoint + XYZ(StartRad,          0, h_cone - barDia / 2 - ydelta - level2Offset + slabSlope * (r_plinth - TopRadBendLocRad))
                rebar_p5 = locPoint + XYZ(StartRad,          0, h_cone - barDia / 2 - ydelta + slabSlope * (r_plinth - TopRadBendLocRad) - StartHookLength - level2Offset)
            else:
                debug_print("Level not found: check excel sheet")
                continue

            # place curves
            curve1 = Line.CreateBound(rebar_p1, rebar_p2)
            curve2 = Line.CreateBound(rebar_p1, rebar_p3)
            curve3 = Line.CreateBound(rebar_p3, rebar_p4)

            has_hook = StartHookLength > 100 * MM_TO_FT
            beyond_bend = StartRad > TopRadBendLocRad

            if has_hook:
                curve4 = Line.CreateBound(rebar_p4, rebar_p5)
                curve_list99g = List[Curve]([curve2, curve3, curve4])


            if has_hook and beyond_bend:
                curve5 = Line.CreateBound(rebar_p2, rebar_p2 - XYZ(0, 0, StartHookLength))
                curve_list99i = List[Curve]([curve1, curve5])

            curve_list20  = List[Curve]([curve1])
            curve_list62  = List[Curve]([curve2, curve3])

            #### Build ####################################################################
            if StartRad < TopRadBendLocRad:
                if not has_hook:
                    rebar = Structure.Rebar.CreateFromCurvesAndShape(
                        doc, sc_62, bar_type, None, None,
                        WTF, XYZ.BasisY, curve_list62,
                        RebarHookOrientation.Left, RebarHookOrientation.Left)
                else:
                    rebar = Structure.Rebar.CreateFromCurves(
                        doc, RebarStyle.Standard, bar_type, None, None,
                        WTF, XYZ.BasisY, curve_list99g,
                        RebarHookOrientation.Left, RebarHookOrientation.Left, 1,0)
            else:
                if not has_hook:
                    rebar = Structure.Rebar.CreateFromCurvesAndShape(
                        doc, sc_20, bar_type, None, None,
                        WTF, XYZ.BasisY, curve_list20,
                        RebarHookOrientation.Left, RebarHookOrientation.Left)
                else:
                    rebar = Structure.Rebar.CreateFromCurvesAndShape(
                        doc, sc_99i, bar_type, None, None,
                        WTF, XYZ.BasisY, curve_list99i,
                        RebarHookOrientation.Left, RebarHookOrientation.Left)

            #### Build Section Bar ####################################################################
            #copy rebar to left and right of the section in the plane Y=0
            section_bar_right_id = ElementTransformUtils.CopyElement(doc, rebar.Id, XYZ(0, 0, 0))[0]
            section_bar_right = doc.GetElement(section_bar_right_id)
            section_bar_right.LookupParameter(PARAM_MARK).Set("SECTION TOP RADIAL")
            section_bar_right.LookupParameter(PARAM_SCHEDULE_MARK).Set(bar_mark)
            section_bar_right.LookupParameter("Rebar r Custom").Set(0)
            section_bar_right.LookupParameter("Rebar Spacing").Set((360.0 / no_bars) / 304.8)
            section_bar_right.LookupParameter("Rebar Quantity").Set(no_bars)

            # rotate the second copy 180 degrees around the vertical axis to mirror it on the other side of the section
            section_bar_left_id = ElementTransformUtils.CopyElement(doc, rebar.Id, XYZ(0, 0, 0))[0]
            section_bar_left = doc.GetElement(section_bar_left_id)
            section_bar_left.Location.Rotate(Line.CreateBound(locPoint, locPoint + XYZ.BasisZ), math.pi)
            section_bar_left.LookupParameter(PARAM_MARK).Set("SECTION TOP RADIAL")
            section_bar_left.LookupParameter(PARAM_SCHEDULE_MARK).Set(bar_mark)
            section_bar_left.LookupParameter("Rebar r Custom").Set(0)
            section_bar_left.LookupParameter("Rebar Spacing").Set((360.0 / no_bars) / 304.8)
            section_bar_left.LookupParameter("Rebar Quantity").Set(no_bars)

            # build radial array
            RotAngle = 2 * math.pi
            elem = RadialArray.ArrayElementWithoutAssociation(
                doc, view, rebar.Id, no_bars,
                Line.CreateBound(locPoint, locPoint + XYZ.BasisZ),
                RotAngle, ArrayAnchorMember.Last)

            # rotate and tag rebar
            for elm in elem:
                el = doc.GetElement(elm)
                el.Location.Rotate(
                    Line.CreateBound(locPoint, locPoint + XYZ.BasisZ),
                    RotSwitch * RotAngle / (no_bars * 2))
                el.LookupParameter(PARAM_MARK).Set("TOP RADIALS")
                el.LookupParameter(PARAM_SCHEDULE_MARK).Set(bar_mark)
                el.LookupParameter("Rebar Quantity").Set(no_bars)
                el.LookupParameter("Rebar r Custom").Set(0)
                el.LookupParameter("Rebar Spacing").Set(360.0/no_bars)

            debug_print('#' * 50)

    t.Commit()

finally:
    workbook.Close(False)
    excel.Quit()









