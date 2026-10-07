from Autodesk.Revit.DB.Structure import *
from Autodesk.Revit.DB.Structure import RebarShape
import math, clr
from Autodesk.Revit.DB import Transaction, Structure, FilteredElementCollector, RadialArray, ArrayAnchorMember
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, Line, XYZ
from Autodesk.Revit.DB import FailureSeverity, FailureProcessingResult, IFailuresPreprocessor
from Autodesk.Revit.DB import *
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
import atexit
from pyrevit import forms

# --- constants ---
MM_TO_FT         = 1.0 / 304.8
TOP_COVER_MM     = 40.0
BOT_COVER_MM     = 50.0
LAP_CONST        = 45
GROUT_OFFSET_MM  = 70.0
MAX_ARC_LEN_MM   = 13000.0
MAX_CHORD_MM     = 2500.0
FULL_CIRCLE_RAD  = 2 * math.pi

doc  = __revit__.ActiveUIDocument.Document
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
    A -> bar mark        e.g. PH100
    B -> empty
    C -> empty
    D -> bar size       e.g. Y16
    E -> empty
    F -> empty
    G -> spacing        e.g. 150

- Bar mark: must start with "PH" to be picked up by the script, followed by the mark number (e.g. 100)
"""


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


def getFilePath():
    fpath = forms.pick_file(file_ext='xlsx', multi_file=False, unc_paths=False)
    if not fpath:
        raise SystemExit
    return fpath


def get_rebar_type(all_rebar_types, size):
    for rebar_type in all_rebar_types:
        name = rebar_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        if name == size:
            return rebar_type
    return None


def create_horizontal_rebar(doc, sc_65, bar_type, wtf_element, curve):
    return Structure.Rebar.CreateFromCurvesAndShape(
        doc, sc_65, bar_type, None, None,
        wtf_element, XYZ.BasisZ, curve,
        RebarHookOrientation.Left, RebarHookOrientation.Left
    )


def set_rebar_params(elem, bar_mark):
    elem.LookupParameter(PARAM_MARK).Set("PLINTH FACE HORIZONTAL")
    elem.LookupParameter(PARAM_SCHEDULE_MARK).Set(bar_mark)


FPath = getFilePath()

excel    = Excel.ApplicationClass()
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

debug_print("BAR MARK  ---  BAR SIZE  ---  SPACING")
debug_print("*" * 50)

# --- collect once, outside all loops ---
rebar_shape_col = FilteredElementCollector(doc).OfClass(RebarShape).WhereElementIsElementType().ToElements()
sc_65 = None
for r_shape in rebar_shape_col:
    if r_shape.LookupParameter("Type Name").AsString() == '65':
        sc_65 = r_shape
        break

all_rebar_types = (FilteredElementCollector(doc)
                   .OfCategory(BuiltInCategory.OST_Rebar)
                   .WhereElementIsElementType()
                   .ToElements())

wtf_collection = (FilteredElementCollector(doc)
                  .OfCategory(BuiltInCategory.OST_GenericModel)
                  .WhereElementIsNotElementType()
                  .ToElements())
wtf_element = None
for element in wtf_collection:
    if element.Name == "1PA_WTF_SteelTower":
        wtf_element = element
        break

wtf_type_id = wtf_element.GetTypeId()
wtf_type    = doc.GetElement(wtf_type_id)
r_base      = wtf_element.LookupParameter("rBase").AsDouble()
h_base      = wtf_element.LookupParameter("hBase").AsDouble()
r_plinth    = wtf_element.LookupParameter("rPlinth").AsDouble()
h_plinth    = wtf_element.LookupParameter("hPlinth").AsDouble()
rTower      = wtf_element.LookupParameter("rTower").AsDouble()
wGroutTop   = wtf_element.LookupParameter("wGroutTop").AsDouble()
r_groutOuter = rTower + (wGroutTop / 2)

wtf_element.Location.Point = XYZ(0, 0, 0)
locPoint = wtf_element.Location.Point

t = Transaction(doc, 'Reinforce')
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SuppressWarnings())
t.SetFailureHandlingOptions(failHandler)
t.Start()

for i in range(2, 200):
    if "PH" in str(xl.Cells(i, 1).Value2).replace(" ", ""):
        debug_print("#" * 45)
        bar_mark = str(xl.Cells(i, 1).Value2)
        bar_size = "Y" + str(xl.Cells(i, 4).Value2)[1:]
        bar_dia  = int(xl.Cells(i, 4).Value2[1:])
        spacing  = int(xl.Cells(i, 7).Value2) * MM_TO_FT

        lap_length = LAP_CONST * bar_dia * MM_TO_FT
        debug_print("Bar Mark: " + bar_mark)
        debug_print("Bar Size: " + bar_size)
        debug_print("Spacing: " + str(spacing / MM_TO_FT))
        debug_print("*" * 50)

        bar_type = get_rebar_type(all_rebar_types, bar_size)
        barDia   = bar_type.LookupParameter("Bar Diameter").AsDouble()

        top_cover = TOP_COVER_MM * MM_TO_FT

        x_outerRange = r_plinth - top_cover - barDia / 2 - spacing / 2
        x_innerRange = r_groutOuter + top_cover + barDia / 2 - GROUT_OFFSET_MM * MM_TO_FT
        x_range      = x_outerRange - x_innerRange
        debug_print("x_outerRange: " + str(x_outerRange / MM_TO_FT))
        debug_print("x_innerRange: " + str(x_innerRange / MM_TO_FT))
        debug_print("x_range: " + str(x_range / MM_TO_FT))
        no_ConRings = math.ceil(x_range / spacing)
        debug_print("No. of concentric rings: " + str(no_ConRings))
        spacing  = x_range / no_ConRings
        debug_print("Spacing: " + str(spacing / MM_TO_FT))

        radius = x_outerRange
        zLevel = h_plinth - top_cover - barDia / 2
        # spliceRot accumulates as radius decreases each iteration
        spliceRot = 2 * lap_length / radius

        mark_prefix   = "".join(c for c in bar_mark if not c.isdigit())
        mark_base_num = int("".join(c for c in bar_mark if c.isdigit()))

        for _ring in range(1, int(no_ConRings) + 1):
            # innermost ring (_ring == no_ConRings) gets the base mark; each outer ring increments by 1
            ring_mark = mark_prefix + str(mark_base_num + int(no_ConRings) - _ring)

            # find no_bars using a bounded loop
            no_bars = 2
            A  = ((math.pi * (radius - barDia / 2) * 2) / no_bars) + lap_length
            r  = radius
            x1 = r - (r * math.cos(A / (2 * r)))
            for _ in range(100):
                if A <= MAX_ARC_LEN_MM * MM_TO_FT and x1 <= MAX_CHORD_MM * MM_TO_FT:
                    break
                no_bars += 1
                A  = ((math.pi * (radius - barDia / 2) * 2) / no_bars) + lap_length
                A  = (round((A * 304.8) / 100) * 100) * MM_TO_FT
                r  = radius
                x1 = r - (r * math.cos(A / (2 * r)))
            else:
                debug_print("bar count loop reached limit")

            debug_print("Number of bars: " + str(no_bars))
            debug_print("A: " + str(A / MM_TO_FT / 1000) + "m")
            debug_print("x1: " + str(round(x1 / MM_TO_FT) / 1000) + "m")
            debug_print("#" * 45)

            # build the construction bar curve
            planeOrigin = locPoint + XYZ(0, 0, zLevel)
            plane       = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, planeOrigin)
            precurve    = [Arc.Create(plane, radius, 0, A / radius)]

            adjValue = barDia * 0.9

            p1    = precurve[0].GetEndPoint(0)
            p2    = XYZ(math.cos(A / (radius * 2)) * radius, math.sin(A / (radius * 2)) * radius, zLevel)
            p3    = precurve[0].GetEndPoint(1)

            p1Vec = p1 - planeOrigin
            p3Vec = p3 - planeOrigin
            p1adj = p1 + p1Vec.Normalize() * adjValue
            p3adj = p3 - p3Vec.Normalize() * adjValue

            curve = [Arc.Create(p1adj, p2, p3adj)]

            rebarCur = create_horizontal_rebar(doc, sc_65, bar_type, wtf_element, curve)
            rebarCur.LookupParameter("A").Set(A)
            rebarCur.LookupParameter("r").Set(r)
            rebarCur.LookupParameter("Rebar r Custom").Set(r)
            rebarCur.LookupParameter(PARAM_MARK).Set("PLINTH FACE HORIZONTAL")
            rebarCur.LookupParameter(PARAM_SCHEDULE_MARK).Set(ring_mark)

            if no_bars > 2:
                arr_elems = RadialArray.ArrayElementWithoutAssociation(
                    doc, view, rebarCur.Id, no_bars,
                    Line.CreateBound(locPoint, XYZ.BasisZ), FULL_CIRCLE_RAD, ArrayAnchorMember.Last
                )
                for elem_id in arr_elems:
                    set_rebar_params(doc.GetElement(elem_id), ring_mark)
                    ElementTransformUtils.RotateElement(
                        doc, elem_id, Line.CreateBound(locPoint, locPoint + XYZ.BasisZ), spliceRot
                    )
            else:
                doc.Delete(rebarCur.Id)
                twobarRebarCur = create_horizontal_rebar(doc, sc_65, bar_type, wtf_element, precurve)
                set_rebar_params(twobarRebarCur, ring_mark)
                rebarCopy = ElementTransformUtils.CopyElement(doc, twobarRebarCur.Id, XYZ(0, 0, barDia))
                ElementTransformUtils.RotateElement(
                    doc, rebarCopy[0], Line.CreateBound(locPoint, XYZ.BasisZ), FULL_CIRCLE_RAD / 2
                )
                set_rebar_params(doc.GetElement(rebarCopy[0]), ring_mark)

            radius    -= spacing
            # intentional accumulation: each ring's splice is offset from the previous
            spliceRot += 2 * lap_length / radius

workbook.Close(False)
excel.Quit()

t.Commit()










