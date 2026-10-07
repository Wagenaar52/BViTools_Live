from Autodesk.Revit.DB.Structure import *
from Autodesk.Revit.DB.Structure import RebarShape
import math, clr
from Autodesk.Revit.DB import Transaction, Structure, FilteredElementCollector, RadialArray, ArrayAnchorMember
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, Line, XYZ
from Autodesk.Revit.DB import FailureSeverity, FailureProcessingResult, IFailuresPreprocessor
from pyrevit import forms
from Autodesk.Revit.DB import *
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
import atexit

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
view = uidoc.ActiveView

DEBUG = True
PARAM_MARK = "Mark"
PARAM_SCHEDULE_MARK = "Schedule Mark"
PARAM_A_CANDIDATES = ["A", "\u00c4"]

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
  A -> bar mark        e.g. PC100
  B -> radius (mm)     e.g. 500
  C -> Y offset (mm)  e.g. 0
  D -> bar size       e.g. Y16
  E -> (empty)
  F -> lap orientation (H or V) e.g. H


- Bar mark: must start with "PC" and be unique for each row, it will be assigned to the "Schedule Mark" parameter of the rebar

"""

#### Constants ####################################################################

MM_TO_FT        = 1.0 / 304.8
TOP_COVER_MM    = 40.0
BOT_COVER_MM    = 50.0
LAP_CONST       = 45
MAX_ARC_LEN_MM  = 13000.0
MAX_CHORD_MM    = 2500.0
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


def getFilePath():
    FPath = forms.pick_file(file_ext='xlsx', multi_file=False, unc_paths=False)
    if not FPath:
        raise SystemExit
    return FPath


def polar_to_car(radius, angle_radians):
    x = radius * math.cos(angle_radians)
    y = radius * math.sin(angle_radians)
    return x, y


def get_rebar_type(all_rebar_types, size):
    for rebar_type in all_rebar_types:
        if rebar_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == size:
            return rebar_type
    return None


def create_concentric_rebar(doc, sc_65, bar_type, wtf_element, curve):
    return Structure.Rebar.CreateFromCurvesAndShape(
        doc, sc_65, bar_type, None, None, wtf_element, XYZ.BasisZ,
        curve, RebarHookOrientation.Left, RebarHookOrientation.Left)


def set_rebar_params(elem, bar_mark, r):
    elem.LookupParameter(PARAM_MARK).Set("PLINTH CONCENTRIC")
    elem.LookupParameter(PARAM_SCHEDULE_MARK).Set(bar_mark)
    elem.LookupParameter("Rebar r Custom").Set(r)
    elem.LookupParameter("Comments").Set(elem.LookupParameter(PARAM_SCHEDULE_MARK).AsValueString()[-2:])


def _get_first_param(elem, names):
    for name in names:
        p = elem.LookupParameter(name)
        if p:
            return p
    return None


def average_a_by_schedule_mark_and_set(doc):
    grouped = {}
    rebar_elems = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()

    for rb in rebar_elems:
        mark_param = rb.LookupParameter(PARAM_MARK)
        if not mark_param:
            continue
        if (mark_param.AsString() or "") != "PLINTH CONCENTRIC":
            continue

        schedule_param = rb.LookupParameter(PARAM_SCHEDULE_MARK)
        a_param = _get_first_param(rb, PARAM_A_CANDIDATES)
        if not schedule_param or not a_param:
            continue
        if a_param.IsReadOnly:
            continue

        schedule_mark = schedule_param.AsString()
        if not schedule_mark:
            continue

        a_val = a_param.AsDouble()
        grouped.setdefault(schedule_mark, []).append((rb, a_param, a_val))

    for schedule_mark, items in grouped.items():
        if not items:
            continue

        avg_ft = sum(x[2] for x in items) / float(len(items))
        avg_mm = avg_ft / MM_TO_FT
        rounded_mm = round(avg_mm / 10.0) * 10.0
        rounded_ft = rounded_mm * MM_TO_FT

        for _, a_param, _ in items:
            a_param.Set(rounded_ft)

        debug_print("Normalized A for", schedule_mark, "count=", len(items), "avg_mm=", round(avg_mm, 3), "set_mm=", rounded_mm)

def place_rebar_array(doc, view, rebarCur, no_bars, sc_65, bar_type, wtf_element,
                      curve, precurve, barDia, p_0, bar_mark, r, spliceRot):
    if no_bars > 2:
        elem = RadialArray.ArrayElementWithoutAssociation(
            doc, view, rebarCur.Id, no_bars,
            Line.CreateBound(p_0, XYZ.BasisZ), FULL_CIRCLE_RAD, ArrayAnchorMember.Last)
        for elm in elem:
            el = doc.GetElement(elm)
            set_rebar_params(el, bar_mark, r)
            ElementTransformUtils.RotateElement(doc, elm, Line.CreateBound(p_0, p_0 + XYZ.BasisZ), spliceRot)
    else:
        doc.Delete(rebarCur.Id)
        twobarRebarCur = create_concentric_rebar(doc, sc_65, bar_type, wtf_element, precurve)
        rebarCopy = ElementTransformUtils.CopyElement(doc, twobarRebarCur.Id, XYZ(0, 0, barDia))
        ElementTransformUtils.RotateElement(doc, rebarCopy[0], Line.CreateBound(p_0, XYZ.BasisZ), FULL_CIRCLE_RAD / 2)
        set_rebar_params(doc.GetElement(rebarCopy[0]), bar_mark, r)
        set_rebar_params(twobarRebarCur, bar_mark, r)


#### Input from excel sheet ####################################################################

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

#### Rebar Shape (collected once) ####################################################################

rebar_shape = FilteredElementCollector(doc).OfClass(RebarShape).WhereElementIsElementType().ToElements()
sc_65 = None
for r_shape in rebar_shape:
    if r_shape.LookupParameter("Type Name").AsString() == '65':
        sc_65 = r_shape
        break

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

wtf_type_id = wtf_element.GetTypeId()
wtf_type    = doc.GetElement(wtf_type_id)
r_base      = wtf_element.LookupParameter("rBase").AsDouble()
h_base      = wtf_element.LookupParameter("hBase").AsDouble()
r_plinth    = wtf_element.LookupParameter("rPlinth").AsDouble()
h_plinth    = wtf_element.LookupParameter("hPlinth").AsDouble()

wtf_element.Location.Point = XYZ(0, 0, 0)
locPoint = wtf_element.Location.Point
p_0      = locPoint

debug_print(str("RADIUS") + "  ---   " + "BAR MARK" + "  ---   " + str('NO. BARS') + "  ---  " + 'BAR SIZE')
debug_print("*" * 45)

# spliceRot accumulates intentionally across rows to offset splices between successive bars
spliceRot = 0

t = Transaction(doc, 'Reinforce')
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SuppressWarnings())
t.SetFailureHandlingOptions(failHandler)
t.Start()

for i in range(5, 200):
    if "PC" in str(xl.Cells(i, 1).Value2):
        radius         = float(xl.Cells(i, 2).Value2) * MM_TO_FT
        bar_mark       = str(xl.Cells(i, 1).Value2)
        Yoffset        = int(xl.Cells(i, 3).Value2) * MM_TO_FT
        bar_size       = "Y" + str(xl.Cells(i, 4).Value2)[1:3]
        bar_dia        = int(xl.Cells(i, 4).Value2[1:3])
        lapOrientation = str(xl.Cells(i, 6).Value2).upper()

        lap_length = LAP_CONST * bar_dia * MM_TO_FT
        debug_print(str(radius / MM_TO_FT) + "  ---  \t " + bar_mark + "  --- \t \t ##  ---\t \t " + bar_size)
        debug_print("*" * 45)

        bar_type = get_rebar_type(all_rebar_types, bar_size)
        barDia   = bar_type.LookupParameter("Bar Diameter").AsDouble()

        ##### Calculate arc length and number of bars ####################################################################

        no_bars = 2
        A       = ((math.pi * (radius - barDia / 2) * 2) / no_bars) + lap_length
        r       = radius
        x1      = r - (r * math.cos(A / (2 * r)))
        whilekill = 0
        while A > MAX_ARC_LEN_MM * MM_TO_FT or x1 > MAX_CHORD_MM * MM_TO_FT:
            no_bars += 1
            A  = ((math.pi * (radius - barDia / 2) * 2) / no_bars) + lap_length
            r  = radius
            x1 = r - (r * math.cos(A / (2 * r)))
            whilekill += 1
            if whilekill > 100:
                debug_print("while loop killed")
                break

        debug_print("Number of bars: " + str(no_bars))
        debug_print("A: "  + str(A / MM_TO_FT / 1000))
        debug_print("x1: " + str(round(x1 / MM_TO_FT) / 1000))
        debug_print("#" * 45)

        ##### Build ####################################################################

        if lapOrientation == "H":
            adjValue = barDia * 0.45
            p1 = XYZ(polar_to_car(radius - adjValue, 0)[0],              polar_to_car(radius - adjValue, 0)[1],              Yoffset)
            p2 = XYZ(polar_to_car(radius, A / (2 * radius))[0],          polar_to_car(radius, A / (2 * radius))[1],          Yoffset)
            p3 = XYZ(polar_to_car(radius + adjValue, A / radius)[0],     polar_to_car(radius + adjValue, A / radius)[1],     Yoffset)
            curve    = [Arc.Create(p1, p3, p2)]
            precurve = curve

        elif lapOrientation == "V":
            preplane  = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, locPoint + XYZ(0, 0, Yoffset))
            precurve  = [Arc.Create(preplane, radius, 0, A / radius)]
            adjValue  = (barDia) * (0.45 * lap_length / A)
            totAdjValue = (adjValue + barDia) / 2
            p1        = precurve[0].GetEndPoint(0) + XYZ(0, 0, totAdjValue)
            p2        = precurve[0].GetEndPoint(1) - XYZ(0, 0, totAdjValue)
            midplane  = Plane.CreateByThreePoints(p1, p2, locPoint + XYZ(0, 0, Yoffset))
            plane     = Plane.CreateByNormalAndOrigin(midplane.Normal, locPoint + XYZ(0, 0, Yoffset))
            curve     = [Arc.Create(plane, radius, 0, A / radius)]

        else:
            debug_print("Unknown lapOrientation: " + lapOrientation + "  skipping " + bar_mark)
            continue

        rebarCur = create_concentric_rebar(doc, sc_65, bar_type, wtf_element, curve)
        rebarCur.LookupParameter("A").Set(A)
        rebarCur.LookupParameter("Rebar r Custom").Set(r)
        rebarCur.LookupParameter(PARAM_MARK).Set("PLINTH CONCENTRIC")
    
        print("Placing array for " + bar_mark)
        place_rebar_array(doc, view, rebarCur, no_bars, sc_65, bar_type, wtf_element,
                          curve, precurve, barDia, p_0, bar_mark, r, spliceRot)

        spliceRot += 2 * lap_length / radius  # intentional accumulation to stagger splices
        debug_print("#" * 50)






average_a_by_schedule_mark_and_set(doc)
t.Commit()

workbook.Close(False)
excel.Quit()