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
MM_TO_FT        = 1.0 / 304.8
TOP_COVER_MM    = 40.0
BOT_COVER_MM    = 50.0
LAP_CONST       = 45
MAX_ARC_LEN_MM  = 13000.0
MAX_CHORD_MM    = 2500.0
FULL_CIRCLE_RAD = 2 * math.pi

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
  A -> bar mark        e.g. PF100
    B -> empty
    C -> empty
    D -> bar size       e.g. Y16
    E -> empty
    F -> empty
    G -> spacing        e.g. 150
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


def create_face_rebar(doc, sc_65, bar_type, wtf_element, curve):
    return Structure.Rebar.CreateFromCurvesAndShape(
        doc, sc_65, bar_type, None, None,
        wtf_element, XYZ.BasisZ, curve,
        RebarHookOrientation.Left, RebarHookOrientation.Left
    )


def set_rebar_params(elem, bar_mark):
    elem.LookupParameter(PARAM_MARK).Set("PLINTH FACE CONCENTRIC")
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

debug_print("RADIUS  ---  BAR MARK  ---  NO. BARS  ---  BAR SIZE")
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
hCone       = wtf_element.LookupParameter("hCone").AsDouble()

wtf_element.Location.Point = XYZ(0, 0, 0)
locPoint = wtf_element.Location.Point
p_0      = locPoint

t = Transaction(doc, 'Reinforce')
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SuppressWarnings())
t.SetFailureHandlingOptions(failHandler)
t.Start()

for i in range(2, 200):
    if str(xl.Cells(i, 1).Value2).replace(" ", "") == 'PF100':
        bar_mark = str(xl.Cells(i, 1).Value2)
        bar_size = "Y" + str(xl.Cells(i, 4).Value2)[1:3]
        bar_dia  = int(xl.Cells(i, 4).Value2[1:3])
        spacing  = int(xl.Cells(i, 7).Value2) * MM_TO_FT
        debug_print(bar_mark + "  ---  " + bar_size)
        debug_print("*" * 50)

        lap_length = LAP_CONST * bar_dia * MM_TO_FT

        bar_type = get_rebar_type(all_rebar_types, bar_size)
        barDia   = bar_type.LookupParameter("Bar Diameter").AsDouble()

        top_cover = TOP_COVER_MM * MM_TO_FT

        y_topRange  = h_plinth - top_cover - barDia * 1.5
        y_Rangebot  = hCone
        y_range     = y_topRange - y_Rangebot
        no_ConRings = math.ceil(y_range / spacing)
        spacing     = y_range / (no_ConRings - 1)
        radius      = r_plinth - top_cover
        Yoffset     = y_Rangebot
        # spliceRot accumulates so each ring's splice is offset from the previous
        spliceRot = 2 * lap_length / radius

        for _ring in range(1, int(no_ConRings) + 1):

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
            debug_print("A: " + str(A * 304.8 / 1000))
            debug_print("x1: " + str(round(x1 * 304.8) / 1000))
            debug_print("#" * 45)

            # build the construction bar curve
            preplane  = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, locPoint + XYZ(0, 0, Yoffset))
            precurve  = [Arc.Create(preplane, radius, 0, A / radius)]

            adjValue    = barDia * (1.45 * lap_length / A)
            totAdjValue = (adjValue + barDia) / 2
            p1          = precurve[0].GetEndPoint(0) + XYZ(0, 0, totAdjValue)
            p2          = precurve[0].GetEndPoint(1) - XYZ(0, 0, totAdjValue)
            midplane    = Plane.CreateByThreePoints(p1, p2, locPoint + XYZ(0, 0, Yoffset))
            plane       = Plane.CreateByNormalAndOrigin(midplane.Normal, locPoint + XYZ(0, 0, Yoffset))
            curve       = [Arc.Create(plane, radius, 0, A / radius)]

            rebarCur = create_face_rebar(doc, sc_65, bar_type, wtf_element, curve)
            rebarCur.LookupParameter("A").Set(A)
            rebarCur.LookupParameter("r").Set(r)
            rebarCur.LookupParameter("Rebar r Custom").Set(0)
            rebarCur.LookupParameter(PARAM_MARK).Set(bar_mark)

            if no_bars > 2:
                arr_elems = RadialArray.ArrayElementWithoutAssociation(
                    doc, view, rebarCur.Id, no_bars,
                    Line.CreateBound(p_0, XYZ.BasisZ), FULL_CIRCLE_RAD, ArrayAnchorMember.Last
                )
                for elem_id in arr_elems:
                    set_rebar_params(doc.GetElement(elem_id), bar_mark)
                    ElementTransformUtils.RotateElement(
                        doc, elem_id, Line.CreateBound(p_0, p_0 + XYZ.BasisZ), spliceRot
                    )
            else:
                doc.Delete(rebarCur.Id)
                twobarRebarCur = create_face_rebar(doc, sc_65, bar_type, wtf_element, precurve)
                rebarCopy = ElementTransformUtils.CopyElement(doc, twobarRebarCur.Id, XYZ(0, 0, barDia))
                ElementTransformUtils.RotateElement(
                    doc, rebarCopy[0], Line.CreateBound(p_0, XYZ.BasisZ), FULL_CIRCLE_RAD / 2
                )
                set_rebar_params(doc.GetElement(rebarCopy[0]), bar_mark)

            Yoffset   += spacing
            # intentional accumulation: each ring's splice is offset from the last
            spliceRot += 2 * lap_length / radius

workbook.Close(False)
excel.Quit()

t.Commit()










