from Autodesk.Revit.DB.Structure import *
import math, clr
from Autodesk.Revit.DB import Transaction, Structure, FilteredElementCollector, RadialArray, ArrayAnchorMember
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, Line, XYZ
from Autodesk.Revit.DB import FailureSeverity, FailureProcessingResult, IFailuresPreprocessor
from pyrevit import forms
from Autodesk.Revit.DB import *
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
import atexit

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
    A -> bar mark             e.g. BC100
    B -> startRadius          e.g. 500
    C -> endRadius            e.g. 1500
    D -> bar size             e.g. Y25
    E -> empty
    F -> empty
    G -> spacing between bars e.g. 150


- Bar mark: must start with "BC" to be picked up by the script, e.g. BC100
"""

#### Constants ####################################################################

MM_TO_FT        = 1.0 / 304.8
TOP_COVER_MM    = 40.0
BOT_COVER_MM    = 50.0
MAX_ARC_LEN_MM  = 13000.0
MAX_CHORD_MM    = 2500.0
LAP_CONST       = 45
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


def concentric_bar(con_barmark, radius, bar_size, top_cover=TOP_COVER_MM * MM_TO_FT, bot_cover=BOT_COVER_MM * MM_TO_FT, lap_length=None):
    if lap_length is None:
        lap_length = LAP_CONST * barDict["BC100"]["bar_dia"]

    bar_type = get_rebar_type(all_rebar_types, bar_size)
    Yoffset  = bot_cover + barDia + 32 * MM_TO_FT

    no_barsStart = 2
    no_bars = no_barsStart
    Acon  = ((math.pi * (radius - barDia / 2) * 2) / no_barsStart) + lap_length
    x1con = radius - (radius * math.cos(Acon / (2 * radius)))
    whilekill = 0
    while Acon > MAX_ARC_LEN_MM * MM_TO_FT or x1con > MAX_CHORD_MM * MM_TO_FT:
        no_barsStart += 1
        no_bars = no_barsStart
        Acon  = ((math.pi * (radius - barDia / 2) * 2) / no_bars) + lap_length
        Acon  = (round((Acon / MM_TO_FT) / 100) * 100) * MM_TO_FT
        x1con = radius - (radius * math.cos(Acon / (2 * radius)))
        whilekill += 1
        debug_print('No_barsStart: ' + str(no_barsStart))
        if whilekill > 100:
            debug_print("while loop killed")
            break
    debug_print("Acon: "    + str(Acon / MM_TO_FT / 1000))
    debug_print("x1con: "   + str(round(x1con / MM_TO_FT) / 1000))
    debug_print("no_bars: " + str(no_bars))
    debug_print("#" * 45)

    planeOrigin = locPoint + XYZ(0, 0, Yoffset)
    plane       = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, planeOrigin)
    precurve    = [Arc.Create(plane, radius, 0, Acon / radius)]
    adjValue    = barDia * 0.5

    p1    = precurve[0].GetEndPoint(0)
    p2    = XYZ(math.cos(Acon / (radius * 2)) * radius, math.sin(Acon / (radius * 2)) * radius, Yoffset)
    p3    = precurve[0].GetEndPoint(1)
    p1Vec = p1 - planeOrigin
    p3Vec = p3 - planeOrigin
    p1adj = p1 + p1Vec.Normalize() * adjValue
    p3adj = p3 - p3Vec.Normalize() * adjValue
    curve = [Arc.Create(p1adj, p2, p3adj)]

    rebarCur = create_concentric_rebar(doc, sc_65, bar_type, wtf_element, curve)

    tot_ss = 1
    if float(endRadius - startRadius) > float(barDict["BC100"]["spacing"] * 5):
        tot_ss = math.floor((endRadius - startRadius) / (spacing * 5))

    if con_barmark is None:
        if Acon > MAX_ARC_LEN_MM * MM_TO_FT or x1con > MAX_CHORD_MM * MM_TO_FT or tot_ss > 1:
            con_barmark = "BC101"
        else:
            con_barmark = "BC100"

    rebarCur.LookupParameter("A").Set(Acon)
    rebarCur.LookupParameter("r").Set(radius)
    rebarCur.LookupParameter(PARAM_MARK).Set("BOTTOM CONCENTRIC")

    if no_bars > 2:
        elem = RadialArray.ArrayElementWithoutAssociation(doc, view, rebarCur.Id, no_bars, Line.CreateBound(locPoint, XYZ.BasisZ), FULL_CIRCLE_RAD, ArrayAnchorMember.Last)
        for elm in elem:
            doc.GetElement(elm).LookupParameter(PARAM_MARK).Set('BOTTOM CONCENTRIC')
            doc.GetElement(elm).LookupParameter(PARAM_SCHEDULE_MARK).Set(con_barmark)
    else:
        doc.Delete(rebarCur.Id)
        twobarRebarCur = create_concentric_rebar(doc, sc_65, bar_type, wtf_element, precurve)
        twobarRebarCur.LookupParameter(PARAM_MARK).Set('BOTTOM CONCENTRIC')
        twobarRebarCur.LookupParameter(PARAM_SCHEDULE_MARK).Set(con_barmark)
        rebarCopy = ElementTransformUtils.CopyElement(doc, twobarRebarCur.Id, XYZ(0, 0, barDia))
        ElementTransformUtils.RotateElement(doc, rebarCopy[0], Line.CreateBound(locPoint, XYZ.BasisZ), FULL_CIRCLE_RAD / 2)
        doc.GetElement(rebarCopy[0]).LookupParameter(PARAM_MARK).Set('BOTTOM CONCENTRIC')
        doc.GetElement(rebarCopy[0]).LookupParameter(PARAM_SCHEDULE_MARK).Set(con_barmark)
    radius -= spacing


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

barDict = {}
for i in range(2, 200):
    if "BC" in str(xl.Cells(i, 1).Value2).replace(" ", ""):
        bar_mark    = str(xl.Cells(i, 1).Value2)
        startRadius = float(xl.Cells(i, 2).Value2) * MM_TO_FT
        endRadius   = float(xl.Cells(i, 3).Value2) * MM_TO_FT
        bar_size    = "Y" + str(xl.Cells(i, 4).Value2)[1:3]
        bar_dia     = int(xl.Cells(i, 4).Value2[1:3]) * MM_TO_FT
        spacing     = int(float(xl.Cells(i, 7).Value2)) * MM_TO_FT
        barDict[bar_mark] = {
            "bar_mark":    bar_mark,
            "startRadius": startRadius,
            "endRadius":   endRadius,
            "bar_size":    bar_size,
            "bar_dia":     bar_dia,
            "spacing":     spacing,
        }

#### Rebar Shape ####################################################################

rebar_shape = FilteredElementCollector(doc).OfClass(RebarShape).WhereElementIsElementType().ToElements()
for r_shape in rebar_shape:
    if r_shape.LookupParameter("Type Name").AsString() == '65':
        sc_65 = r_shape

#### Element Host ####################################################################

wtf_collection = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
wtf_element = None
for element in wtf_collection:
    if element.Name == "1PA_WTF_SteelTower":
        wtf_element = element
        break

r_base = wtf_element.LookupParameter("rBase").AsDouble()

top_cover = TOP_COVER_MM * MM_TO_FT
bot_cover = BOT_COVER_MM * MM_TO_FT

#### Rebar Types (collected once) ####################################################################

all_rebar_types = FilteredElementCollector(doc) \
    .OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsElementType() \
    .ToElements()

#### Start transaction ####################################################################

t = Transaction(doc, 'Reinforce')
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SuppressWarnings())
t.SetFailureHandlingOptions(failHandler)
t.Start()

bar_type = get_rebar_type(all_rebar_types, barDict["BC100"]["bar_size"])
barDia   = bar_type.LookupParameter("Bar Diameter").AsDouble()

wtf_element.Location.Point = XYZ(0, 0, 0)
locPoint = wtf_element.Location.Point

startRadius = barDict["BC100"]["startRadius"]
spacing     = barDict["BC100"]["spacing"]
lap_length  = barDict["BC100"]["bar_dia"] * LAP_CONST
Yoffset     = bot_cover + barDict["BC100"]["bar_dia"] / 2 + 32 * MM_TO_FT

A  = 13500 * MM_TO_FT
x1 = startRadius - (startRadius * math.cos(A / (2 * startRadius)))
while A > MAX_ARC_LEN_MM * MM_TO_FT or x1 > MAX_CHORD_MM * MM_TO_FT:
    A  -= 100 * MM_TO_FT
    x1  = startRadius - (startRadius * math.cos(A / (2 * startRadius)))
debug_print(" start A: " + str(A / MM_TO_FT / 1000))
debug_print("start x1: " + str(round(x1 / MM_TO_FT) / 1000))
debug_print("startRad:  " + str(startRadius / MM_TO_FT))

r1    = startRadius + barDict["BC100"]["bar_dia"]
r3    = ((startRadius + math.sqrt(startRadius**2 + 4 * (spacing / (2 * math.pi)) * A)) / 2) - barDict["BC100"]["bar_dia"]
r2    = (r1 + r3) / 2
r3spl = (startRadius + math.sqrt(startRadius**2 + 4 * (spacing / (2 * math.pi)) * (A - lap_length))) / 2
theta1    = 0
theta3    = A / r2 + theta1
theta2    = (theta3 - theta1) / 2 + theta1
theta3spl = (A - lap_length) / r2

p1 = XYZ(polar_to_car(r1, theta1)[0], polar_to_car(r1, theta1)[1], Yoffset)
p2 = XYZ(polar_to_car(r2, theta2)[0], polar_to_car(r2, theta2)[1], Yoffset)
p3 = XYZ(polar_to_car(r3, theta3)[0], polar_to_car(r3, theta3)[1], Yoffset)
preCurve  = Arc.Create(p1, p3, p2)
bar_type  = get_rebar_type(all_rebar_types, barDict["BC100"]["bar_size"])
rebarCur1 = create_concentric_rebar(doc, sc_65, bar_type, wtf_element, [preCurve])
rebarCur1.LookupParameter(PARAM_MARK).Set('BOTTOM CONCENTRIC')

skipped_rows = []

# Sort by startRadius: r2 carries over between rows and only ever grows, so the rows
# must be walked outward.  Dict order is not the sheet order under IronPython.
for entry in sorted(barDict.values(), key=lambda e: e["startRadius"]):
    bar_mark    = entry["bar_mark"]
    startRadius = entry["startRadius"]
    endRadius   = entry["endRadius"]
    bar_size    = entry["bar_size"]
    bar_dia     = entry["bar_dia"]
    spacing     = entry["spacing"]

    if endRadius > r_base:
        endRadius = r_base - bot_cover - barDia * 2.5
        debug_print("Adjusted endRadius: " + str(endRadius / MM_TO_FT))

    bar_type = get_rebar_type(all_rebar_types, bar_size)
    barDia   = bar_type.LookupParameter("Bar Diameter").AsDouble()

    lap_length   = LAP_CONST * bar_dia
    Yoffset      = bot_cover + barDia / 2 + 32 * MM_TO_FT
    radius_limit = endRadius - bot_cover - (barDia * 1.5)

    # If the spiral has already run past this row's outer limit, both while loops below
    # are false on entry and the row produces nothing at all.  Report it rather than
    # letting the bar mark vanish silently.
    if r2 >= radius_limit:
        skipped_rows.append(bar_mark)
        debug_print("SKIPPED {} -- r2 ({:.0f}) is already past its limit ({:.0f})".format(
            bar_mark, r2 / MM_TO_FT, radius_limit / MM_TO_FT))
        continue

    debug_print("*" * 45)
    debug_print("bar_mark: "    + str(bar_mark))
    debug_print("startRadius: " + str(startRadius / MM_TO_FT))
    debug_print("endRadius: "   + str(endRadius / MM_TO_FT))
    debug_print("bar_size: "    + str(bar_size))
    debug_print("bar_dia: "     + str(bar_dia / MM_TO_FT))
    debug_print("spacing: "     + str(spacing / MM_TO_FT))
    debug_print("lap_length: "  + str(lap_length / MM_TO_FT))
    debug_print("*" * 45)

    if startRadius > 6000 * MM_TO_FT and barDia < 32 * MM_TO_FT:
        A = MAX_ARC_LEN_MM * MM_TO_FT

    tot_subset  = 1
    if endRadius - startRadius > spacing * 3:
        tot_subset = math.floor((endRadius - startRadius) / (spacing * 3))
        debug_print("subset: " + str(tot_subset))
    subset      = 1
    subsetRange = (endRadius - startRadius) / tot_subset

    if A < MAX_ARC_LEN_MM * MM_TO_FT:
        while r2 < radius_limit:
            while r2 < startRadius + subsetRange * subset:
                r1    = r3spl - (bar_dia / 2)
                r3    = (r1 + math.sqrt(r1**2 + 4 * (spacing / (2 * math.pi)) * (A - lap_length))) / 2
                r2    = (r1 + r3) / 2
                r3spl = ((r1 + math.sqrt(r1**2 + 4 * (spacing / (2 * math.pi)) * (A - lap_length))) + bar_dia) / 2
                theta1    = theta3spl
                theta3    = theta1 + A / r2
                theta2    = (theta3 - theta1) / 2 + theta1
                theta3spl = (A - lap_length) / r2 + theta1
                p1 = XYZ(polar_to_car(r1, theta1)[0], polar_to_car(r1, theta1)[1], Yoffset)
                p2 = XYZ(polar_to_car(r2, theta2)[0], polar_to_car(r2, theta2)[1], Yoffset)
                p3 = XYZ(polar_to_car(r3, theta3)[0], polar_to_car(r3, theta3)[1], Yoffset)
                preCurve = Arc.Create(p1, p3, p2)
                rebarCur = create_concentric_rebar(doc, sc_65, bar_type, wtf_element, [preCurve])
                rebarCur.LookupParameter(PARAM_MARK).Set('BOTTOM CONCENTRIC')
                rebarCur.LookupParameter(PARAM_SCHEDULE_MARK).Set("BC" + str(int(str(bar_mark)[2:]) + subset - 1))

            subset += 1
            A_live  = (theta3 - theta1) * r2
            x1_live = r2 - (r2 * math.cos(A / (2 * r2)))
            while A_live < MAX_ARC_LEN_MM * MM_TO_FT and x1_live < MAX_CHORD_MM * MM_TO_FT:
                A      += 100 * MM_TO_FT
                x1      = r2 - (r2 * math.cos(A / (2 * r2)))
                A_live  = A
                x1_live = x1
            if r2 > 6000 * MM_TO_FT and barDia < 25 * MM_TO_FT:
                A = MAX_ARC_LEN_MM * MM_TO_FT

    elif A > 12999 * MM_TO_FT and A < 13001 * MM_TO_FT:
        while r2 < radius_limit:
            r1    = r3spl - (bar_dia / 2)
            r3    = (r1 + math.sqrt(r1**2 + 4 * (spacing / (2 * math.pi)) * (A - lap_length))) / 2
            r2    = (r1 + r3) / 2
            r3spl = ((r1 + math.sqrt(r1**2 + 4 * (spacing / (2 * math.pi)) * (A - lap_length))) + bar_dia) / 2
            theta1    = theta3spl
            theta3    = theta1 + A / r2
            theta2    = (theta3 - theta1) / 2 + theta1
            theta3spl = (A - lap_length) / r2 + theta1
            p1 = XYZ(polar_to_car(r1, theta1)[0], polar_to_car(r1, theta1)[1], Yoffset)
            p2 = XYZ(polar_to_car(r2, theta2)[0], polar_to_car(r2, theta2)[1], Yoffset)
            p3 = XYZ(polar_to_car(r3, theta3)[0], polar_to_car(r3, theta3)[1], Yoffset)
            preCurve = Arc.Create(p1, p3, p2)
            rebarCur = create_concentric_rebar(doc, sc_65, bar_type, wtf_element, [preCurve])
            rebarCur.LookupParameter(PARAM_MARK).Set('BOTTOM CONCENTRIC')
            rebarCur.LookupParameter(PARAM_SCHEDULE_MARK).Set(bar_mark)

        while A < MAX_ARC_LEN_MM * MM_TO_FT and x1 < MAX_CHORD_MM * MM_TO_FT:
            A  += 100 * MM_TO_FT
            x1  = r2 - (r2 * math.cos(A / (2 * r2)))
        debug_print(" start A: " + str(A / MM_TO_FT / 1000))
        debug_print("start x1: " + str(round(x1 / MM_TO_FT) / 1000))
        debug_print("subset: "   + str(subset))

outConDia    = max(str(bm)[2:] for bm in barDict)
barDia       = barDict["BC" + str(outConDia)]["bar_dia"]
outConRadius = r_base - bot_cover - barDia * 0.5
debug_print('#' * 100)
debug_print("outConDia: "    + str(outConDia))
debug_print("outConRadius: " + str(outConRadius / MM_TO_FT))
debug_print('#' * 100)
startRadius = outConRadius

outCon = concentric_bar("BC" + str(outConDia), outConRadius, barDict["BC" + str(outConDia)]["bar_size"],
                        top_cover=top_cover, bot_cover=bot_cover, lap_length=LAP_CONST * barDia)
concentric_bar(None, barDict["BC100"]["startRadius"], barDict["BC100"]["bar_size"])
doc.Delete(rebarCur1.Id)

t.Commit()

if skipped_rows:
    forms.alert(
        "No bars were generated for: {}\n\n"
        "The spiral had already passed these rows' outer radius, so that zone is "
        "covered by the neighbouring bar mark at the neighbour's spacing and bar size. "
        "Check the schedule and the row radii in the excel sheet.".format(", ".join(skipped_rows)),
        title="Rows skipped")

#### Post-process ####################################################################

FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
scheduleMarkList = list({elem.LookupParameter(PARAM_SCHEDULE_MARK).AsString() for elem in FEC})

t = Transaction(doc, "Update r")
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SuppressWarnings())
t.SetFailureHandlingOptions(failHandler)
t.Start()
for mark in scheduleMarkList:
    if "BC" in mark:
        sum_r   = 0
        count_r = 0
        for elem in FEC:
            if elem.LookupParameter(PARAM_SCHEDULE_MARK).AsString() == mark:
                sum_r   += elem.LookupParameter("r").AsDouble()
                count_r += 1
        r = round((sum_r / count_r) * 100) / 100
        r = round(r / MM_TO_FT) * MM_TO_FT
        # print(mark)
        # print(r / MM_TO_FT)
        for elem in FEC:
            if elem.LookupParameter(PARAM_SCHEDULE_MARK).AsString() == mark:
                elem.LookupParameter("Rebar r Custom").Set(r)
        # print("updated r")
t.Commit()

t = Transaction(doc, "Update A")
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SuppressWarnings())
t.SetFailureHandlingOptions(failHandler)
t.Start()
for mark in scheduleMarkList:
    if "BC" in mark:
        sum_A   = 0
        count_A = 0
        A_max   = 0
        for elem in FEC:
            if elem.LookupParameter(PARAM_SCHEDULE_MARK).AsString() == mark:
                A_val   = elem.LookupParameter("A").AsDouble()
                sum_A   += A_val
                count_A += 1
                if A_val > A_max:
                    A_max = A_val
        A = round((sum_A / count_A) * 100) / 100
        A = round(A / MM_TO_FT) * MM_TO_FT
        # print(mark)
        # print(A / MM_TO_FT)
        # print(A_max / MM_TO_FT)
        for elem in FEC:
            if elem.LookupParameter(PARAM_SCHEDULE_MARK).AsString() == mark:
                elem.LookupParameter("A").Set((round((A_max / MM_TO_FT) / 100) * 100) * MM_TO_FT)
t.Commit()

workbook.Close(False)
excel.Quit()
print ("DONE")
debug_print('#' * 100)











