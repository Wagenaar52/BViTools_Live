from Autodesk.Revit.DB.Structure import * 
from Autodesk.Revit.DB.Structure import RebarShape
import math, clr
from Autodesk.Revit.DB import Transaction, Structure, FilteredElementCollector, RadialArray, ArrayAnchorMember
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, Line, XYZ
from Autodesk.Revit.DB import FailureSeverity, FailureProcessingResult,IFailuresPreprocessor
from pyrevit import forms
from Autodesk.Revit.DB import *
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
import atexit
import Functions as func

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
    A -> bar mark             e.g. TC100
    B -> startRadius          e.g. 500
    C -> endRadius            e.g. 1500
    D -> bar size             e.g. Y25
    E -> empty
    F -> empty
    G -> spacing between bars e.g. 150


- Bar mark: must start with "TC" to be picked up by the script, e.g. TC100
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

def polar_to_car(radius, angle_radians):
    x = radius * math.cos(angle_radians)
    y = radius * math.sin(angle_radians)
    return x, y

# FPath =   forms.pick_file(file_ext='xlsx', multi_file=False, unc_paths=False)
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


lapConst = 55

# get all inputs from excel in a dictionary
barDict = {}
for i in range(2, 101):
    if "TC" in str(xl.Cells(i, 1).Value2).replace(" ", ""):
        bar_mark    = str(xl.Cells(i, 1).Value2)
        startRadius = float(xl.Cells(i, 2).Value2) / 304.8
        endRadius   = float(xl.Cells(i, 3).Value2) / 304.8
        bar_size    = "Y" + str(xl.Cells(i, 4).Value2)[1:3]
        bar_dia     = int(xl.Cells(i, 4).Value2[1:3]) / 304.8
        spacing     = int(xl.Cells(i, 7).Value2) / 304.8
        barDict[bar_mark] = {
            "bar_mark":    bar_mark,
            "startRadius": startRadius,
            "endRadius":   endRadius,
            "bar_size":    bar_size,
            "bar_dia":     bar_dia,
            "spacing":     spacing,
        }

excel.Quit()

    
#####Rebar Shape ################################################################

rebar_shape = FilteredElementCollector(doc).OfClass(RebarShape).WhereElementIsElementType().ToElements()  
for r_shape in rebar_shape:
    if r_shape.LookupParameter("Type Name").AsString() == '65':
        sc_65 = r_shape




##### Element Host ###############################################################

WTF = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
for element in WTF:
    if element.Name == "1PA_WTF_SteelTower":
        WTF = element
        break

type_id  = WTF.GetTypeId()
wtf_type = doc.GetElement(type_id)
r_base = WTF.LookupParameter("rBase").AsDouble()
h_base = WTF.LookupParameter("hBase").AsDouble()
r_plinth = WTF.LookupParameter("rPlinth").AsDouble()
h_plinth = WTF.LookupParameter("hPlinth").AsDouble()
hCone = WTF.LookupParameter("hCone").AsDouble()
slabSlope = (hCone - h_base)/(r_base - r_plinth)

def radius_Yoff(radius):
    if radius > r_plinth-100/304.8:
        return ((r_base - radius)*slabSlope)+h_base
    debug_print("radius is less than plinth")
    return 0.0



#Start a transaction ####################################################################   
t = Transaction(doc, 'Reinforce')
t.Start()
# Rebar type #######################################################
all_rebar_types = FilteredElementCollector(doc) \
    .OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsElementType() \
    .ToElements()

def get_rebar_type(bar_size_name):
    for rtype in all_rebar_types:
        if rtype.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == bar_size_name:
            return rtype
    return None

bar_type = get_rebar_type(bar_size)
barDia = bar_type.LookupParameter("Bar Diameter").AsDouble()

# set bar WTF location to origin
WTF.Location.Point = XYZ(0,0,0)
locPoint = WTF.Location.Point
##### Build single concentric bar at start radius############################################################################
startRadius = barDict["TC100"]["startRadius"]
bar_size = barDict["TC100"]["bar_size"]
radius = startRadius
conYoffset = hCone-40/304.8

def concentric_bar(con_barmark, conYoffset, radius, bar_size, spacing,
                   endRadius=None, top_cover=40/304.8, bot_cover=50/304.8, lap_length=None):

    bar_type     = get_rebar_type(bar_size)
    barDia_local = bar_type.LookupParameter("Bar Diameter").AsDouble()
    if lap_length is None:
        lap_length = 45 * barDia_local
    if endRadius is None:
        endRadius = radius

    Yoffset = bot_cover + barDia_local + 32/304.8

    # determine number of bars
    no_bars = 2
    Acon  = ((math.pi*(radius - barDia_local/2)*2)/no_bars) + lap_length
    x1con = radius - (radius*math.cos(Acon/(2*radius)))
    whilekill = 0
    while Acon > 13000/304.8 or x1con > 2500/304.8:
        no_bars += 1
        Acon  = ((math.pi*(radius - barDia_local/2)*2)/no_bars) + lap_length
        Acon  = (round((Acon*304.8)/100)*100)/304.8
        x1con = radius - (radius*math.cos(Acon/(2*radius)))
        whilekill += 1
        if whilekill > 100:
            debug_print("while loop killed")
            break
    debug_print("Acon: "    + str(Acon*304.8/1000))
    debug_print("x1con: "   + str(round(x1con*304.8)/1000))
    debug_print("no_bars: " + str(no_bars))
    debug_print("#"*45)

    # draw arc to place the bar
    planeOrigin = locPoint + XYZ(0, 0, conYoffset)
    plane    = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, planeOrigin)
    precurve = [Arc.Create(plane, radius, 0, Acon/radius)]

    p1 = precurve[0].GetEndPoint(0)
    p2 = XYZ(math.cos(Acon/(radius*2))*radius, math.sin(Acon/(radius*2))*radius, conYoffset)
    p3 = precurve[0].GetEndPoint(1)

    curve = [Arc.Create(p1, p2, p3)]

    #build construction bar
    rebarCur = Structure.Rebar.CreateFromCurvesAndShape(doc, sc_65, bar_type, None, None, WTF, XYZ.BasisZ, curve, RebarHookOrientation.Left, RebarHookOrientation.Left)

    # determine bar mark
    tot_ss = 1
    if float(endRadius - radius) > float(spacing * 5):
        tot_ss = math.floor((endRadius - radius) / (spacing * 5))

    if con_barmark is None:
        if Acon > 13000/304.8 or x1con > 2500/304.8 or tot_ss > 1:
            con_barmark = "TC101"
        else:
            con_barmark = "TC100"

    # set construction bar properties
    rebarCur.LookupParameter("A").Set(Acon)
    rebarCur.LookupParameter("r").Set(radius)
    rebarCur.LookupParameter(PARAM_MARK).Set("TOP CONCENTRIC")

    #build radial array
    RotAngle = 360*math.pi/180
    if no_bars > 2:
        elems = RadialArray.ArrayElementWithoutAssociation(doc, view, rebarCur.Id, no_bars, Line.CreateBound(locPoint, XYZ.BasisZ), RotAngle, ArrayAnchorMember.Last)
        for elem in elems:
            doc.GetElement(elem).LookupParameter(PARAM_MARK).Set("TOP CONCENTRIC")
            doc.GetElement(elem).LookupParameter(PARAM_SCHEDULE_MARK).Set(con_barmark)
    else:
        doc.Delete(rebarCur.Id)
        twobarRebarCur = Structure.Rebar.CreateFromCurvesAndShape(doc, sc_65, bar_type, None, None, WTF, XYZ.BasisZ, precurve, RebarHookOrientation.Left, RebarHookOrientation.Left)
        rebarCopy = ElementTransformUtils.CopyElement(doc, twobarRebarCur.Id, XYZ(0, 0, barDia_local))
        elem = ElementTransformUtils.RotateElement(doc, rebarCopy[0], Line.CreateBound(locPoint, XYZ.BasisZ), RotAngle/2)
        doc.GetElement(rebarCopy[0]).LookupParameter(PARAM_SCHEDULE_MARK).Set(con_barmark)
        doc.GetElement(rebarCopy[0]).LookupParameter(PARAM_MARK).Set("TOP CONCENTRIC")

    failHandler = t.GetFailureHandlingOptions()
    failHandler.SetFailuresPreprocessor(SuppressWarnings())
    t.SetFailureHandlingOptions(failHandler)

concentric_bar(None, conYoffset, barDict["TC100"]["startRadius"], barDict["TC100"]["bar_size"],
               barDict["TC100"]["spacing"], endRadius=barDict["TC100"]["endRadius"])

#######################################################################################################################################################
startRadius = barDict["TC100"]["startRadius"]
lap_length = barDict["TC100"]["bar_dia"]*lapConst
top_cover =  40/304.8
bot_cover =  50/304.8
Yoffset  =  -top_cover - barDict["TC100"]["bar_dia"]/2 

A = 13500/304.8  
x1 = startRadius - (startRadius*math.cos(A/(2*startRadius)))
while A > 13000/304.8 or x1 > 2500/304.8:
    A = A - 100/304.8
    x1 = startRadius - (startRadius*math.cos(A/(2*startRadius)))
debug_print(" start A: " + str(A*304.8/1000))
debug_print("start x1: " + str(round(x1*304.8)/1000))
debug_print("startRad:  " + str(startRadius*304.8))
r1 = startRadius 
r3 = ((startRadius + math.sqrt(startRadius**2 + 4*(barDict["TC100"]["spacing"]/(2*math.pi))*A))/2)-barDict["TC100"]["bar_dia"]
r2 = (r1+r3)/2
r3spl = (startRadius + math.sqrt(startRadius**2 + 4*(barDict["TC100"]["spacing"]/(2*math.pi))*(A-lap_length)))/2
theta1 = 0
theta3 = A/r2  + theta1
theta2 = (theta3-theta1)/2 + theta1
theta3spl = (A-lap_length)/r2

p1 = XYZ(polar_to_car(r1,theta1)[0],polar_to_car(r1,theta1)[1],Yoffset+radius_Yoff(r1))
p2 = XYZ(polar_to_car(r2,theta2)[0],polar_to_car(r2,theta2)[1],Yoffset+radius_Yoff(r2))
p3 = XYZ(polar_to_car(r3,theta3)[0],polar_to_car(r3,theta3)[1],Yoffset+radius_Yoff(r3))
preCurve = Arc.Create(p1,p3,p2)
bar_size = barDict["TC100"]["bar_size"]
bar_type = get_rebar_type(bar_size)
rebarCur = Structure.Rebar.CreateFromCurvesAndShape(doc, sc_65, bar_type, None, None, WTF, XYZ.BasisZ, [preCurve], RebarHookOrientation.Left, RebarHookOrientation.Left)
            
rebarCur.LookupParameter(PARAM_MARK).Set("TOP CONCENTRIC")
rebarCur.LookupParameter(PARAM_SCHEDULE_MARK).Set("TC100")


for bar_params in barDict.values():
    bar_mark    = bar_params["bar_mark"]
    startRadius = bar_params["startRadius"]
    endRadius   = bar_params["endRadius"]
    bar_size    = bar_params["bar_size"]
    bar_dia     = bar_params["bar_dia"]
    spacing     = bar_params["spacing"]

    if endRadius > r_base:
        endRadius = r_base-bot_cover-barDia*2.5
        debug_print("Adjusted endRadius: " + str(endRadius*304.8))
        
    bar_type = get_rebar_type(bar_size)
    barDia = bar_type.LookupParameter("Bar Diameter").AsDouble()

    #### Calculate from input parameters ############################################
    lap_length = lapConst*bar_dia
    Yoffset  =  -top_cover - barDia/2 
    # print(str('radius') + "  ---  \t " + bar_mark + "  --- \t \t " + str('##') + "  ---\t \t " + bar_size + "  ---\t \t " + str('##') + "  ---\t \t " + str('spacing') + "  ---\t \t " + str(spacing))
    debug_print("*"*45)  

    debug_print("bar_mark: " + str(bar_mark))
    debug_print("startRadius: " + str(startRadius*304.8))
    debug_print("endRadius: " + str(endRadius*304.8))
    debug_print("bar_size: " + str(bar_size))
    debug_print("bar_dia: " + str(bar_dia*304.8))
    debug_print("spacing: " + str(spacing*304.8))
    debug_print("lap_length: " + str(lap_length*304.8))
    debug_print("*"*35)
    debug_print("A: " + str(A*304.8/1000))

    if startRadius > 6000/304.8 and barDia < 25/304.8:
        A = 13000/304.8

    tot_subset = 1
    if endRadius-startRadius > spacing*5:
        tot_subset = math.floor((endRadius-startRadius)/(spacing*5))
        debug_print("subset: " + str(tot_subset))
    subset = 1
    subsetRange = (endRadius - startRadius)/tot_subset
    debug_print("#"*45)
    debug_print(A*304.8/1000)
    if A < 13000/304.8:
        while r2 < (endRadius-bot_cover-(barDia*1.5)):
            while r2 < startRadius+subsetRange*subset:
                r1 = r3spl-(bar_dia/2)
                r3 = (r1 + math.sqrt(r1**2 + 4*(spacing/(2*math.pi))*(A-lap_length)))/2
                r2 = (r1+r3)/2
                r3spl = ((r1 + math.sqrt(r1**2 + 4*(spacing/(2*math.pi))*(A-lap_length)))+bar_dia)/2
                theta1 = theta3spl
                theta3 = theta1 + A/r2
                theta2 = (theta3-theta1)/2 +theta1 #+ A/(r2*2)
                theta3spl = (A-lap_length)/r2 + theta1
                #build curve list
                p1 = XYZ(polar_to_car(r1,theta1)[0],polar_to_car(r1,theta1)[1],Yoffset+radius_Yoff(r1))
                p2 = XYZ(polar_to_car(r2,theta2)[0],polar_to_car(r2,theta2)[1],Yoffset+radius_Yoff(r2))
                p3 = XYZ(polar_to_car(r3,theta3)[0],polar_to_car(r3,theta3)[1],Yoffset+radius_Yoff(r3))
                preCurve = Arc.Create(p1,p3,p2)
                #build bar
                rebarCur = Structure.Rebar.CreateFromCurvesAndShape(doc, sc_65, bar_type, None, None, WTF, XYZ.BasisZ, [preCurve], RebarHookOrientation.Left, RebarHookOrientation.Left)
                # set mark and schedule mark
                rebarCur.LookupParameter(PARAM_MARK).Set("TOP CONCENTRIC")
                rebarCur.LookupParameter(PARAM_SCHEDULE_MARK).Set("TC" + str(int(str(bar_mark)[2:]) + subset - 1))
            
            subset += 1
            A_live = (theta3 - theta1)*r2
            x1_live = r2 - (r2*math.cos(A/(2*r2)))
            while A_live < 13000/304.8 and x1_live < 2500/304.8:
                A = A + 100/304.8
                x1_live = r2 - (r2*math.cos(A/(2*r2)))
                A_live = A
            if r2 > 6000/304.8 and barDia < 32/304.8 or A > 12500/304.8:
                A = 13000/304.8
                x1_live = x1
            # print(" start A: " + str(A*304.8/1000))
            # print("start x1: " + str(round(x1*304.8)/1000))
            # print("subset: " + str(subset))

    elif A > 12999/304.8 and A < 13001/304.8:
        debug_print(" in loop " + str(bar_mark))
        while r2 < (endRadius-bot_cover-(barDia*1.5)):
            r1 = r3spl-(bar_dia/2)
            r3 = (r1 + math.sqrt(r1**2 + 4*(spacing/(2*math.pi))*(A-lap_length)))/2
            r2 = (r1+r3)/2
            r3spl = ((r1 + math.sqrt(r1**2 + 4*(spacing/(2*math.pi))*(A-lap_length)))+bar_dia)/2
            theta1 = theta3spl
            theta3 = theta1 + A/r2
            theta2 = (theta3-theta1)/2 +theta1 #+ A/(r2*2)
            theta3spl = (A-lap_length)/r2 + theta1
            #build curve list
            p1 = XYZ(polar_to_car(r1,theta1)[0],polar_to_car(r1,theta1)[1],Yoffset+radius_Yoff(r1))
            p2 = XYZ(polar_to_car(r2,theta2)[0],polar_to_car(r2,theta2)[1],Yoffset+radius_Yoff(r2))
            p3 = XYZ(polar_to_car(r3,theta3)[0],polar_to_car(r3,theta3)[1],Yoffset+radius_Yoff(r3))
            preCurve = Arc.Create(p1,p3,p2)

            #build bar
            rebarCur = Structure.Rebar.CreateFromCurvesAndShape(doc, sc_65, bar_type, None, None, WTF, XYZ.BasisZ, [preCurve], RebarHookOrientation.Left, RebarHookOrientation.Left)
            # set mark and schedule mark
            rebarCur.LookupParameter(PARAM_MARK).Set("TOP CONCENTRIC")
            rebarCur.LookupParameter(PARAM_SCHEDULE_MARK).Set(bar_mark)
            
        
        subset += 1
        A_live = (theta3 - theta1)*r2
        x1_live = r2 - (r2*math.cos(A/(2*r2)))
        while A < 13000/304.8 and x1 < 2500/304.8:
            A = A + 100/304.8
            x1 = r2 - (r2*math.cos(A/(2*r2)))
        debug_print(" start A: " + str(A*304.8/1000))
        debug_print("start x1: " + str(round(x1*304.8)/1000))
        debug_print("subset: " + str(subset))

outConDia    = max((str(bm)[2:] for bm in barDict), key=int)
barDia_out   = barDict["TC" + outConDia]["bar_dia"]
outConRadius = r_base - bot_cover - barDia_out * 0.5
debug_print('#'*100)
debug_print("outConDia: "    + str(outConDia))
debug_print("outConRadius: " + str(outConRadius*304.8))
debug_print('#'*100)
conYoffset = h_base - bot_cover - barDia_out * 0.5
concentric_bar("TC" + outConDia, conYoffset, outConRadius,
               barDict["TC" + outConDia]["bar_size"],
               barDict["TC" + outConDia]["spacing"],
               lap_length=45 * barDia_out)



t.Commit()




FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
scheduleMarkList = []
for elem in FEC:
    if elem.LookupParameter(PARAM_SCHEDULE_MARK).AsString() not in scheduleMarkList:
        scheduleMarkList.append(elem.LookupParameter(PARAM_SCHEDULE_MARK).AsString())



t = Transaction(doc, "Update r")
t.Start()
for sm in scheduleMarkList:
    if "TC" in sm:
        tc_elems = [e for e in FEC if e.LookupParameter(PARAM_SCHEDULE_MARK).AsString() == sm]
        sum_r  = sum(e.LookupParameter("r").AsDouble() for e in tc_elems)
        r = round((sum_r / len(tc_elems)) * 304.8) / 304.8
        for elem in tc_elems:
            elem.LookupParameter("Rebar r Custom").Set(r)


t.Commit()

t = Transaction(doc, "Update A")
t.Start()
for sm in scheduleMarkList:
    if "TC" in sm:
        tc_elems = [e for e in FEC if e.LookupParameter(PARAM_SCHEDULE_MARK).AsString() == sm]
        A_vals   = [e.LookupParameter("A").AsDouble() for e in tc_elems]
        A_max    = max(A_vals)
        A_round  = (round((A_max * 304.8) / 100) * 100) / 304.8
        for elem in tc_elems:
            elem.LookupParameter("A").Set(A_round)


t.Commit()









