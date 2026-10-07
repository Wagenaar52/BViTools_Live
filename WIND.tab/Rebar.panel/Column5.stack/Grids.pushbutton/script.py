from Autodesk.Revit.DB.Structure import * 
from Autodesk.Revit.DB import CurveByPoints, ReferencePointArray, ReferencePoint, CurveArray, PolyLine, Plane, SketchPlane, ElementTransformUtils, Arc
from System.Collections.Generic import List
from Autodesk.Revit.DB import Curve, Line, XYZ, ElementId
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

DEBUG = True
PARAM_MARK = "Mark"
PARAM_SCHEDULE_MARK = "Schedule Mark"

def debug_print(*args):
    if DEBUG:
        import sys
        sys.stdout.write(" ".join([str(a) for a in args]) + "\n")

def get_str(elem, param_name):
    """Parameter text as a string. Returns "" if the parameter is missing or unset."""
    p = elem.LookupParameter(param_name)
    if p is None:
        return ""
    return p.AsString() or ""

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
    A -> bar mark        e.g. GR100
    B -> empty 
    C -> end radius      e.g. 500
    D -> bar size        e.g. Y25
    E -> empty
    F -> height          e.g. 1500
    G -> spacing         e.g. 150

Bar mark: must start with "GR" to be picked up by the script, e.g. GR100

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

# FPath =  forms.pick_file(file_ext='xlsx', multi_file=False, unc_paths=False)
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

#### Element Host ####################################################################
            
WTF = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
for element in WTF:
    if element.Name == "1PA_WTF_SteelTower":
        WTF = element
        break


type_id = WTF.GetTypeId()
type = doc.GetElement(type_id)
r_base = WTF.LookupParameter("rBase").AsDouble()
h_base = WTF.LookupParameter("hBase").AsDouble()
r_plinth = WTF.LookupParameter("rPlinth").AsDouble()
h_cone = WTF.LookupParameter("hCone").AsDouble()
h_plinth = WTF.LookupParameter("hPlinth").AsDouble()
slabSlope = (h_cone - h_base)/(r_base - r_plinth)
rPitOuter = WTF.LookupParameter("rVoidOuter").AsDouble()
rPitInner = WTF.LookupParameter("rVoidInner").AsDouble()
hPit = WTF.LookupParameter("hBottomVoid").AsDouble()
locPoint = WTF.Location.Point
dGroutMid = WTF.LookupParameter("dGroutMiddle").AsDouble()

####  Get number of Anchor Cage Parameters #################################################################
AnchorBolts = func.getAnchorBolts(xl)

# #####Rebar Shape ####################################################################
rebar_shape = FilteredElementCollector(doc).OfClass(RebarShape).WhereElementIsElementType().ToElements()  
for r_shape in rebar_shape:
    if r_shape.LookupParameter("Type Name").AsString() == '20':
        sc_20 = r_shape
        break
for r_shape in rebar_shape:
    if r_shape.LookupParameter("Type Name").AsString() == '38':
        sc_38 = r_shape
        break
    
## define function to determine length of rebar ####################################################################

def bar_length(radius, offset):
    return 1000/304.8 if (2*math.sqrt(radius**2 - offset**2)<1000/304.8) else 2*math.sqrt(radius**2 - offset**2)


def range_length(radius, spacing):
    l = math.sqrt(radius**2 - (500/304.8)**2)
    l = math.floor(l/spacing)*spacing
    return l


def pack_bars_to_average(doc, rebar_elements):

    by_mark = {}
    for elem in rebar_elements:
        sm = get_str(elem, PARAM_SCHEDULE_MARK)
        if sm not in by_mark:
            by_mark[sm] = []
        by_mark[sm].append(elem)

    for sm, bars in by_mark.items():
        x_pack = []   # bars running along X, distributed at different X offsets
        y_pack = []   # bars running along Y, distributed at different Y offsets
        curve_cache = {}

        for bar in bars:
            curves = bar.GetCenterlineCurves(
                True, True, True,
                Structure.MultiplanarOption.IncludeOnlyPlanarCurves, 0)
            if not curves or len(curves) == 0:
                continue
            c = curves[0]
            curve_cache[bar.Id.IntegerValue] = c
            sp = c.GetEndPoint(0)
            ep = c.GetEndPoint(1)
            if abs(ep.X - sp.X) >= abs(ep.Y - sp.Y):
                x_pack.append(bar)
            else:
                y_pack.append(bar)

        # X-pack: move each bar to the average X of the pack
        if x_pack:
            x_positions = []
            for bar in x_pack:
                c = curve_cache[bar.Id.IntegerValue]
                x_positions.append((c.GetEndPoint(0).X + c.GetEndPoint(1).X) / 2.0)
            avg_x = sum(x_positions) / len(x_positions)
            for bar, bar_x in zip(x_pack, x_positions):
                delta = avg_x - bar_x
                if abs(delta) > 1e-9:
                    ElementTransformUtils.MoveElement(doc, bar.Id, XYZ(delta, 0, 0))

        # Y-pack: move each bar to the average Y of the pack
        if y_pack:
            y_positions = []
            for bar in y_pack:
                c = curve_cache[bar.Id.IntegerValue]
                y_positions.append((c.GetEndPoint(0).Y + c.GetEndPoint(1).Y) / 2.0)
            avg_y = sum(y_positions) / len(y_positions)
            for bar, bar_y in zip(y_pack, y_positions):
                delta = avg_y - bar_y
                if abs(delta) > 1e-9:
                    ElementTransformUtils.MoveElement(doc, bar.Id, XYZ(0, delta, 0))


#### Transaction ################################################################################
t = Transaction(doc, 'Reinforce')
t.Start()

for i in range(5,200):
    if "GR" in str(xl.Cells(i, 1).Value2):
        bar_mark = str(xl.Cells(i,1).Value2)
        endRad = float(xl.Cells(i,3).Value2)/304.8
        size = "Y" + str(xl.Cells(i,4).Value2)[1:]
        height = float(xl.Cells(i,6).Value2)/304.8
        Spacing = int(xl.Cells(i,7).Value2)/304.8

        debug_print("#"*50)
        debug_print("  --   " + "BAR MARK"+ "  ----   " + "END RAD" + "  ---  " + "       SIZE"+ "  ---  " + "       HEIGHT" + "  ---  " + "            SPACING" + "  ---  " )
        debug_print("  --   " + bar_mark + " \t ---- \t\t\t  " + str(endRad*304.7) + "  \t\t---  " + size + "  --- \t\t\t " + str(height*304.8) + "  --- \t\t\t " + str(Spacing*304.8) )
        debug_print("*"*10)
        

        ##### Rebar type ####################################################################
            
        all_rebar_types = FilteredElementCollector(doc) \
            .OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsElementType() \
            .ToElements()

        for  rebar_type in all_rebar_types:
            rebar_name = rebar_type.get_Parameter(BuiltInParameter \
                .SYMBOL_NAME_PARAM).AsString()
            if rebar_name == size:
                bar_type = rebar_type
                break

        barDia = bar_type.LookupParameter("Bar Diameter").AsDouble()
        alphaList = ['a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z']
        #### Rebar Shape Properties ####################################################################


        count = 0   
        MarkCount1 = 0
        MarkCount2 = 0
        rebarList = []
        for count in range(0, int((range_length(endRad,Spacing)/Spacing))+1):
            bar_y = bar_length(endRad,abs(-range_length(endRad,Spacing) + count*Spacing))/2
            MarkCount2 += 1
            if count == 0:
                previous_bar_y = bar_y
            if count*Spacing <= range_length(endRad,Spacing):
                if bar_y < (previous_bar_y + Spacing):
                    bar_y = previous_bar_y
                    MarkCount2 = MarkCount1

            rebar_p1 = locPoint+XYZ(-range_length(endRad,Spacing)+ count*Spacing        , bar_y , height+barDia/2)
            rebar_p2 = locPoint+XYZ(-range_length(endRad,Spacing)+ count*Spacing        ,-bar_y , height+barDia/2)

            curve1 = Line.CreateBound(rebar_p1, rebar_p2)

            rebar20 = Structure.Rebar.CreateFromCurves(doc, 
                                                RebarStyle.Standard, 
                                                bar_type, 
                                                None, 
                                                None, 
                                                WTF, 
                                                XYZ.BasisZ, 
                                                [curve1], 
                                                RebarHookOrientation.Left, 
                                                RebarHookOrientation.Left,1,0)
            MarkCount1 = MarkCount2
            rebar20.LookupParameter(PARAM_MARK).Set("GRIDS")
            rebar20.LookupParameter(PARAM_SCHEDULE_MARK).Set(bar_mark+str(alphaList[MarkCount2]))
            count += 1
            previous_bar_y = bar_y
            #print("MarkCount: " + str(MarkCount2))
            rebarList.append(rebar20)
            # if count*Spacing == range_length(endRad,Spacing):
            #         rebarList.remove(rebar20)
            #         print("rebarList: " + str(len(rebarList)))


        rebarListId =[]
        for rebar in rebarList[:-1]:
            rebarListId.append(rebar.Id)
        element_id = List[ElementId](rebarListId)
        
        copiedBars = ElementTransformUtils.CopyElements(doc, element_id, XYZ(0,0,0))#,None,None)
        ElementTransformUtils.RotateElements(doc, element_id, Line.CreateBound(XYZ(0,0,0),XYZ(0,0,1)), math.pi)

        copiedBars = [doc.GetElement(id) for id in copiedBars]

        for rebar in copiedBars:
            rebar.LookupParameter(PARAM_MARK).Set("GRIDS")
            #print("rebar set")

    ############################### 2nd layer of rebar ########################################


        count = 0   
        MarkCount1 = 0
        MarkCount2 = 0
        rebarList = []
        for count in range(0, int((range_length(endRad,Spacing)/Spacing))+1):
            bar_y = bar_length(endRad,abs(-range_length(endRad,Spacing) + count*Spacing))/2
            MarkCount2 += 1
            if count == 0:
                previous_bar_y = bar_y
            if count*Spacing <= range_length(endRad,Spacing):
                if bar_y < (previous_bar_y + Spacing):
                    bar_y = previous_bar_y
                    MarkCount2 = MarkCount1

            
            rebar_p1 = locPoint+XYZ(-bar_y , -range_length(endRad,Spacing)+ count*Spacing        , height-barDia/2)
            rebar_p2 = locPoint+XYZ(bar_y  , -range_length(endRad,Spacing)+ count*Spacing        , height-barDia/2)
            curve1 = Line.CreateBound(rebar_p1, rebar_p2)

            rebar20 = Structure.Rebar.CreateFromCurves(doc, 
                                                RebarStyle.Standard, 
                                                bar_type, 
                                                None, 
                                                None, 
                                                WTF, 
                                                XYZ.BasisZ, 
                                                [curve1], 
                                                RebarHookOrientation.Left, 
                                                RebarHookOrientation.Left,1,0)
            MarkCount1 = MarkCount2
            rebar20.LookupParameter(PARAM_MARK).Set("GRIDS")
            rebar20.LookupParameter(PARAM_SCHEDULE_MARK).Set(bar_mark+str(alphaList[MarkCount2]))
            count += 1
            previous_bar_y = bar_y
            #print("MarkCount: " + str(MarkCount2))
            rebarList.append(rebar20)
            # if count*Spacing == range_length(endRad,Spacing):
            #         rebarList.remove(rebar20)
            #         print("rebarList: " + str(len(rebarList)))
            #         print(rebarList)

        rebarListId =[]
        for rebar in rebarList[:-1]:
            rebarListId.append(rebar.Id)
        element_id = List[ElementId](rebarListId)
        
        copiedBars = ElementTransformUtils.CopyElements(doc, element_id, XYZ(0,0,0))#,None,None)
        ElementTransformUtils.RotateElements(doc, element_id, Line.CreateBound(XYZ(0,0,0),XYZ(0,0,1)), math.pi)

        copiedBars = [doc.GetElement(id) for id in copiedBars]

        for rebar in copiedBars:
            rebar.LookupParameter(PARAM_MARK).Set("GRIDS")
            #print("rebar set")



        debug_print('#'*50)
        debug_print("rebar ran")
        debug_print('#'*50)

    # except Exception as e:
    #         print("Error at row: " + str(i) + " #### " + str(e))
            
 #       delete construction bar
        #doc.Delete(rebar.Id)  

## Supress warnings ################################################################
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SuppressWarnings())
t.SetFailureHandlingOptions(failHandler)
t.Commit()

#close excel
workbook.Close(False)
excel.Quit()

#### Update A ####################################################################

FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
scheduleMarkList = []
for elem in FEC:
    sched_mark = get_str(elem, PARAM_SCHEDULE_MARK)
    if sched_mark not in scheduleMarkList and "GR" in sched_mark and "GRIDS" in get_str(elem, PARAM_MARK):
        scheduleMarkList.append(sched_mark)

#print(scheduleMarkList)

#### Pack bars to average position ####################################################################

t = Transaction(doc, "Update A")    
t.Start()
for i in range(len(scheduleMarkList)):
    if "GR" in scheduleMarkList[i]:
        sum_A = 0
        count_A = 0
        A_max = 0
        for elem in FEC:
            if elem.LookupParameter(PARAM_SCHEDULE_MARK).AsString() == scheduleMarkList[i]:
                sum_A += elem.LookupParameter("A").AsDouble()
                count_A += 1
                if elem.LookupParameter("A").AsDouble() > A_max:
                    A_max = elem.LookupParameter("A").AsDouble()
        A = round((sum_A/count_A)*10)/10
        A = round(A*304.8)/304.8
    #     print(scheduleMarkList[i])
    #     print(A*304.8)
    #     print(A_max*304.8)
    #     print(count_A)
    # print(A*304.8)
    # print(A_max*304.8)

    for elem in FEC:
        if elem.LookupParameter(PARAM_SCHEDULE_MARK).AsString() == scheduleMarkList[i]:
            suffix = elem.LookupParameter(PARAM_SCHEDULE_MARK).AsString()[-1:]
            elem.LookupParameter("A").Set(A_max)#((round((A_max*304.8)/10)*10)/304.8))
            elem.LookupParameter("Rebar r Custom").Set(0)#((round((A_max*304.8)/10)*10)/304.8))
            elem.LookupParameter("Comments").Set(suffix)
            elem.LookupParameter("Rebar Spacing").Set(Spacing)


doc.Regenerate()
t.Commit()

t_pack = Transaction(doc, "Pack bars to average")
t_pack.Start()
doc.Regenerate()

gr_bars = [e for e in FEC if "GRIDS" in get_str(e, PARAM_MARK)]
pack_bars_to_average(doc, gr_bars)

doc.Regenerate()
t_pack.Commit()







