from Autodesk.Revit.DB.Structure import *
from System.Collections.Generic import List
from Autodesk.Revit.DB import Curve, Line, XYZ, Plane, SketchPlane
import math, clr
from Autodesk.Revit.DB import Transaction, Structure, FilteredElementCollector, RadialArray, ArrayAnchorMember
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, FailureSeverity, FailureProcessingResult, IFailuresPreprocessor
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
    A -> bar mark        e.g. ST100
    B -> link radius     e.g. 500
    C -> empty
    D -> bar size        e.g. Y25
    E -> rotation switch e.g. 1 for 360/no_bars*4, 2 for 2*360/no_bars*4 etc.
    F -> empty
    G -> number of bars factor e.g. 0.5 for half the anchor bolts, 1 for same as anchor bolts, 2 for double the anchor bolts etc.

 - Bar mark: must start with ST
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




#### Input from excel sheet ####################################################################

FPath = func.getFilePath() #forms.pick_file(file_ext='xlsx', multi_file=False, unc_paths=False)

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


AnchorBolts =  func.getAnchorBolts(xl) #AnchorCage.LookupParameter("nBolts").AsInteger()

#####Rebar Shape ####################################################################

rebar_shape = FilteredElementCollector(doc).OfClass(RebarShape).WhereElementIsElementType().ToElements()  
for r_shape in rebar_shape:
    if r_shape.LookupParameter("Type Name").AsString() == '99d':
        sc_99d = r_shape
        break


SlabFace_dia = 0

#####ANCHORAGE LENGTHS#############################################################################################################################
anchorage_lengths = {
    12: {20: 47, 25: 42, 30: 38, 35: 36, 40: 33, 45: 31, 50: 30, 60: 27},
    14: {20: 50, 25: 44, 30: 41, 35: 38, 40: 35, 45: 33, 50: 31, 60: 29},
    16: {20: 52, 25: 46, 30: 42, 35: 39, 40: 37, 45: 35, 50: 33, 60: 30},
    20: {20: 56, 25: 50, 30: 46, 35: 42, 40: 40, 45: 37, 50: 35, 60: 32},
    25: {20: 60, 25: 54, 30: 49, 35: 46, 40: 43, 45: 40, 50: 38, 60: 35},
    28: {20: 63, 25: 56, 30: 51, 35: 47, 40: 44, 45: 42, 50: 40, 60: 36},
    32: {20: 65, 25: 58, 30: 53, 35: 49, 40: 46, 45: 44, 50: 41, 60: 38}
}
def get_anchorage_length(bar_diameter, concrete_grade):
    debug_print("bardiameter   ", bar_diameter)
    debug_print("concrete_grade   ", concrete_grade)
    if bar_diameter in anchorage_lengths and concrete_grade in anchorage_lengths[bar_diameter]:
        return int(anchorage_lengths[bar_diameter][concrete_grade])*bar_diameter/304.8
    else:
        debug_print("No appropriate bars found to calculate anchorage length.")
        return None

def set_rebar_params(rebar, bar_mark, no_bars):
    rebar.LookupParameter(PARAM_MARK).Set("STOOLS")
    rebar.LookupParameter(PARAM_SCHEDULE_MARK).Set(bar_mark)
    rebar.LookupParameter("Rebar Quantity").Set(no_bars)
    rebar.LookupParameter("Rebar r Custom").Set(float(0))
    rebar.LookupParameter("Rebar Spacing").Set((360.0/no_bars)/304.8)

####################################################################################################################################################################################

t = Transaction(doc, 'Reinforce')
t.Start()



for i in range(5,200):
    if "ST" in str(xl.Cells(i, 1).Value2):
        bar_mark = str(xl.Cells(i,1).Value2)
        no_bars_factor = float(xl.Cells(i,7).Value2)
        LinkRadius = float(xl.Cells(i,2).Value2)/304.8
        size = "Y" + str(xl.Cells(i,4).Value2)[1:]
        RotSwitch = float(xl.Cells(i,5).Value2)
        

        debug_print("  ----   " + "BAR MARK"+ "  ----   " + "NO BARS FACTOR" + "  ---  " + "SIZE" + "  ---  " + "Link Radius")
        debug_print("  ----   " + bar_mark + "  ----   " + str(no_bars_factor) + "  ---  " + size + "  ---  " + str(LinkRadius*304.8))
        debug_print("*"*30)

        no_bars = no_bars_factor*AnchorBolts

        ##### Rebar type ####################################################################
            
        all_rebar_types = FilteredElementCollector(doc) \
            .OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsElementType() \
            .ToElements()

        for rebar_type in all_rebar_types:
            rebar_name = rebar_type.get_Parameter(BuiltInParameter \
                .SYMBOL_NAME_PARAM).AsString()
            if rebar_name == size:
                bar_type = rebar_type
                break

        barDia = bar_type.LookupParameter("Bar Diameter").AsDouble()

        ##### Element Host ####################################################################
            
        WTF = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
        for element in WTF:
            if element.Name == "1PA_WTF_SteelTower":
                WTF = element
                break

        type_id = WTF.GetTypeId()
        r_base = WTF.LookupParameter("rBase").AsDouble()
        h_base = WTF.LookupParameter("hBase").AsDouble()
        r_plinth = WTF.LookupParameter("rPlinth").AsDouble()
        h_cone = WTF.LookupParameter("hCone").AsDouble()
        h_plinth = WTF.LookupParameter("hPlinth").AsDouble()
        slabSlope = (h_cone - h_base)/(r_base - r_plinth)

        #cover
        top_cover = 40/304.8
        bot_cover = 50/304.8
        radius = r_base - bot_cover - barDia/2

        y_Rangetop = h_base + slabSlope*bot_cover -top_cover -barDia*1.5
        y_Rangebot = bot_cover + barDia
        B = y_Rangetop - y_Rangebot

        locPoint = WTF.Location.Point


        #### Rebar Shape Properties ####################################################################

        #Rebar Properties
        #A
        A = (get_anchorage_length(int(str(size)[1:]), 35))*1.2 +7.5*barDia/304.8
        #B
        y_Rangetop = h_cone - slabSlope*(LinkRadius-r_plinth) -top_cover -barDia*1.5
        y_Rangebot = bot_cover + barDia/2 + 32/304.8
        B = y_Rangetop - y_Rangebot
        #C
        C =  get_anchorage_length(int(str(size)[1:]), 35) +7.5*barDia/304.8


        debug_print("A: " + str(A*304.8))
        debug_print("B: " + str(B*304.8))
        debug_print("C: " + str(C*304.8))


        #### Build ####################################################################

        rebar_p1 = locPoint+XYZ(LinkRadius                                    ,0, B + bot_cover+barDia/2 + 32/304.8 )
        rebar_p2 = locPoint+XYZ(LinkRadius + ((A**2)/(slabSlope+1) )**0.5     ,0, B + bot_cover+barDia/2 + 32/304.8 - slabSlope*((A**2)/(slabSlope+1) )**0.5)
        rebar_p3 = locPoint+XYZ(LinkRadius                                    ,0,bot_cover+barDia/2 + 32/304.8)
        rebar_p4 = locPoint+XYZ(LinkRadius + C                                ,0,bot_cover+barDia/2 + 32/304.8)


        curve1 = Line.CreateBound(rebar_p2, rebar_p1)
        curve2 = Line.CreateBound(rebar_p1, rebar_p3)
        curve3 = Line.CreateBound(rebar_p3, rebar_p4)


        curves = [ curve1, curve2, curve3]
        
        # Cast the list to IList<Curve>
        curve_list = List[Curve](curves)

        rebar = Structure.Rebar.CreateFromCurves(doc, 
                                                RebarStyle.Standard, 
                                                bar_type, 
                                                None, 
                                                None, 
                                                WTF, 
                                                XYZ.BasisY, 
                                                curve_list, 
                                                RebarHookOrientation.Left, 
                                                RebarHookOrientation.Left,1,0)

        #build radial array
        RotAngle = 360*math.pi/180
        if no_bars <= 200:
            elem = RadialArray.ArrayElementWithoutAssociation(doc,
                                                            view, 
                                                            rebar.Id,
                                                            no_bars, 
                                                            Line.CreateBound(locPoint,locPoint+XYZ.BasisZ), 
                                                            RotAngle, ArrayAnchorMember.Last)
        elif no_bars > 200:
            elem = RadialArray.ArrayElementWithoutAssociation(doc,
                                                view, 
                                                rebar.Id,
                                                no_bars/4, 
                                                Line.CreateBound(locPoint,locPoint+XYZ.BasisZ), 
                                                RotAngle, ArrayAnchorMember.Last)
            elem2 = RadialArray.ArrayElementWithoutAssociation(doc,
                                                            view, 
                                                            rebar.Id,
                                                            no_bars/4, 
                                                            Line.CreateBound(locPoint,locPoint+XYZ.BasisZ), 
                                                            RotAngle, ArrayAnchorMember.Last)
            elem3 = RadialArray.ArrayElementWithoutAssociation(doc,
                                                            view, 
                                                            rebar.Id,
                                                            no_bars/4, 
                                                            Line.CreateBound(locPoint,locPoint+XYZ.BasisZ), 
                                                            RotAngle, ArrayAnchorMember.Last)
            elem4 = RadialArray.ArrayElementWithoutAssociation(doc,
                                                            view, 
                                                            rebar.Id,
                                                            no_bars/4, 
                                                            Line.CreateBound(locPoint,locPoint+XYZ.BasisZ), 
                                                            RotAngle, ArrayAnchorMember.Last)
            ###########
            rebar2 = Structure.Rebar.CreateFromCurves(doc, 
                                    RebarStyle.Standard, 
                                    bar_type, 
                                    None, 
                                    None, 
                                    WTF, 
                                    XYZ.BasisY, 
                                    curve_list, 
                                    RebarHookOrientation.Left, 
                                    RebarHookOrientation.Left,1,0)
            set_rebar_params(rebar2, bar_mark, no_bars)

            doc.GetElement(rebar2.Id).Location.Rotate( Line.CreateBound(locPoint, locPoint + XYZ.BasisZ), (RotAngle)/(no_bars))
            ###########
            rebar2 = Structure.Rebar.CreateFromCurves(doc, 
                                    RebarStyle.Standard, 
                                    bar_type, 
                                    None, 
                                    None, 
                                    WTF, 
                                    XYZ.BasisY, 
                                    curve_list, 
                                    RebarHookOrientation.Left, 
                                    RebarHookOrientation.Left,1,0)
            set_rebar_params(rebar2, bar_mark, no_bars)
            doc.GetElement(rebar2.Id).Location.Rotate( Line.CreateBound(locPoint, locPoint + XYZ.BasisZ), (RotAngle)/(no_bars))
            ###########
            rebar3 = Structure.Rebar.CreateFromCurves(doc, 
                                    RebarStyle.Standard, 
                                    bar_type, 
                                    None, 
                                    None, 
                                    WTF, 
                                    XYZ.BasisY, 
                                    curve_list, 
                                    RebarHookOrientation.Left, 
                                    RebarHookOrientation.Left,1,0)
            set_rebar_params(rebar3, bar_mark, no_bars)
            doc.GetElement(rebar3.Id).Location.Rotate( Line.CreateBound(locPoint, locPoint + XYZ.BasisZ), 2*(RotAngle)/(no_bars))
            ###########          
            rebar4 = Structure.Rebar.CreateFromCurves(doc, 
                                    RebarStyle.Standard, 
                                    bar_type, 
                                    None, 
                                    None, 
                                    WTF, 
                                    XYZ.BasisY, 
                                    curve_list, 
                                    RebarHookOrientation.Left, 
                                    RebarHookOrientation.Left,1,0)
            set_rebar_params(rebar4, bar_mark, no_bars)
            doc.GetElement(rebar4.Id).Location.Rotate( Line.CreateBound(locPoint, locPoint + XYZ.BasisZ), 3*(RotAngle)/(no_bars))
            ###########
            #rotare elem2 with half the angle
            for elm in elem2:
                doc.GetElement(elm).Location.Rotate( Line.CreateBound(locPoint, locPoint + XYZ.BasisZ), (RotAngle)/(no_bars))
            set_rebar_params(doc.GetElement(elm), bar_mark, no_bars)

            #rotare elem3 with half the angle
            for elm in elem3:
                doc.GetElement(elm).Location.Rotate( Line.CreateBound(locPoint, locPoint + XYZ.BasisZ), 2*(RotAngle)/(no_bars))
            set_rebar_params(doc.GetElement(elm), bar_mark, no_bars)

            #rotare elem4 with half the angle
            for elm in elem4:
                doc.GetElement(elm).Location.Rotate( Line.CreateBound(locPoint, locPoint + XYZ.BasisZ), 3*(RotAngle)/(no_bars))
                set_rebar_params(doc.GetElement(elm), bar_mark, no_bars)

            for elm in elem:
                set_rebar_params(doc.GetElement(elm), bar_mark, no_bars)
            ########
            doc.Delete(rebar.Id)
            doc.Delete(rebar2.Id)
            rebar = Structure.Rebar.CreateFromCurves(doc, 
                                    RebarStyle.Standard, 
                                    bar_type, 
                                    None, 
                                    None, 
                                    WTF, 
                                    XYZ.BasisY, 
                                    curve_list, 
                                    RebarHookOrientation.Left, 
                                    RebarHookOrientation.Left,1,0)
            set_rebar_params(rebar, bar_mark, no_bars)
        #Rotate rebar 
        if no_bars <= 200:
            for elm in elem:
                doc.GetElement(elm).Location.Rotate( Line.CreateBound(locPoint, locPoint + XYZ.BasisZ), RotSwitch*RotAngle/(no_bars*4))
                set_rebar_params(doc.GetElement(elm), bar_mark, no_bars)
            
        #delete construction bar
        #doc.Delete(rebar.Id)  

## Supress warnings ################################################################
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SuppressWarnings())
t.SetFailureHandlingOptions(failHandler)
t.Commit()








