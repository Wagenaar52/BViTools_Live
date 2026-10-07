from Autodesk.Revit.DB.Structure import * 
from Autodesk.Revit.DB import Arc
from System.Collections.Generic import List
from Autodesk.Revit.DB import Curve, Line, XYZ, ElementId, ElementTransformUtils
from Autodesk.Revit.DB.Structure import RebarShape
import math, clr
from Autodesk.Revit.DB import Transaction, Structure, FilteredElementCollector, RadialArray, ArrayAnchorMember
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, FailureSeverity, FailureProcessingResult, IFailuresPreprocessor
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
        try:
            import sys
            sys.stdout.write(" ".join([str(a) for a in args]) + "\n")
        except:
            pass

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
    A -> bar mark        e.g. HP100
    B -> startRadius          e.g. 1800
    C -> vertical offset from bot grout   e.g. 100
    D -> size              e.g. Y25
    E -> rotation switch   e.g. 0.5
    F -> no bars factor    e.g. 1.5
    G -> B (distance between legs) e.g. 200
    H -> (not used)
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

##### Element Host ####################################################################
            
WTF = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
for element in WTF:
    if element.Name == "1PA_WTF_SteelTower":
        WTF = element
        break

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

#####Rebar Shape ###############################################################################
rebar_shape = FilteredElementCollector(doc).OfClass(RebarShape).WhereElementIsElementType().ToElements()  
for r_shape in rebar_shape:
    if r_shape.LookupParameter("Type Name").AsString() == '39':
        sc_39 = r_shape
        break

####  Get number of Anchor Bolts#################################################################
AnchorBolts = func.getAnchorBolts(xl)

### Get Plinth Face Conduit Diameter ############################################################
for i in range(5,200):
    if "PF" in str(xl.Cells(i, 1).Value2):
        plinthFaceConDia = int(str(xl.Cells(i,4).Value2)[1:])/304.8

#### Transaction ################################################################################
t = Transaction(doc, 'Reinforce')
t.Start()

for i in range(5,200):
    if "HP" in str(xl.Cells(i, 1).Value2):
        bar_mark = str(xl.Cells(i,1).Value2)
        no_bars_factor = float(xl.Cells(i,7).Value2)
        size = "Y" + str(xl.Cells(i,4).Value2)[1:]
        RotSwitch = float(xl.Cells(i,5).Value2)
        StartRad =  int(float(xl.Cells(i,2).Value2))/304.8
        VerOffset = int(float(xl.Cells(i,3).Value2))/304.8
        B = float(xl.Cells(i,8).Value2)/304.8

        debug_print("#"*50)
        debug_print("  --   " + "BAR MARK"+ "  ----   " + "NO BARS FACTOR" + "  ---  " + "SIZE"+ "  ---  " + "ROTATION SWITCH" + "  ---  " + "START RAD" + "  ---  " + "h Offset_bot grout" + "  ---  " )
        debug_print("  --   " + bar_mark + " \t ---- \t\t\t  " + str(no_bars_factor) + "  \t\t---  " + size + "  --- \t\t\t " + str(RotSwitch) + " \t\t\t --- \t " + str(StartRad*304.8) + "  ---  " + str(VerOffset*304.8) )
        debug_print("*"*10)
        
        no_bars = no_bars_factor*AnchorBolts


        ##### Rebar type ####################################################################
            
        all_rebar_types = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsElementType().ToElements()

        for rebar_type in all_rebar_types:
            rebar_name = rebar_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            if rebar_name == size:
                bar_type = rebar_type
                break

        barDia = bar_type.LookupParameter("Bar Diameter").AsDouble()

        

        #cover
        top_cover = 40/304.8
        bot_cover = 50/304.8

        #### Rebar Shape Properties ####################################################################

        # place points
        rebar_p1 = locPoint+XYZ(StartRad                                                    ,0 , h_plinth - dGroutMid - top_cover - barDia/2 - VerOffset)
        rebar_p2 = locPoint+XYZ(r_plinth - top_cover - plinthFaceConDia  - B/2      ,0 , h_plinth - dGroutMid - top_cover - barDia/2 - VerOffset)
        rebar_p3 = locPoint+XYZ(r_plinth - top_cover - plinthFaceConDia - barDia*0.5         ,0 , h_plinth - dGroutMid - top_cover - VerOffset - B/2)
        rebar_p4 = locPoint+XYZ(r_plinth - top_cover - plinthFaceConDia   - B/2     ,0 , h_plinth - dGroutMid - top_cover + barDia/2 - VerOffset - B)
        rebar_p5 = locPoint+XYZ(StartRad                                                    ,0 , h_plinth - dGroutMid - top_cover + barDia/2 - VerOffset - B)


        #place curves
        curve1 = Line.CreateBound(rebar_p1, rebar_p2)
        curve2 = Arc.Create(rebar_p2, rebar_p4, rebar_p3)
        curve3 = Line.CreateBound(rebar_p4, rebar_p5)
   
        #Cast the list to IList<Curve>
        curve_list39 = List[Curve]([curve1, curve2, curve3])

        #### Bluid ####################################################################

        # rebar = Structure.Rebar.CreateFromCurvesAndShape(doc,                                               
        #                                         sc_39, #RebarStyle.Standard,
        #                                         bar_type, 
        #                                         None,
        #                                         None, 
        #                                         WTF, 
        #                                         XYZ.BasisY, 
        #                                         curve_list39,
        #                                         RebarHookOrientation.Left, 
        #                                         RebarHookOrientation.Left)#,1,0)
        
        rebar = Structure.Rebar.CreateFromCurves(doc, 
                                            RebarStyle.Standard, 
                                            bar_type, 
                                            None, 
                                            None, 
                                            WTF, 
                                            XYZ.BasisY, 
                                            curve_list39, 
                                            RebarHookOrientation.Left, 
                                            RebarHookOrientation.Left,1,0)	#bool useExistingShapeIfPossible,
	                                                                        #bool createNewShape
        
        #### Build Section Bar ####################################################################
        #copy rebar to left and right of the section in the plane Y=0
        section_bar_right_id = ElementTransformUtils.CopyElement(doc, rebar.Id, XYZ(0, 0, 0))[0]
        section_bar_right = doc.GetElement(section_bar_right_id)
        section_bar_right.LookupParameter(PARAM_MARK).Set("SECTION HAIR PINS")
        section_bar_right.LookupParameter(PARAM_SCHEDULE_MARK).Set(bar_mark)
        section_bar_right.LookupParameter("Rebar r Custom").Set(0)
        section_bar_right.LookupParameter("Rebar Spacing").Set((360.0 / no_bars) / 304.8)
        section_bar_right.LookupParameter("Rebar Quantity").Set(no_bars)

        # rotate the second copy 180 degrees around the vertical axis to mirror it on the other side of the section
        section_bar_left_id = ElementTransformUtils.CopyElement(doc, rebar.Id, XYZ(0, 0, 0))[0]
        section_bar_left = doc.GetElement(section_bar_left_id)
        section_bar_left.Location.Rotate(Line.CreateBound(locPoint, locPoint + XYZ.BasisZ), math.pi)
        section_bar_left.LookupParameter(PARAM_MARK).Set("SECTION HAIR PINS")
        section_bar_left.LookupParameter(PARAM_SCHEDULE_MARK).Set(bar_mark)
        section_bar_left.LookupParameter("Rebar r Custom").Set(0)
        section_bar_left.LookupParameter("Rebar Spacing").Set((360.0 / no_bars) / 304.8)
        section_bar_left.LookupParameter("Rebar Quantity").Set(no_bars)


        #build radial array
        RotAngle = 360*math.pi/180

        elem = RadialArray.ArrayElementWithoutAssociation(doc,
                                                           doc.GetElement(ElementId(2964800)),
                                                           rebar.Id,
                                                           no_bars, 
                                                           Line.CreateBound(locPoint,locPoint+XYZ.BasisZ), 
                                                           RotAngle, ArrayAnchorMember.Last)
        rebar.LookupParameter(PARAM_MARK).Set("HAIR PINS")
        rebar.LookupParameter(PARAM_SCHEDULE_MARK).Set(bar_mark)
        rebar.LookupParameter("Rebar Quantity").Set(no_bars)
        rebar.LookupParameter("Rebar r Custom").Set(float(0))
        rebar.LookupParameter("Rebar Spacing").Set((360.0/no_bars)/304.8)


        #Rotate rebar 
        for elm in elem:
            doc.GetElement(elm).Location.Rotate( Line.CreateBound(locPoint, locPoint + XYZ.BasisZ), RotSwitch*RotAngle/(no_bars*2))
            doc.GetElement(elm).LookupParameter(PARAM_MARK).Set("HAIR PINS")
            doc.GetElement(elm).LookupParameter(PARAM_SCHEDULE_MARK).Set(bar_mark)
            doc.GetElement(elm).LookupParameter("Rebar Quantity").Set(no_bars)
            doc.GetElement(elm).LookupParameter("Rebar r Custom").Set(float(0))
            doc.GetElement(elm).LookupParameter("Rebar Spacing").Set((360.0/no_bars)/304.8)

        debug_print('#'*50)
        debug_print("rebar ran")
        debug_print('#'*50)
 #       delete construction bar
        #doc.Delete(rebar.Id)  

## Supress warnings ################################################################
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SuppressWarnings())
t.SetFailureHandlingOptions(failHandler)
t.Commit()

