from Autodesk.Revit.DB.Structure import * 
from Autodesk.Revit.DB import CurveByPoints, ReferencePointArray, ReferencePoint, CurveArray, PolyLine, Plane, SketchPlane, ElementTransformUtils
from System.Collections.Generic import List
from Autodesk.Revit.DB import Curve, Line, XYZ
from Autodesk.Revit.DB.Structure import RebarShape
import math, clr
from Autodesk.Revit.DB import Transaction, Structure, FilteredElementCollector, RadialArray, ArrayAnchorMember
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, Line, XYZ, FailureSeverity, FailureProcessingResult,IFailuresPreprocessor
from pyrevit import forms

clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
view = doc.ActiveView

class SupressWarnings(IFailuresPreprocessor):
    
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
            print(traceback.format_exc())
        
        return FailureProcessingResult.Continue


#### Input from excel sheet ####################################################################

FPath =  forms.pick_file(file_ext='xlsx', multi_file=False, unc_paths=False)

excel = Excel.ApplicationClass()
excel.Visible = False
workbook = excel.Workbooks.Open(FPath)
xl = workbook.Worksheets['A']

##### Element Host ####################################################################
            
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


#####Rebar Shape ####################################################################

rebar_shape = FilteredElementCollector(doc).OfClass(RebarShape).WhereElementIsElementType().ToElements()   

                
for r_shape in rebar_shape:
    if str(r_shape.ShapeFamilyId) == '5344399':
        sc_62 = r_shape
        break

for r_shape in rebar_shape:
    if str(r_shape.ShapeFamilyId) == '5344662':
        sc_99g = r_shape
        break

for r_shape in rebar_shape:
    if str(r_shape.ShapeFamilyId) == '5323829':
        sc_20 = r_shape
        break

####  Get number of Anchor Bolts#################################################################
AC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
for Anchorcage in AC:
    if Anchorcage.Name == "1PA_AnchorCage_Assembly 2":
        AnchorCage = Anchorcage
        break
if AnchorCage == None:
    forms.alert("Anchor Cage not found")

AnchorBolts = AnchorCage.LookupParameter("nBolts").AsInteger()

if not isinstance(AnchorBolts, int):
    AnchorBolts = int(xl.Cells(3, 3).Value2)

TRMaxDia = 0
for i in range(5,200):
    if "TR" in str(xl.Cells(i, 1).Value2):
        if TRMaxDia < int(str(xl.Cells(i,4).Value2)[1:]):
            TRMaxDia = int(str(xl.Cells(i,4).Value2)[1:])

for i in range(5,200):
    if "SF" in str(xl.Cells(i, 1).Value2):
        slabFaceConDia = int(str(xl.Cells(i,4).Value2)[1:])/304.8


#### Transaction ################################################################################
t = Transaction(doc, 'Reinforce')
t.Start()

for i in range(5,200):
    if "TR" in str(xl.Cells(i, 1).Value2):
        bar_mark = str(xl.Cells(i,1).Value2)
        no_bars_factor = float(xl.Cells(i,7).Value2)
        size = "Y" + str(xl.Cells(i,4).Value2)[1:]
        RotSwitch = float(xl.Cells(i,5).Value2)
        StartRad = int(xl.Cells(i,2).Value2)/304.8
        EndRad = int(xl.Cells(i,3).Value2)/304.8
        StartHookLength = float(xl.Cells(i,11).Value2)/304.8
        Level = int(xl.Cells(i,10).Value2)
        i += 1


        print("#"*50)
        print("  --   " + "BAR MARK"+ "  ----   " + "NO BARS FACTOR" + "  ---  " + "SIZE"+ "  ---  " + "ROTATION SWITCH" + "  ---  " + "START RAD" + "  ---  " + "END RAD" + "  ---  " + "START HOOK" + "  ---  " + "LEVEL")
        print("  --   " + bar_mark + " \t ---- \t\t\t  " + str(no_bars_factor) + "  \t\t---  " + size + "  --- \t\t\t " + str(RotSwitch) + " \t\t\t --- \t " + str(StartRad*304.8) + "  ---  " + str(EndRad*304.8) + "  ---  " + str(StartHookLength*304.8) +" \t\t --- \t\t " + str(Level))
        print("*"*10)
        
        no_bars = no_bars_factor*AnchorBolts


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

        

        #cover
        top_cover = (40+32)/304.8
        bot_cover = 50/304.8

        if EndRad > (r_base - bot_cover - slabFaceConDia):
            EndRad = r_base - bot_cover - slabFaceConDia  
            print("End Radius adjusted to: " + str(EndRad*304.8))
            print("Check End Radius")

        #### Rebar Shape Properties ####################################################################

        # place points
        level2Offset = (32 + TRMaxDia/2)/304.8 +barDia/2

        slabSlope = (h_cone - h_base)/(r_base - r_plinth)
        theta = math.atan(h_cone - h_base)/(r_base - r_plinth)
        ydelta = top_cover/math.cos(theta)        
        if Level == 1:
            rebar_p1 = locPoint+XYZ(EndRad       ,0 , h_base + slabSlope*((r_base-r_plinth)-(EndRad - r_plinth))-barDia/2 - ydelta)
            rebar_p2 = locPoint+XYZ(StartRad     ,0 , h_base + slabSlope*((r_base-r_plinth)-(StartRad - r_plinth))-barDia/2 - ydelta)
            rebar_p3 = locPoint+XYZ(r_plinth     ,0 , h_cone -barDia/2 - ydelta)
            rebar_p4 = locPoint+XYZ(StartRad     ,0 , h_cone -barDia/2 - ydelta)
            rebar_p5 = locPoint+XYZ(StartRad     ,0 , h_cone -barDia/2 - ydelta-StartHookLength)
        elif Level == 2:
            rebar_p1 = locPoint+XYZ(EndRad       ,0 , h_base + slabSlope*((r_base-r_plinth)-(EndRad - r_plinth))-barDia/2 - ydelta - level2Offset)
            rebar_p2 = locPoint+XYZ(StartRad     ,0 , h_base + slabSlope*((r_base-r_plinth)-(StartRad - r_plinth))-barDia/2 - ydelta- level2Offset)
            rebar_p3 = locPoint+XYZ(r_plinth     ,0 , h_cone -barDia/2 - ydelta- level2Offset)
            rebar_p4 = locPoint+XYZ(StartRad     ,0 , h_cone -barDia/2 - ydelta- level2Offset)
            rebar_p5 = locPoint+XYZ(StartRad     ,0 , h_cone -barDia/2 - ydelta-StartHookLength- level2Offset)
        else:
            print("Level not found: check excel sheet")

        #place curves
        curve1 = Line.CreateBound(rebar_p1, rebar_p2)
        curve2 = Line.CreateBound(rebar_p1, rebar_p3)
        curve3 = Line.CreateBound(rebar_p3, rebar_p4)
   
        if StartHookLength > 100/304.8:
            curve4 = Line.CreateBound(rebar_p4, rebar_p5)
            curve_list99g = List[Curve]([ curve2, curve3, curve4 ])

        #### Cast the list to IList<Curve>
        curve_list20 = List[Curve]([curve1])
        curve_list62 = List[Curve]([curve2, curve3])

        #### Bluid ####################################################################
        if StartRad < r_plinth :        
            if StartHookLength < 100/304.8:
                rebar = Structure.Rebar.CreateFromCurvesAndShape(doc, 
                                                        sc_62, #RebarStyle.Standard,
                                                        bar_type, 
                                                        None,
                                                        None, 
                                                        WTF, 
                                                        XYZ.BasisY, 
                                                        curve_list62,
                                                        RebarHookOrientation.Left, 
                                                        RebarHookOrientation.Left)#,1,0)
            else:
                rebar = Structure.Rebar.CreateFromCurves(doc, 
                                                    RebarStyle.Standard, 
                                                    bar_type, 
                                                    None, 
                                                    None, 
                                                    WTF, 
                                                    XYZ.BasisY, 
                                                    curve_list99g, 
                                                    RebarHookOrientation.Left, 
                                                    RebarHookOrientation.Left,1,0)  
                                                                                    #         RebarStyle style,
                                                                                    #         RebarBarType rebarType,
                                                                                    #         RebarHookType startHook,
                                                                                    #         RebarHookType endHook,
                                                                                    #         Element host,
                                                                                    #         XYZ norm,
                                                                                    #         IList<Curve> curves,
                                                                                    #         RebarHookOrientation startHookOrient,
                                                                                    #         RebarHookOrientation endHookOrient,
                                                                                    #         bool useExistingShapeIfPossible,
                                                                                    #         bool createNewShape

        else:
            rebar = Structure.Rebar.CreateFromCurvesAndShape(doc, 
                                                        sc_20, #RebarStyle.Standard,
                                                        bar_type, 
                                                        None,
                                                        None, 
                                                        WTF, 
                                                        XYZ.BasisY, 
                                                        curve_list20,
                                                        RebarHookOrientation.Left, 
                                                        RebarHookOrientation.Left)#,1,0)

    #set construction link properties
        # rebar.LookupParameter("A").Set(A)
        # rebar.LookupParameter("B").Set(B)
        # rebar.LookupParameter("C").Set(C)
        # rebar.LookupParameter("D").Set(D)
        
        
        #build radial array
        RotAngle = 360*math.pi/180

        elem = RadialArray.ArrayElementWithoutAssociation(doc,
                                                           view, 
                                                           rebar.Id,
                                                           no_bars, 
                                                           Line.CreateBound(locPoint,locPoint+XYZ.BasisZ), 
                                                           RotAngle, ArrayAnchorMember.Last)


        
        #Rotate rebar 
        for elm in elem:
            doc.GetElement(elm).Location.Rotate( Line.CreateBound(locPoint, locPoint + XYZ.BasisZ), RotSwitch*RotAngle/(no_bars*2))
            doc.GetElement(elm).LookupParameter("Mark").Set("TOP RADIALS")
            doc.GetElement(elm).LookupParameter("Schedule Mark").Set(bar_mark)

        print('#'*50)
 #       delete construction bar
        #doc.Delete(rebar.Id)  

## Supress warnings ################################################################
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SupressWarnings())
t.SetFailureHandlingOptions(failHandler)
t.Commit()

