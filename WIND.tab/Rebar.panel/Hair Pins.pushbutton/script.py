from Autodesk.Revit.DB.Structure import * 
from Autodesk.Revit.DB import CurveByPoints, ReferencePointArray, ReferencePoint, CurveArray, PolyLine, Plane, SketchPlane, ElementTransformUtils, Arc
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
dGroutMid = WTF.LookupParameter("dGroutMiddle").AsDouble()

#####Rebar Shape ####################################################################

rebar_shape = FilteredElementCollector(doc).OfClass(RebarShape).WhereElementIsElementType().ToElements()   

                
for r_shape in rebar_shape:
    if str(r_shape.ShapeFamilyId) == '5871777':
        sc_39 = r_shape
        break


####  Get number of Anchor Bolts#################################################################
AC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
for Anchorcage in AC:
    if Anchorcage.Name == "1PA_AnchorCage_Assembly 2":
        AnchorCage = Anchorcage
        
if AnchorCage == None:
    forms.alert("Anchor Cage not found")

AnchorBolts = AnchorCage.LookupParameter("nBolts").AsInteger()

if not isinstance(AnchorBolts, int):
    AnchorBolts = int(xl.Cells(3, 3).Value2)


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
        StartRad = int(xl.Cells(i,2).Value2)/304.8
        hOffset = int(xl.Cells(i,3).Value2)/304.8
        B = float(xl.Cells(i,8).Value2)/304.8

        i += 1


        print("#"*50)
        print("  --   " + "BAR MARK"+ "  ----   " + "NO BARS FACTOR" + "  ---  " + "SIZE"+ "  ---  " + "ROTATION SWITCH" + "  ---  " + "START RAD" + "  ---  " + "h Offset_bot grout" + "  ---  " )
        print("  --   " + bar_mark + " \t ---- \t\t\t  " + str(no_bars_factor) + "  \t\t---  " + size + "  --- \t\t\t " + str(RotSwitch) + " \t\t\t --- \t " + str(StartRad*304.8) + "  ---  " + str(hOffset*304.8) )
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
        top_cover = 40/304.8
        bot_cover = 50/304.8

        #### Rebar Shape Properties ####################################################################

        # place points

        rebar_p1 = locPoint+XYZ(StartRad                                                    ,0 , h_plinth - dGroutMid - top_cover - barDia/2 - hOffset)
        rebar_p2 = locPoint+XYZ(r_plinth - top_cover - plinthFaceConDia  - B/2      ,0 , h_plinth - dGroutMid - top_cover - barDia/2 - hOffset)
        rebar_p3 = locPoint+XYZ(r_plinth - top_cover - plinthFaceConDia - barDia*0.5         ,0 , h_plinth - dGroutMid - top_cover - hOffset - B/2)
        rebar_p4 = locPoint+XYZ(r_plinth - top_cover - plinthFaceConDia   - B/2     ,0 , h_plinth - dGroutMid - top_cover + barDia/2 - hOffset - B)
        rebar_p5 = locPoint+XYZ(StartRad                                                    ,0 , h_plinth - dGroutMid - top_cover + barDia/2 - hOffset - B)


        #place curves
        curve1 = Line.CreateBound(rebar_p1, rebar_p2)
        curve2 = Arc.Create(rebar_p2, rebar_p4, rebar_p3)
        curve3 = Line.CreateBound(rebar_p4, rebar_p5)

        geomPlane = Plane.CreateByThreePoints(rebar_p1, rebar_p2, rebar_p3)
        sketch = SketchPlane.Create(doc, geomPlane)

        model_line = doc.Create.NewModelCurve(curve1, sketch)
        model_line = doc.Create.NewModelCurve(curve2, sketch)
        model_line = doc.Create.NewModelCurve(curve3, sketch)
   


        #### Cast the list to IList<Curve>
    
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
                                            RebarHookOrientation.Left,1,1)




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
            doc.GetElement(elm).LookupParameter("Mark").Set("HAIR PINS")
            doc.GetElement(elm).LookupParameter("Schedule Mark").Set(bar_mark)

        print('#'*50)
        print("rebar ran")
        print('#'*50)
 #       delete construction bar
        #doc.Delete(rebar.Id)  

## Supress warnings ################################################################
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SupressWarnings())
t.SetFailureHandlingOptions(failHandler)
t.Commit()

