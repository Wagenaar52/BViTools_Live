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

FPath =  forms.pick_file(file_ext='xlsx', multi_file=False, unc_paths=False) #"C:\Users\Wagner.Human\Desktop\Wolf_RebarData_RevD_V162r5.xlsx" #

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
wGroutTop = WTF.LookupParameter("wGroutTop").AsDouble()
rTower = WTF.LookupParameter("rTower").AsDouble()
hPit = WTF.LookupParameter("hBottomVoid").AsDouble()

#####Rebar Shape ####################################################################

rebar_shape = FilteredElementCollector(doc).OfClass(RebarShape).WhereElementIsElementType().ToElements()   

for r_shape in rebar_shape:
    if r_shape.LookupParameter("Type Name").AsString() == '74':
        sc_74 = r_shape


####  Get number of Anchor Bolts#################################################################
AC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
for Anchorcage in AC:
    if Anchorcage.Name == "1PA_AnchorCage_Assembly":
        AnchorCage = Anchorcage
        break
if AnchorCage == None:
    forms.alert("Anchor Cage not found")

AnchorBolts = AnchorCage.LookupParameter("nBolts").AsInteger()
wbotFlange = AnchorCage.LookupParameter("wFlangeBot").AsDouble()


if not isinstance(AnchorBolts, int):
    AnchorBolts = int(xl.Cells(3, 3).Value2)


for i in range(5,200):
    if "PF1" in str(xl.Cells(i, 1).Value2):
        plinthFaceConDia = int(str(xl.Cells(i,4).Value2)[1:])/304.8
    elif "PH1" in str(xl.Cells(i, 1).Value2):
        plinthHorConDia = int(str(xl.Cells(i,4).Value2)[1:])/304.8
    elif "GR1" in str(xl.Cells(i, 1).Value2):
        botGridDia = int(str(xl.Cells(i,4).Value2)[1:])/304.8

#### Transaction ################################################################################
t = Transaction(doc, 'Reinforce')
t.Start()

for i in range(5,200):
    if "PV20" in str(xl.Cells(i, 1).Value2):
        bar_mark = str(xl.Cells(i,1).Value2)
        no_bars_factor = float(xl.Cells(i,7).Value2)
        size = "Y" + str(xl.Cells(i,4).Value2)[1:]
        RotSwitch = float(xl.Cells(i,5).Value2)
        horOffsetInner = float(xl.Cells(i,3).Value2)/304.8
        i += 1


        print("#"*50)
        print("  --   " + "BAR MARK"+ "  ----   " + "NO BARS FACTOR" + "  ---  " + "SIZE"+ "  ---  " + "ROTATION SWITCH"  )
        print("  --   " + bar_mark + " \t ---- \t\t\t  " + str(no_bars_factor) + "  \t\t---  " + size + "  --- \t\t\t " + str(RotSwitch)   )
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

        A = (math.ceil(barDia*14.5*304.8/50))*50/304.8 #14.5xdia and rounded up to the nearest 50mm

        ACoutRad = max(wGroutTop, wbotFlange)
        r = ((r_plinth- top_cover - plinthFaceConDia -  (barDia/2))-(horOffsetInner+(barDia/2)+rTower+(ACoutRad/2)))/2
        #### Rebar Shape Properties ####################################################################

        # place points
        rebar_p1 = locPoint+XYZ(horOffsetInner+(barDia/2)+rTower+(ACoutRad/2)+A              , 0 ,  -hPit+bot_cover + barDia/2 +(2*botGridDia))
        rebar_p2 = locPoint+XYZ(horOffsetInner+(barDia/2)+rTower+(ACoutRad/2)                , 0 ,  -hPit+bot_cover + barDia/2 +(2*botGridDia))
        rebar_p3 = locPoint+XYZ(horOffsetInner+(barDia/2)+rTower+(ACoutRad/2)                , 0 ,  h_plinth -top_cover - barDia/2 -plinthHorConDia)
        #rebar_p4 = locPoint+XYZ(horOffsetInner+(barDia/2)+rTower+(ACoutRad/2)+r              , 0 ,  h_plinth -top_cover - barDia/2 -plinthHorConDia)
        rebar_p5 = locPoint+XYZ(r_plinth- top_cover - plinthFaceConDia -  (barDia/2)                , 0 ,  h_plinth -top_cover - barDia/2 -plinthHorConDia)
        rebar_p6 = locPoint+XYZ(r_plinth- top_cover - plinthFaceConDia -  (barDia/2)                , 0 ,  -hPit+bot_cover + barDia/2 +(2*botGridDia) )
        rebar_p7 = locPoint+XYZ(horOffsetInner+(barDia/2)+rTower+(ACoutRad/2)                        , 0 ,  -hPit+bot_cover + barDia/2 +(2*botGridDia) +A )
        
        #place curves
        curve1 = Line.CreateBound(rebar_p2, rebar_p3)
        curve2 = Line.CreateBound(rebar_p3, rebar_p5)
        curve3 = Line.CreateBound(rebar_p5, rebar_p6)
        #curve3 = Arc.Create(rebar_p3, rebar_p5, rebar_p4)
        curve4 = Line.CreateBound(rebar_p6, rebar_p2)
        #curve5 = Line.CreateBound(rebar_p6, rebar_p7)

        geomPlane = Plane.CreateByThreePoints(rebar_p5, rebar_p2, rebar_p3)
        sketch = SketchPlane.Create(doc, geomPlane)

        model_line = doc.Create.NewModelCurve(curve1, sketch)
        model_line = doc.Create.NewModelCurve(curve2, sketch)
        model_line = doc.Create.NewModelCurve(curve3, sketch)
        model_line = doc.Create.NewModelCurve(curve4, sketch)
        # model_line = doc.Create.NewModelCurve(curve5, sketch)
   

        #### Cast the list to IList<Curve>
    
        curve_list99a = List[Curve]([curve1, curve2, curve3, curve4])#, curve5])

        #### Bluid ####################################################################


        
        rebar = Structure.Rebar.CreateFromCurves(doc, 
                                            RebarStyle.Standard, 
                                            bar_type, 
                                            None, 
                                            None, 
                                            WTF, 
                                            XYZ.BasisY, 
                                            curve_list99a, 
                                            RebarHookOrientation.Left, 
                                            RebarHookOrientation.Left,0,1)




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
            doc.GetElement(elm).LookupParameter("Mark").Set("PLINTH VERTICAL")
            doc.GetElement(elm).LookupParameter("Schedule Mark").Set(bar_mark)

        print('#'*50)
        print("PV200 ran")
        print('#'*50)
    #    delete construction bar
    #     doc.Delete(rebar.Id)  
    elif "PV10" in str(xl.Cells(i, 1).Value2):
        bar_mark = str(xl.Cells(i,1).Value2)
        no_bars_factor = float(xl.Cells(i,7).Value2)
        size = "Y" + str(xl.Cells(i,4).Value2)[1:]
        RotSwitch = float(xl.Cells(i,5).Value2)
        horOffsetInner = float(xl.Cells(i,3).Value2)/304.8
        startRad = float(xl.Cells(i,2).Value2)/304.8
        i += 1


        print("#"*50)
        print("  --   " + "BAR MARK"+ "  ----   " + "NO BARS FACTOR" + "  ---  " + "SIZE"+ "  ---  " + "ROTATION SWITCH"  )
        print("  --   " + bar_mark + " \t ---- \t\t\t  " + str(no_bars_factor) + "  \t\t---  " + size + "  --- \t\t\t " + str(RotSwitch)   )
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

        A = (math.ceil(barDia*14.5*304.8/50))*50/304.8 #14.5xdia and rounded up to the nearest 50mm

        ACoutRad = rTower - max(wGroutTop, wbotFlange)/2
        r = ((ACoutRad - horOffsetInner -barDia/2) - (startRad))/2
        #### Rebar Shape Properties ####################################################################

        # place points
        #rebar_p1 = locPoint+XYZ(startRad - A              , 0 ,  -hPit+bot_cover + barDia/2 +(2*botGridDia))
        rebar_p2 = locPoint+XYZ(startRad                  , 0 ,  -hPit+bot_cover + barDia/2 +(2*botGridDia))
        rebar_p3 = locPoint+XYZ(startRad                  , 0 ,  h_plinth -top_cover - barDia/2 -plinthHorConDia-r)
        #rebar_p4 = locPoint+XYZ(startRad + r              , 0 ,  h_plinth -top_cover - barDia/2 -plinthHorConDia)
        rebar_p5 = locPoint+XYZ(startRad + 2*r            , 0 ,  h_plinth -top_cover - barDia/2 -plinthHorConDia-r)
        rebar_p6 = locPoint+XYZ(startRad + 2*r            , 0 ,  -hPit+bot_cover + barDia/2 +(2*botGridDia) )
        #rebar_p7 = locPoint+XYZ(startRad + 2*r + A        , 0 ,  -hPit+bot_cover + barDia/2 +(2*botGridDia) )
        
        #place curves
        curve1 = Line.CreateBound(rebar_p3, rebar_p5)
        curve2 = Line.CreateBound(rebar_p5, rebar_p6)
        # curve3 = Arc.Create(rebar_p3, rebar_p5, rebar_p4)
        # curve4 = Line.CreateBound(rebar_p5, rebar_p6)
        curve3 = Line.CreateBound(rebar_p6, rebar_p2)

        geomPlane = Plane.CreateByThreePoints(rebar_p5, rebar_p2, rebar_p3)
        sketch = SketchPlane.Create(doc, geomPlane)

        model_line = doc.Create.NewModelCurve(curve1, sketch)
        model_line = doc.Create.NewModelCurve(curve2, sketch)
        model_line = doc.Create.NewModelCurve(curve3, sketch)
        # model_line = doc.Create.NewModelCurve(curve4, sketch)
        # model_line = doc.Create.NewModelCurve(curve5, sketch)
   

        #### Cast the list to IList<Curve>
    
        curve_list99a = List[Curve]([curve1, curve2, curve3])#, curve4, curve5])

        #### Bluid ####################################################################

        
        rebar = Structure.Rebar.CreateFromCurves(doc, 
                                            RebarStyle.Standard, 
                                            bar_type, 
                                            None, 
                                            None, 
                                            WTF, 
                                            XYZ.BasisY, 
                                            curve_list99a, 
                                            RebarHookOrientation.Left, 
                                            RebarHookOrientation.Left,1,0)




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
            doc.GetElement(elm).LookupParameter("Mark").Set("PLINTH VERTICAL")
            doc.GetElement(elm).LookupParameter("Schedule Mark").Set(bar_mark)

        print('#'*50)
        print("PV100 ran")
        print('#'*50)
## Supress warnings ################################################################
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SupressWarnings())
t.SetFailureHandlingOptions(failHandler)
t.Commit()

