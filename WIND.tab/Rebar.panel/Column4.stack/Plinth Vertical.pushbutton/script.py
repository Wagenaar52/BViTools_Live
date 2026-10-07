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
    A -> bar mark       e.g. PV100
    B -> start radius of the rebar from the center of the plinth for PV1__ bars; for PV2__ bars this column is ignored
    C -> horizontal offset of the rebar from the inner face of the plinth (in mm, e.g. 50)
    D -> bar size       e.g. Y16
    E -> rotation switch (0 or 1) 0 means no rotation, 1 means rotate by half of the angle between the bars
    F -> empty
    G -> no bars factor (e.g. 1.5 means 1.5 times the number of anchor bolts, 2 means twice the number of anchor bolts)

- Bar mark: must be "PV100" to be picked up by the script
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

def getFilePath():
    FPath = forms.pick_file(file_ext='xlsx', multi_file=False, unc_paths=False)
    if not FPath:
        raise SystemExit
    return FPath

def getAnchorBolts(xl):
    AC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
    AnchorBolts = 0
    for Anchorcage in AC:
        try:
            if "1PA_AnchorCage_Assembly 2" in Anchorcage.Name: 
                AnchorCage = Anchorcage
                AnchorBolts = AnchorCage.LookupParameter("nBolts").AsInteger()
            elif "1PA_AnchorCage_Assembly1" in Anchorcage.Name:
                AnchorCage = Anchorcage
                AnchorBolts = AnchorCage.LookupParameter("nBolts").AsInteger()
        except:
            pass
        
    if AnchorBolts is None or AnchorBolts == 0:
        AnchorBolts = int(xl.Cells(3, 3).Value2)
        
    return AnchorBolts


#### Input from excel sheet ####################################################################

# FPath =  forms.pick_file(file_ext='xlsx', multi_file=False, unc_paths=False) #"C:\Users\Wagner.Human\Desktop\Wolf_RebarData_RevD_V162r5.xlsx" #
FPath = getFilePath()

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

wtf_collection = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
WTF = None
AC  = None
for element in wtf_collection:
    if element.Name == "1PA_WTF_SteelTower":
        WTF = element
    elif element.Name == "1PA_AnchorCage_Assembly":
        AC = element
    if WTF is not None and AC is not None:
        break


type_id = WTF.GetTypeId()
wtf_type = doc.GetElement(type_id)
r_base = WTF.LookupParameter("rBase").AsDouble()
h_base = WTF.LookupParameter("hBase").AsDouble()
r_plinth = WTF.LookupParameter("rPlinth").AsDouble()
h_cone = WTF.LookupParameter("hCone").AsDouble()
h_plinth = WTF.LookupParameter("hPlinth").AsDouble()
slabSlope = (h_cone - h_base)/(r_base - r_plinth)
rPitOuter = WTF.LookupParameter("rVoidOuter").AsDouble()
rPitInner = WTF.LookupParameter("rVoidInner").AsDouble()
locPoint = WTF.Location.Point
dGroutMid = WTF.LookupParameter("dGroutMiddle").AsDouble()
wGroutTop = WTF.LookupParameter("wGroutTop").AsDouble()
rTower = WTF.LookupParameter("rTower").AsDouble()
hPit = WTF.LookupParameter("hBottomVoid").AsDouble()
wbotFlange = AC.LookupParameter("wFlangeBot").AsDouble()

LAP_CONST = 50
MM_TO_FT = 1.0 / 304.8


#####Rebar Shape ####################################################################

rebar_shape = FilteredElementCollector(doc).OfClass(RebarShape).WhereElementIsElementType().ToElements()   

for r_shape in rebar_shape:
    if r_shape.LookupParameter("Type Name").AsString() == '99c':
        sc_99c = r_shape


####  Get number of Anchor Bolts#################################################################
AnchorBolts = getAnchorBolts(xl)


for i in range(5,200):
    if "PF1" in str(xl.Cells(i, 1).Value2):
        plinthFaceConDia = int(str(xl.Cells(i,4).Value2)[1:])/304.8
    elif "PH1" in str(xl.Cells(i, 1).Value2):
        plinthHorConDia = int(str(xl.Cells(i,4).Value2)[1:])/304.8
    elif "GR1" in str(xl.Cells(i, 1).Value2):
        botGridDia = int(str(xl.Cells(i,4).Value2)[1:])/304.8
    elif "GR2" in str(xl.Cells(i, 1).Value2):
        mid1GridDia = int(str(xl.Cells(i,4).Value2)[1:])/304.8
    elif "GR3" in str(xl.Cells(i, 1).Value2):
        mid2GridDia = int(str(xl.Cells(i,4).Value2)[1:])/304.8
    elif "GR4" in str(xl.Cells(i, 1).Value2):
        topGridDia = int(str(xl.Cells(i,4).Value2)[1:])/304.8

    
# try GR4 if it exists, if not use mid2 grid dia for the top bars

if topGridDia:
    gridDia = topGridDia
elif mid2GridDia:
    gridDia = mid2GridDia
elif mid1GridDia:
    gridDia = mid1GridDia   


#### Transaction ################################################################################


t = Transaction(doc, 'Reinforce Plinth Verticals')
t.Start()

for i in range(5,200):
    if "PV20" in str(xl.Cells(i, 1).Value2):
        bar_mark = str(xl.Cells(i,1).Value2)
        no_bars_factor = float(xl.Cells(i,7).Value2)
        size = "Y" + str(xl.Cells(i,4).Value2)[1:]
        RotSwitch = float(xl.Cells(i,5).Value2)
        horOffsetInner = float(xl.Cells(i,3).Value2)/304.8
        lap_length = LAP_CONST * int(str(xl.Cells(i,4).Value2)[1:]) * MM_TO_FT


        debug_print("#"*50)
        debug_print("  --   " + "BAR MARK"+ "  ----   " + "NO BARS FACTOR" + "  ---  " + "SIZE"+ "  ---  " + "ROTATION SWITCH"  )
        debug_print("  --   " + bar_mark + " \t ---- \t\t\t  " + str(no_bars_factor) + "  \t\t---  " + size + "  --- \t\t\t " + str(RotSwitch)   )
        debug_print("*"*10)
        
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

        

        #cover
        top_cover = 40/304.8
        bot_cover = 50/304.8

        A = (math.ceil(barDia*14.5*304.8/50))*50/304.8 #14.5xdia and rounded up to the nearest 50mm

        ACoutRad = max(wGroutTop, wbotFlange)
        r = r_plinth- top_cover - plinthFaceConDia #-  (barDia/2))-(horOffsetInner+(barDia/2)+rTower+(ACoutRad/2)))/2

        inX = rTower + ACoutRad/2  + (barDia/2) + horOffsetInner
        outX = r_plinth - top_cover - barDia/2 - plinthFaceConDia 

        botZ = -hPit+bot_cover + barDia/2 +(2*botGridDia) + 12/304.8 #12mm extra cover for the bottom bar for Y12's to ensure level placement
        topZ = h_plinth -top_cover - barDia/2 -plinthHorConDia
        #### Rebar Shape Properties ####################################################################

        # place points
        rebar_p1 = locPoint+XYZ(inX              , 0 ,  botZ + h_cone/2 +lap_length/2)
        rebar_p2 = locPoint+XYZ(inX           , 0 ,  botZ)
        rebar_p3 = locPoint+XYZ(outX           , 0 ,  botZ)
        rebar_p4 = locPoint+XYZ(outX           , 0 ,  topZ)
        rebar_p5 = locPoint+XYZ(inX           , 0 ,  topZ)
        rebar_p6 = locPoint+XYZ(inX           , 0 ,  botZ + h_cone/2 -lap_length/2)


        #place curves
        curve1 = Line.CreateBound(rebar_p1, rebar_p2)
        curve2 = Line.CreateBound(rebar_p2, rebar_p3)
        curve3 = Line.CreateBound(rebar_p3, rebar_p4)
        curve4 = Line.CreateBound(rebar_p4, rebar_p5)
        curve5 = Line.CreateBound(rebar_p5, rebar_p6)

        geomPlane = Plane.CreateByThreePoints(rebar_p5, rebar_p2, rebar_p3)
        sketch = SketchPlane.Create(doc, geomPlane)

        # model_line = doc.Create.NewModelCurve(curve1, sketch)
        # model_line = doc.Create.NewModelCurve(curve2, sketch)
        # model_line = doc.Create.NewModelCurve(curve3, sketch)
        # model_line = doc.Create.NewModelCurve(curve4, sketch)
        # model_line = doc.Create.NewModelCurve(curve5, sketch)
   

        #### Cast the list to IList<Curve>
    
        curve_list99c = List[Curve]([curve1, curve2, curve3, curve4, curve5])

        #### Bluid ####################################################################


        
        rebar = Structure.Rebar.CreateFromCurves(doc, 
                                            RebarStyle.Standard, 
                                            bar_type, 
                                            None, 
                                            None, 
                                            WTF, 
                                            XYZ.BasisY, 
                                            curve_list99c, 
                                            RebarHookOrientation.Left, 
                                            RebarHookOrientation.Left,
                                            1,                          	#bool useExistingShapeIfPossible,
                                            0)                          	#bool createNewShape

        rebar.LookupParameter("Rebar r Custom").Set(0)
        rebar.LookupParameter("Rebar Spacing").Set(360/(no_bars)/304.8)    
        rebar.LookupParameter("Rebar Quantity").Set(int(no_bars))


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
            doc.GetElement(elm).LookupParameter(PARAM_MARK).Set("PLINTH VERTICAL")
            doc.GetElement(elm).LookupParameter(PARAM_SCHEDULE_MARK).Set(bar_mark)
            doc.GetElement(elm).LookupParameter("Rebar r Custom").Set(0)
            doc.GetElement(elm).LookupParameter("Rebar Spacing").Set(360/(no_bars)/304.8)
            doc.GetElement(elm).LookupParameter("Rebar Quantity").Set(int(no_bars))

        debug_print('#'*50)
        debug_print("PV200 ran")
        debug_print('#'*50)
    #    delete construction bar
    #     doc.Delete(rebar.Id)  
    elif "PV10" in str(xl.Cells(i, 1).Value2):
        bar_mark = str(xl.Cells(i,1).Value2)
        no_bars_factor = float(xl.Cells(i,7).Value2)
        size = "Y" + str(xl.Cells(i,4).Value2)[1:]
        RotSwitch = float(xl.Cells(i,5).Value2)
        horOffsetInner = float(xl.Cells(i,3).Value2)/304.8
        startRad = float(xl.Cells(i,2).Value2)/304.8
        lap_length = LAP_CONST * int(str(xl.Cells(i,4).Value2)[1:]) * MM_TO_FT


        debug_print("#"*50)
        debug_print("  --   " + "BAR MARK"+ "  ----   " + "NO BARS FACTOR" + "  ---  " + "SIZE"+ "  ---  " + "ROTATION SWITCH"  )
        debug_print("  --   " + bar_mark + " \t ---- \t\t\t  " + str(no_bars_factor) + "  \t\t---  " + size + "  --- \t\t\t " + str(RotSwitch)   )
        debug_print("*"*10)
        
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

        

        #cover
        top_cover = 40/304.8
        bot_cover = 50/304.8

        # A = (math.ceil(barDia*14.5*304.8/50))*50/304.8 #14.5xdia and rounded up to the nearest 50mm


        outRad = (rTower - max(wGroutTop, wbotFlange)/2 - horOffsetInner -barDia/2) 
        botZ = -hPit+bot_cover + barDia/2 +(2*botGridDia) + 12/304.8 #12mm extra cover for the bottom bar for Y12's to ensure level placement
        topZ = h_plinth -top_cover - barDia/2 -gridDia
        #### Rebar Shape Properties ####################################################################

        # place points
        rebar_p1 = locPoint+XYZ(startRad               , 0 ,  botZ + h_cone/2 +lap_length/2)
        rebar_p2 = locPoint+XYZ(startRad               , 0 ,  botZ)
        rebar_p3 = locPoint+XYZ(outRad                 , 0 ,  botZ)
        rebar_p4 = locPoint+XYZ(outRad                 , 0 ,  topZ)
        rebar_p5 = locPoint+XYZ(startRad               , 0 ,  topZ)
        rebar_p6 = locPoint+XYZ(startRad               , 0 ,  botZ + h_cone/2 -lap_length/2)

        #place curves
        curve1 = Line.CreateBound(rebar_p1, rebar_p2)
        curve2 = Line.CreateBound(rebar_p2, rebar_p3)
        curve3 = Line.CreateBound(rebar_p3, rebar_p4)
        curve4 = Line.CreateBound(rebar_p4, rebar_p5)
        curve5 = Line.CreateBound(rebar_p5, rebar_p6)
   

        geomPlane = Plane.CreateByThreePoints(rebar_p5, rebar_p2, rebar_p3)
        sketch = SketchPlane.Create(doc, geomPlane)

        # model_line = doc.Create.NewModelCurve(curve1, sketch)
        # model_line = doc.Create.NewModelCurve(curve2, sketch)
        # model_line = doc.Create.NewModelCurve(curve3, sketch)
        # model_line = doc.Create.NewModelCurve(curve4, sketch)
        # model_line = doc.Create.NewModelCurve(curve5, sketch)
   

        #### Cast the list to IList<Curve>
    
        curve_list99c = List[Curve]([curve1, curve2, curve3, curve4, curve5])

        #### Bluid ####################################################################

        
        rebar = Structure.Rebar.CreateFromCurves(doc, 
                                            RebarStyle.Standard, 
                                            bar_type, 
                                            None, 
                                            None, 
                                            WTF, 
                                            XYZ.BasisY, 
                                            curve_list99c, 
                                            RebarHookOrientation.Left, 
                                            RebarHookOrientation.Left,1,0)
        
        rebar.LookupParameter("Rebar r Custom").Set(0)
        rebar.LookupParameter("Rebar Spacing").Set(360/(no_bars)/304.8)    
        rebar.LookupParameter("Rebar Quantity").Set(int(no_bars))


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
            doc.GetElement(elm).LookupParameter(PARAM_MARK).Set("PLINTH VERTICAL")
            doc.GetElement(elm).LookupParameter(PARAM_SCHEDULE_MARK).Set(bar_mark)
            doc.GetElement(elm).LookupParameter("Rebar r Custom").Set(0)
            doc.GetElement(elm).LookupParameter("Rebar Spacing").Set(360/(no_bars)/304.8)    
            doc.GetElement(elm).LookupParameter("Rebar Quantity").Set(int(no_bars))

        debug_print('#'*50)
        debug_print("PV100 ran")
        debug_print('#'*50)
## Supress warnings ################################################################
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SuppressWarnings())
t.SetFailureHandlingOptions(failHandler)
t.Commit()









