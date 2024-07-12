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

FPath = forms.pick_file(file_ext='xlsx', multi_file=False, unc_paths=False)

excel = Excel.ApplicationClass()
excel.Visible = False
workbook = excel.Workbooks.Open(FPath)
xl = workbook.Worksheets['A']


####  Get number of Anchor Bolts#################################################################
AC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
for Anchorcage in AC:
    if Anchorcage.Name == "1PA_AnchorCage_Assembly 2":
        AnchorCage = Anchorcage
        break
if AnchorCage == None:
    forms.alert("Anchor Cage not found")
    

AnchorBolts = AnchorCage.LookupParameter("nBolts").AsInteger()


SlabFace_dia = 0

t = Transaction(doc, 'Reinforce')
t.Start()
# for i in range(5,200):
#     if "SF100" in str(xl.Cells(i, 1).Value2):
#         SlabFace_dia = int(str(xl.Cells(i,4).Value2)[1:])/304.8
        


for i in range(5,200):
    if "BV100" in str(xl.Cells(i, 1).Value2):
        bar_mark = str(xl.Cells(i,1).Value2)
        no_bars_factor = float(xl.Cells(i,7).Value2)
        size = "Y" + str(xl.Cells(i,4).Value2)[1:]
        RotSwitch = float(xl.Cells(i,5).Value2)
        RadiusOut = float(xl.Cells(i,2).Value2)/304.8
        RadiusIN = float(xl.Cells(i,3).Value2)/304.8
        i += 1

        print("  ----   " + "BAR MARK"+ "  ----   " + "NO BARS FACTOR" + "  ---  " + "SIZE")
        print("  ----   " + bar_mark + "  ----   " + str(no_bars_factor) + "  ---  " + size)
        print("*"*30)

        no_bars = no_bars_factor*AnchorBolts

  

        #####Rebar Shape ####################################################################

        rebar_shape = FilteredElementCollector(doc).OfClass(RebarShape).WhereElementIsElementType().ToElements()   

                      
        for r_shape in rebar_shape:
            if str(r_shape.ShapeFamilyId) == '437791':
                sc_97 = r_shape
                break
        
        for r_shape in rebar_shape:
            if str(r_shape.ShapeFamilyId) == '384440':
                sc_37 = r_shape
                break
        
        for r_shape in rebar_shape:
            if str(r_shape.ShapeFamilyId) == '383060':
                sc_20 = r_shape
                break

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


        AC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
        for Anchorcage in AC:
            if Anchorcage.Name == "1PA_AnchorCage_Assembly 2":
                AnchorCage = Anchorcage
                break

        AnchorBoltNr = AnchorCage.LookupParameter("nBolts").AsInteger()
        # print(AnchorBoltNr)

        #cover
        top_cover = 40/304.8
        bot_cover = 50/304.8
       


        locPoint = WTF.Location.Point

        #### Rebar Shape Properties ####################################################################


        #Rebar Properties
        #A
        A = (RadiusOut-RadiusIN) # + (slabSlope*(h_base*2))**2)**0.5
        #B
        y_Rangetop = h_base + slabSlope*(r_base-RadiusOut) -top_cover -barDia*1.5
        y_Rangebot = bot_cover + barDia
        B = y_Rangetop - y_Rangebot
        #C
        C =  A*slabSlope
        #D
        D = 13.5*barDia
        #J
        J = B+C


        print("A: " + str(A*304.8))
        print("B: " + str(B*304.8))
        print("C: " + str(C*304.8))
        print("D: " + str(D*304.8))
        print("J: " + str(J*304.8))



        #### Bluid ####################################################################

        rebar_p1 = locPoint+XYZ(RadiusOut- A            ,0,bot_cover+J)
        rebar_p2 = locPoint+XYZ(RadiusOut-(barDia/2)    ,0,bot_cover+B+barDia)
        rebar_p3 = locPoint+XYZ(RadiusOut-(barDia/2)    ,0,bot_cover+barDia/2)
        rebar_p4 = locPoint+XYZ(RadiusOut-A             ,0,bot_cover+barDia/2)
        rebar_p5 = locPoint+XYZ(RadiusOut- A +D         ,0,bot_cover+barDia/2)

        # E = (C**2 + A**2)**0.5
        # dx_p6 = D*(A/E)
        # dy_p6 = D*(C/E)

        rebar_p6 = locPoint+XYZ(RadiusOut- A        ,0,bot_cover+D + barDia/2)


        curve1 = Line.CreateBound(rebar_p5, rebar_p4)
        curve2 = Line.CreateBound(rebar_p4, rebar_p1)
        curve3 = Line.CreateBound(rebar_p1, rebar_p2)
        curve4 = Line.CreateBound(rebar_p2, rebar_p3)
        curve5 = Line.CreateBound(rebar_p3, rebar_p4)
        curve6 = Line.CreateBound(rebar_p4, rebar_p6)

        # curveList =  [curve1, curve2, curve3, curve4, curve5, curve6]
       
        # curveList = [curve3, curve4]
       
        # Add curves to the CurveArray
        # curve_array = CurveArray()
                
        # curve_array.Append(curve1)
        # curve_array.Append(curve2)
        # curve_array.Append(curve3)
        # curve_array.Append(curve4)

        #rebar = Structure.Rebar.CreateFromRebarShape(doc, sc_97, bar_type, WTF,  rebar_p1+XYZ(-1,0,0), -XYZ.BasisX, -XYZ.BasisZ)
        

        
        # Create a list of Curve objects
        curves = [ curve1, curve2, curve3, curve4, curve5, curve6]
        
        # Cast the list to IList<Curve>
        curve_list = List[Curve](curves)

        # Create a polyline from the points
        # points = [rebar_p1, rebar_p2, rebar_p3, rebar_p4, rebar_p5, rebar_p6]
        # polyline = PolyLine.Create(points)
        # #draw polyline from curveList   

        # curve = PolyLine.Create(curveList)
        
        geomPlane = Plane.CreateByThreePoints(rebar_p1, rebar_p2, rebar_p3)
        # Create a sketch plane in current document
        sketch = SketchPlane.Create(doc, geomPlane)
        #create model lines form curve

        # model_line = doc.Create.NewModelCurve(curve1, sketch)
        # model_line = doc.Create.NewModelCurve(curve2, sketch)
        # model_line = doc.Create.NewModelCurve(curve3, sketch)
        # model_line = doc.Create.NewModelCurve(curve4, sketch)
        # model_line = doc.Create.NewModelCurve(curve5, sketch)
        # model_line = doc.Create.NewModelCurve(curve6, sketch)
       
        #for curve in curve_array:
        #for curve in curveList:
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
        
        #for curve in curves:
        # rebar = Structure.Rebar.CreateFromCurvesAndShape(doc, 
        #                                         sc_20, #RebarStyle.Standard,
        #                                         bar_type, 
        #                                         None,
        #                                         None, 
        #                                         WTF, 
        #                                         XYZ.BasisY, 
        #                                         curve_list,
        #                                         RebarHookOrientation.Left, 
        #                                         RebarHookOrientation.Left)#,1,0)



    #set construction link properties
        # rebar.LookupParameter("A").Set(A)
        # rebar.LookupParameter("B").Set(B)
        # rebar.LookupParameter("C").Set(C)
        # rebar.LookupParameter("D").Set(D)
        
        
        #build radial array
        RotAngle = 360*math.pi/180
        if no_bars <= 200:
            elem = RadialArray.ArrayElementWithoutAssociation(doc,
                                                            view, 
                                                            rebar.Id,
                                                            no_bars, 
                                                            Line.CreateBound(locPoint,locPoint+XYZ.BasisZ), 
                                                            RotAngle, ArrayAnchorMember.Last)
        if no_bars > 200:
            elem = RadialArray.ArrayElementWithoutAssociation(doc,
                                                view, 
                                                rebar.Id,
                                                no_bars/2, 
                                                Line.CreateBound(locPoint,locPoint+XYZ.BasisZ), 
                                                RotAngle, ArrayAnchorMember.Last)
            elem2 = RadialArray.ArrayElementWithoutAssociation(doc,
                                                            view, 
                                                            rebar.Id,
                                                            no_bars/2, 
                                                            Line.CreateBound(locPoint,locPoint+XYZ.BasisZ), 
                                                            RotAngle, ArrayAnchorMember.Last)
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
            rebar.LookupParameter("Mark").Set("BASE VERTICALS")
            rebar.LookupParameter("Schedule Mark").Set(bar_mark)

            #rotare elem2 with half the angle
            for elm in elem2:
                doc.GetElement(elm).Location.Rotate( Line.CreateBound(locPoint, locPoint + XYZ.BasisZ), (RotAngle)/(no_bars))
                doc.GetElement(elm).LookupParameter("Mark").Set("BASE VERTICALS")
                doc.GetElement(elm).LookupParameter("Schedule Mark").Set(bar_mark)
            for elm in elem:
                doc.GetElement(elm).LookupParameter("Mark").Set("BASE VERTICALS")
                doc.GetElement(elm).LookupParameter("Schedule Mark").Set(bar_mark)  

        
        #Rotate rebar 
        if no_bars <= 200:
            for elm in elem:
                doc.GetElement(elm).Location.Rotate( Line.CreateBound(locPoint, locPoint + XYZ.BasisZ), RotSwitch*RotAngle/(no_bars*4))
                doc.GetElement(elm).LookupParameter("Mark").Set("BASE VERTICALS")
                doc.GetElement(elm).LookupParameter("Schedule Mark").Set(bar_mark)
            
 #       delete construction bar
        #doc.Delete(rebar.Id)  

## Supress warnings ################################################################
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SupressWarnings())
t.SetFailureHandlingOptions(failHandler)
t.Commit()

