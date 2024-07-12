from Autodesk.Revit.DB.Structure import * 
from Autodesk.Revit.DB import CurveByPoints, ReferencePointArray, ReferencePoint, CurveArray, PolyLine, Plane, SketchPlane, ElementTransformUtils, Arc
from System.Collections.Generic import List
from Autodesk.Revit.DB import Curve, Line, XYZ, ElementId
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
# AC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
# for AnCage in AC:
#     if "1PA_AnchorCage_Assembly 2" or "1PA_AnchorCage_Assembly"in AnCage.Name: 
#         AnchorCage = AnCage
#         break
# print("#"*50)
# print(AnchorCage.Name)
# rBoltInner = AnchorCage.LookupParameter("rBoltInner").AsDouble()

# #####Rebar Shape ####################################################################

rebar_shape = FilteredElementCollector(doc).OfClass(RebarShape).WhereElementIsElementType().ToElements()   

                
for r_shape in rebar_shape:
    if str(r_shape.ShapeFamilyId) == '5323829':
        sc_20 = r_shape
        
for r_shape in rebar_shape:
    if str(r_shape.ShapeFamilyId) == '5379978':
        sc_38 = r_shape

## define function to determine length of rebar ####################################################################

def bar_length(radius, offset):
    return 1000/304.8 if (2*math.sqrt(radius**2 - offset**2)<1000/304.8) else 2*math.sqrt(radius**2 - offset**2)


def range_length(radius, spacing):
    l = math.sqrt(radius**2 - (500/304.8)**2)
    l = math.floor(l/spacing)*spacing
    return l


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
        i += 1

        print("#"*50)
        print("  --   " + "BAR MARK"+ "  ----   " + "END RAD" + "  ---  " + "SIZE"+ "  ---  " + "HEIGHT" + "  ---  " + "SPACING" + "  ---  " )
        print("  --   " + bar_mark + " \t ---- \t\t\t  " + str(endRad*304.7) + "  \t\t---  " + size + "  --- \t\t\t " + str(height*304.8) + "  --- \t\t\t " + str(Spacing*304.8) )
        print("*"*10)
        

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
            rebar20.LookupParameter("Mark").Set("GRIDS")
            rebar20.LookupParameter("Schedule Mark").Set(bar_mark+str(alphaList[MarkCount2]))
            count += 1
            previous_bar_y = bar_y
            print("MarkCount: " + str(MarkCount2))
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
            rebar.LookupParameter("Mark").Set("GRIDS")
            print("rebar set")

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
            rebar20.LookupParameter("Mark").Set("GRIDS")
            rebar20.LookupParameter("Schedule Mark").Set(bar_mark+str(alphaList[MarkCount2]))
            count += 1
            previous_bar_y = bar_y
            print("MarkCount: " + str(MarkCount2))
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
            rebar.LookupParameter("Mark").Set("GRIDS")
            print("rebar set")



        print('#'*50)
        print("rebar ran")
        print('#'*50)

    # except Exception as e:
    #         print("Error at row: " + str(i) + " #### " + str(e))
            
 #       delete construction bar
        #doc.Delete(rebar.Id)  

## Supress warnings ################################################################
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SupressWarnings())
t.SetFailureHandlingOptions(failHandler)
t.Commit()

#close excel
workbook.Close(False)
excel.Quit()

#### Update A ####################################################################

FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
scheduleMarkList = []
for elem in FEC:
    if elem.LookupParameter("Schedule Mark").AsString() not in scheduleMarkList and "GR" in elem.LookupParameter("Schedule Mark").AsString() and "GRIDS" in elem.LookupParameter("Mark").AsString() :
        scheduleMarkList.append(elem.LookupParameter("Schedule Mark").AsString())

print(scheduleMarkList)

t = Transaction(doc, "Update A")    
t.Start()
for i in range(len(scheduleMarkList)):
    if "GR" in scheduleMarkList[i]:
        sum_A = 0
        count_A = 0
        A_max = 0
        for elem in FEC:
            if elem.LookupParameter("Schedule Mark").AsString() == scheduleMarkList[i]:
                sum_A += elem.LookupParameter("A").AsDouble()
                count_A += 1
                if elem.LookupParameter("A").AsDouble() > A_max:
                    A_max = elem.LookupParameter("A").AsDouble()
        A = round((sum_A/count_A)*10)/10
        A = round(A*304.8)/304.8
        print(scheduleMarkList[i])
        print(A*304.8)
        print(A_max*304.8)
        print(count_A)
    print(A*304.8)
    print(A_max*304.8)

    for elem in FEC:
        if elem.LookupParameter("Schedule Mark").AsString() == scheduleMarkList[i]:
            elem.LookupParameter("A").Set(A_max)#((round((A_max*304.8)/10)*10)/304.8))



t.Commit()






































        # count = 0   
        # MarkCount1 = 0
        # MarkCount2 = 0
        # rebarList = []
        # for count in range(0, int((range_length(endRad,Spacing)/Spacing))+1):
        #     bar_y = bar_length(endRad,abs(-range_length(endRad,Spacing) + count*Spacing))/2
        #     MarkCount2 += 1
        #     if count == 0:
        #         previous_bar_y = bar_y
        #     if count*Spacing <= range_length(endRad,Spacing):
        #         if bar_y < (previous_bar_y + Spacing):
        #             bar_y = previous_bar_y
        #             MarkCount2 = MarkCount1

        #     rebar_p1 = locPoint+XYZ(-bar_y , -range_length(endRad,Spacing)+ count*Spacing        , height-barDia/2)
        #     rebar_p2 = locPoint+XYZ(bar_y  , -range_length(endRad,Spacing)+ count*Spacing        , height-barDia/2)


        #     curve1 = Line.CreateBound(rebar_p1, rebar_p2)

        #     rebar20 = Structure.Rebar.CreateFromCurves(doc, 
        #                                         RebarStyle.Standard, 
        #                                         bar_type, 
        #                                         None, 
        #                                         None, 
        #                                         WTF, 
        #                                         XYZ.BasisZ, 
        #                                         [curve1], 
        #                                         RebarHookOrientation.Left, 
        #                                         RebarHookOrientation.Left,1,0)
        #     MarkCount1 = MarkCount2
        #     rebar20.LookupParameter("Mark").Set("GRIDS")
        #     rebar20.LookupParameter("Schedule Mark").Set(bar_mark+str(alphaList[MarkCount2]))
        #     count += 1
        #     previous_bar_y = bar_y
        #     print("MarkCount: " + str(MarkCount2))
        #     rebarList.append(rebar20)
        #     # if count*Spacing == range_length(endRad,Spacing):
        #     #         rebarList.remove(rebar20)
        #     #         print("rebarList: " + str(len(rebarList)))
        #     #         print(rebarList)

        # rebarListId =[]
        # for rebar in rebarList[:-1]:
        #     rebarListId.append(rebar.Id)
        # element_id = List[ElementId](rebarListId)
        
        # ElementTransformUtils.CopyElements(doc, element_id, XYZ(0,0,0))#,None,None)
        # ElementTransformUtils.RotateElements(doc, element_id, Line.CreateBound(XYZ(0,0,0),XYZ(0,0,1)), math.pi)


########################################################################################################################################################################################################################

        # count = 0   
        # rebarList = []
        # for count in range(0, int((range_length(endRad,Spacing)/Spacing))+1):
        #     bar_y = bar_length(endRad,abs(-range_length(endRad,Spacing) + count*Spacing))/2
        #     if count == 0:
        #         previous_bar_y = bar_y
        #     if count*Spacing <= range_length(endRad,Spacing):
        #         if bar_y < (previous_bar_y + Spacing):
        #             bar_y = previous_bar_y

        #     rebar_p1 = locPoint+XYZ(-bar_y , -range_length(endRad,Spacing)+ count*Spacing        , height-barDia/2)
        #     rebar_p2 = locPoint+XYZ(bar_y  , -range_length(endRad,Spacing)+ count*Spacing        , height-barDia/2)

        #     curve1 = Line.CreateBound(rebar_p1, rebar_p2)

        #     rebar20 = Structure.Rebar.CreateFromCurves(doc, 
        #                                         RebarStyle.Standard, 
        #                                         bar_type, 
        #                                         None, 
        #                                         None, 
        #                                         WTF, 
        #                                         XYZ.BasisZ, 
        #                                         [curve1], 
        #                                         RebarHookOrientation.Left, 
        #                                         RebarHookOrientation.Left,1,0)
        #     rebar20.LookupParameter("Mark").Set("GRIDS")
        #     rebar20.LookupParameter("Schedule Mark").Set(bar_mark)
        #     count += 1
        #     previous_bar_y = bar_y
            
        #     rebarList.append(rebar20)

        # rebarListId =[]
        # for rebar in rebarList[:-1]:
        #     rebarListId.append(rebar.Id)
        # element_id = List[ElementId](rebarListId)
        
        # ElementTransformUtils.CopyElements(doc, element_id, XYZ(0,0,0))#,None,None)
        # ElementTransformUtils.RotateElements(doc, element_id, Line.CreateBound(XYZ(0,0,0),XYZ(0,0,1)), math.pi)




        
        # for count in range(0, int((range_length(endRad,Spacing)/Spacing)*2)+1):
        #     rebar_p1 = locPoint+XYZ(bar_length(endRad,abs(-range_length(endRad,Spacing)+ count*Spacing))/2, -range_length(endRad,Spacing)+ count*Spacing        ,height+barDia/2)
        #     rebar_p2 = locPoint+XYZ(-bar_length(endRad,abs(-range_length(endRad,Spacing)+ count*Spacing))/2,  -range_length(endRad,Spacing)+ count*Spacing      ,height+barDia/2)
        #     curve1 = Line.CreateBound(rebar_p1, rebar_p2)
        #     # geomPlane = Plane.CreateByThreePoints(rebar_p1, rebar_p2, XYZ(0,0,height))
        #     # sketch = SketchPlane.Create(doc, geomPlane)
        #     # model_line = doc.Create.NewModelCurve(curve1, sketch)
        #     rebar_20 = Structure.Rebar.CreateFromCurves(doc, 
        #                                         RebarStyle.Standard, 
        #                                         bar_type, 
        #                                         None, 
        #                                         None, 
        #                                         WTF, 
        #                                         XYZ.BasisZ, 
        #                                         [curve1], 
        #                                         RebarHookOrientation.Left, 
        #                                         RebarHookOrientation.Left,1,0)
        #     rebar_20.LookupParameter("Mark").Set("GRIDS")
        #     rebar_20.LookupParameter("Schedule Mark").Set(bar_mark)
        #     count += 1
        #     # rebarList.append(rebar_20)


    # for rebar in rebarList:
    #     rebar.LookupParameter("GRIDS").Set(bar_mark)
    #     rebar.LookupParameter("Schedule Mark").Set(bar_mark)





