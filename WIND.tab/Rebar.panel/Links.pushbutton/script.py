from Autodesk.Revit.DB.Structure import * 
from Autodesk.Revit.DB.Structure import RebarShape
import math, clr
from System.Collections.Generic import List
from Autodesk.Revit.DB import Curve, Line, XYZ, Plane, SketchPlane, Family
from Autodesk.Revit.DB import Transaction, Structure, FilteredElementCollector, RadialArray, ArrayAnchorMember
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, Line, XYZ, FailureSeverity, FailureProcessingResult,IFailuresPreprocessor
from pyrevit import forms
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
view = doc.ActiveView

rebar_shape = FilteredElementCollector(doc).OfClass(RebarShape).WhereElementIsElementType().ToElements()  
for r_shape in rebar_shape:
    if r_shape.LookupParameter("Type Name").AsString() == '99z':
        sc_99z = r_shape


# rebar_shape = FilteredElementCollector(doc).OfClass(RebarShape).WhereElementIsElementType().ToElements()  
# for r_shape in rebar_shape:
#     if str(r_shape.ShapeFamilyId) == '723320':
#         sc_99z = r_shape
#         break

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
# sheetCounter = int(xl.Cells(3, 50).Value2)
rotCount = 1
t = Transaction(doc, 'Reinforce')
t.Start()

for i in range(5,200):
    if "ST" in str(xl.Cells(i, 1).Value2):
        radius = float(xl.Cells(i,2).Value2)/304.8
        bar_mark = str(xl.Cells(i,1).Value2)
        no_bars = int(xl.Cells(i,3).Value2)/2
        size =  str(xl.Cells(i,4).Value2)
        i += 1

        print(str(radius*304.8) + "  ----   " + bar_mark + "  ----   " + str(no_bars) + "  ---  " + size)
        print("*"*30)

        deg = (360/(no_bars*2))*math.pi/180


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


        locPoint = WTF.Location.Point
        p_1 = locPoint
        p_2 = locPoint + XYZ.BasisY
        p_3 = locPoint + XYZ.BasisZ
        #### Link Properties ####################################################################

        #cover
        top_cover = 40/304.8
        bot_cover = 50/304.8

        #Rebar Properties
        #A
        A = (math.ceil(barDia*14.5*304.8/50))*50/304.8 #14.5xdia and rounded up to the nearest 50mm
        #B
        ystool = (r_base - radius)*((h_cone-h_base)/(r_base - r_plinth))
        B = ystool -top_cover - bot_cover + h_base - 0.5*barDia
        #C
        C = ((math.pi*radius*2)/(no_bars*2))
        #D
        D = B
        
        # offsetDict = {"Y16": -59.3/304.8 , "Y20": -11.3/304.8 , "Y25": 86.2/304.8, "Y32": 187.5/304.8}
        # offset = offsetDict[size]

        # print(offset*304.8)
        # print(locPoint.X + radius + offset)
        # print(locPoint.X )
        # print( radius*304.8 )
        # print(offset*304.8)
        # print((locPoint.X + radius + offset)*304.8) 

        origin = XYZ(locPoint.X + radius +1  , locPoint.Y -0.5*C , locPoint.Z + 500/304.8 + bot_cover+ 0.5*barDia) 
        xVec = XYZ.BasisY
        yVec = -XYZ.BasisZ
        #### Bluid ####################################################################
        #build construction link
        rebar = Structure.Rebar.CreateFromRebarShape(doc, sc_99z, bar_type, WTF, origin, xVec , yVec)
    #set construction link properties
        rebar.LookupParameter("A").Set(A)
        rebar.LookupParameter("B").Set(B)
        rebar.LookupParameter("C").Set(C)
        rebar.LookupParameter("D").Set(D)
        rebar.LookupParameter("Mark").Set("STOOLS")

############################################################################################################################################################################

    #build radial array
        RotAngle = 360*math.pi/180
        elem = RadialArray.ArrayElementWithoutAssociation(doc, view, rebar.Id, no_bars, Line.CreateBound(p_1,p_2), RotAngle, ArrayAnchorMember.Last)
    #get hold of all elements in radial array to change barmarks
        for elem in elem:
            doc.GetElement(elem).LookupParameter("Mark").Set("STOOLS")
            doc.GetElement(elem).LookupParameter("Schedule Mark").Set(bar_mark)
            if (rotCount % 2) == 0:
                doc.GetElement(elem).Location.Rotate(Line.CreateBound(p_1,p_3), deg)
        rotCount += 1




## Supress warnings ################################################################
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SupressWarnings())
t.SetFailureHandlingOptions(failHandler)
t.Commit()

#close excel    
workbook.Close()
excel.Quit()


FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
t = Transaction(doc, 'Stretch')
t.Start()
for rebar in FEC:
    if rebar.LookupParameter("Mark").AsString() == "STOOLS":
        rebar.LookupParameter("A").Set(20)
t.Commit()



FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
t = Transaction(doc, 'Revert')
t.Start()
for rebar in FEC:
    if rebar.LookupParameter("Mark").AsString() == "STOOLS":
        tempDia = rebar.LookupParameter("Bar Diameter").AsDouble()
        rebar.LookupParameter("A").Set((math.ceil(tempDia*14.5*304.8/50))*50/304.8)
t.Commit()

