from Autodesk.Revit.DB.Structure import *
import math, clr
from System.Collections.Generic import List
from Autodesk.Revit.DB import Curve, Line, XYZ, Plane, SketchPlane, Family
from Autodesk.Revit.DB import Transaction, Structure, FilteredElementCollector, RadialArray, ArrayAnchorMember
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, FailureSeverity, FailureProcessingResult, IFailuresPreprocessor
from pyrevit import forms
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
import atexit
import Functions as func

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
  A -> bar mark           e.g. BC100
  B -> radius             e.g. 500
  C -> number of stools   e.g. 12
  D -> bar size           e.g. Y25
"""

rebar_shape = FilteredElementCollector(doc).OfClass(RebarShape).WhereElementIsElementType().ToElements()  
for r_shape in rebar_shape:
    if r_shape.LookupParameter("Type Name").AsString() == '99h':
        sc_99z = r_shape

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
# sheetCounter = int(xl.Cells(3, 50).Value2)
rotCount = 1
t = Transaction(doc, 'Reinforce')
t.Start()

for i in range(5,200):
    if "ST" in str(xl.Cells(i, 1).Value2):
        radius = float(xl.Cells(i,2).Value2)/304.8
        bar_mark = str(xl.Cells(i,1).Value2)
        no_stools = int(xl.Cells(i,3).Value2)
        size =  str(xl.Cells(i,4).Value2)
        i += 1

        debug_print(str(radius*304.8) + "  ----   " + bar_mark + "  ----   " + str(no_stools) + "  ---  " + size)
        debug_print("*"*30)

        deg = (360/(no_stools*2))*math.pi/180


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
        r_base = WTF.LookupParameter("rBase").AsDouble()
        h_base = WTF.LookupParameter("hBase").AsDouble()
        r_plinth = WTF.LookupParameter("rPlinth").AsDouble()
        h_cone = WTF.LookupParameter("hCone").AsDouble()


        locPoint = WTF.Location.Point
        p_1 = locPoint
        p_2 = locPoint + XYZ.BasisZ
        #### Link Properties ####################################################################

        #cover
        top_cover = 40/304.8
        bot_cover = 50/304.8

        #Rebar Properties
        if size == "Y32*":
            MandrelRatio = 13.28
        elif size == "Y25*":
            MandrelRatio = 14
        elif size == "Y20*":
            MandrelRatio = 17.5
        elif size == "Y16*":
            MandrelRatio = 17.5
        else:
            MandrelRatio = 14.5
            
        #A       
        A = (math.ceil((barDia*MandrelRatio*304.8/50)))*50/304.8 #14.5xdia and rounded up to the nearest 50mm
        if A < 500/304.8:
            A = 450/304.8
     
        #B
        ystool = (r_base - radius)*((h_cone-h_base)/(r_base - r_plinth))
        B = ystool -top_cover - bot_cover + h_base - 0.5*barDia
        #C
        C = ((math.pi*radius*2)/(no_stools*2))
        #D
        D = B
        


        origin = XYZ(locPoint.X + radius +1  , locPoint.Y -0.5*C , locPoint.Z + 500/304.8 + bot_cover+ 0.5*barDia) 
        xVec = XYZ.BasisY
        yVec = -XYZ.BasisZ
        #### Build ####################################################################
        #build construction link
        rebar = Structure.Rebar.CreateFromRebarShape(doc, sc_99z, bar_type, WTF, origin, xVec , yVec)
        #set construction link properties
        rebar.LookupParameter("A").Set(A)
        rebar.LookupParameter("B").Set(B)
        rebar.LookupParameter("C").Set(C)
        # rebar.LookupParameter("D").Set(D)
        rebar.LookupParameter(PARAM_MARK).Set("STOOLS")
        
############################################################################################################################################################################
    #build radial array
        RotAngle = 360*math.pi/180
        elem = RadialArray.ArrayElementWithoutAssociation(doc, view, rebar.Id, no_stools, Line.CreateBound(p_1,p_2), RotAngle, ArrayAnchorMember.Last)
    #get hold of all elements in radial array to change barmarks
        for elem in elem:
            doc.GetElement(elem).LookupParameter(PARAM_MARK).Set("STOOLS")
            doc.GetElement(elem).LookupParameter(PARAM_SCHEDULE_MARK).Set(bar_mark)
            doc.GetElement(elem).LookupParameter("Rebar Spacing").Set((360.0/no_stools)/304.8)
            doc.GetElement(elem).LookupParameter("Rebar Quantity").Set(no_stools)
            doc.GetElement(elem).LookupParameter("Rebar r Custom").Set(float(0))
            if (rotCount % 2) == 0:
                doc.GetElement(elem).Location.Rotate(Line.CreateBound(p_1,p_2), deg)
        rotCount += 1




## Supress warnings ################################################################
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SuppressWarnings())
t.SetFailureHandlingOptions(failHandler)
t.Commit()

#close excel    
workbook.Close()
excel.Quit()











