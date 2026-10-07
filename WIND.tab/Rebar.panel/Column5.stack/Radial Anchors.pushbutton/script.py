from Autodesk.Revit.DB.Structure import *
import math, clr
from Autodesk.Revit.DB import Transaction, Structure, FilteredElementCollector, RadialArray, ArrayAnchorMember
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, Line, XYZ, ElementId
from Autodesk.Revit.DB import FailureSeverity, FailureProcessingResult, IFailuresPreprocessor
from Autodesk.Revit.DB import Plane, Arc, ElementTransformUtils
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
    A -> bar mark        e.g. RA100
    B -> radius          e.g. 500
    C -> Height of pin   e.g. 1500
    D -> bar size        e.g. Y25

Bar mark: must start with "RA" to be picked up by the script, e.g. RA100
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


# FPath = forms.pick_file(file_ext='xlsx', multi_file=False, unc_paths=False)
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


debug_print(str("RADIUS") + "  ---   " + "BAR MARK" + "  ---   " + str('NO. BARS') + "  ---  " + 'BAR SIZE')
debug_print("*"*45)

# Rotate construction bar to prevent splice in same location
spliceRot = 0 


t = Transaction(doc, 'Reinforce')
t.Start()
for i in range(5,150):
    if "RA" in str(xl.Cells(i, 1).Value2) :#!= None:
        radius = float(xl.Cells(i,2).Value2)/304.8
        bar_mark = str(xl.Cells(i,1).Value2)
        Yoffset = int(float(xl.Cells(i,3).Value2))/304.8
        bar_size = "Y" + str(xl.Cells(i,4).Value2)[1:3]
        bar_dia = int(float(xl.Cells(i,4).Value2[1:3]))

        #### Calculate from input parameters ####################################################################


        lap_length = 45*bar_dia/304.8
        debug_print(str(radius*304.8) + "  ---  \t " + bar_mark + "  --- \t \t " + str('##') + "  ---\t \t " + bar_size)
        debug_print("*"*45)  

        #####Rebar Shape ####################################################################
        rebar_shape = FilteredElementCollector(doc).OfClass(RebarShape).WhereElementIsElementType().ToElements()  
        for r_shape in rebar_shape:
            if r_shape.LookupParameter("Type Name").AsString() == '65':
                sc_65 = r_shape
                break

        # Rebar type ####################################################################

        all_rebar_types = FilteredElementCollector(doc) \
            .OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsElementType() \
            .ToElements()

        for rebar_type in all_rebar_types:
            rebar_name = rebar_type.get_Parameter(BuiltInParameter \
                .SYMBOL_NAME_PARAM).AsString()
            if rebar_name == bar_size:
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
        h_plinth = WTF.LookupParameter("hPlinth").AsDouble()

        WTF.Location.Point = XYZ(0,0,0)

        locPoint = WTF.Location.Point
        p_0 = locPoint

        ##### Radial Anchor Properties ####################################################################

        #cover
        top_cover = 40/304.8
        bot_cover = 50/304.8

        # start number of bars 
        no_bars = 2
        A = ((math.pi*(radius -barDia/2)*2)/no_bars)+lap_length    
        r = radius 
        x1 = r - (r*math.cos(A/(2*r)))
        whilekill = 0
        while A > 13000/304.8 or x1 > 2500/304.8:
            no_bars += 1
            A = ((math.pi*(radius -barDia/2)*2)/no_bars)+lap_length
            A = (round((A*304.8)/100)*100)/304.8
            r = radius 
            x1 = r - (r*math.cos(A/(2*r)))
            whilekill += 1
            if whilekill > 100:
                debug_print("while loop killed")
                break

        debug_print("Number of bars: " + str(no_bars))
        debug_print("A: " + str(A*304.8/1000))
        debug_print("x1: " + str(round(x1*304.8)/1000))
        debug_print("#"*45)
        ##### Build ############################################################################
        # draw a  line  to place the bar 
        preplane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, locPoint + XYZ(0,0,Yoffset))
        #plane = Plane.CreateByThreePoints(p_1, XYZ(), p_3)
        precurve = [Arc.Create(preplane, radius, 0, A/radius)]

        adjValue = (barDia)*(1.45*lap_length/A)
        totAdjValue = (adjValue + barDia)/2
        p1 = precurve[0].GetEndPoint(0) + XYZ(0,0,totAdjValue)
        p2 = precurve[0].GetEndPoint(1) - XYZ(0,0,totAdjValue)
        midplane = Plane.CreateByThreePoints(p1,p2, locPoint+ XYZ(0,0,Yoffset))
        # get normal of plane
        normal = midplane.Normal
        plane = Plane.CreateByNormalAndOrigin(normal, locPoint+ XYZ(0,0,Yoffset))
        curve = [Arc.Create(plane, radius, 0, A/radius)]

        #build construction bar
        rebarCur = Structure.Rebar.CreateFromCurvesAndShape(doc, sc_65, bar_type, None, None, WTF, XYZ.BasisZ, curve, RebarHookOrientation.Left, RebarHookOrientation.Left)
       
        # set construction bar properties
        # rebarCur.LookupParameter("A").Set(A)
        # rebarCur.LookupParameter("r").Set(r)
        rebarCur.LookupParameter("Rebar r Custom").Set(r) 
        rebarCur.LookupParameter("Rebar Quantity").Set(no_bars)
        rebarCur.LookupParameter(PARAM_MARK).Set("RADIAL ANCHORS")


        #build radial array
        RotAngle = 360*math.pi/180
        if no_bars > 2:
            elem = RadialArray.ArrayElementWithoutAssociation(doc, view, rebarCur.Id, no_bars, Line.CreateBound(p_0,XYZ.BasisZ), RotAngle, ArrayAnchorMember.Last)
            for elem in elem:
                doc.GetElement(elem).LookupParameter(PARAM_MARK).Set("RADIAL ANCHORS")
                doc.GetElement(elem).LookupParameter(PARAM_SCHEDULE_MARK).Set(bar_mark) 
                doc.GetElement(elem).LookupParameter("Rebar r Custom").Set(r) 
                # doc.GetElement(elem).LookupParameter("Rebar Spacing").Set((360.0/no_bars)/304.8)
                doc.GetElement(elem).LookupParameter("Rebar Quantity").Set(no_bars)
                rebarCur_rotate = ElementTransformUtils.RotateElement(doc, elem, Line.CreateBound(p_0,p_0+XYZ.BasisZ), spliceRot)
        else:
            doc.Delete(rebarCur.Id)
            twobarRebarCur = Structure.Rebar.CreateFromCurvesAndShape(doc, sc_65, bar_type, None, None, WTF, XYZ.BasisZ, precurve, RebarHookOrientation.Left, RebarHookOrientation.Left)
            rebarCopy = ElementTransformUtils.CopyElement(doc, twobarRebarCur.Id, XYZ(0, 0, barDia))
            elem = ElementTransformUtils.RotateElement(doc, rebarCopy[0], Line.CreateBound(p_0,XYZ.BasisZ), RotAngle/2)
            doc.GetElement(rebarCopy[0]).LookupParameter(PARAM_MARK).Set("RADIAL ANCHORS")
            doc.GetElement(rebarCopy[0]).LookupParameter(PARAM_SCHEDULE_MARK).Set(bar_mark) 
            doc.GetElement(rebarCopy[0]).LookupParameter("Rebar r Custom").Set(r)
            doc.GetElement(rebarCopy[0]).LookupParameter("A").Set(twobarRebarCur.LookupParameter("A").AsDouble())
            twobarRebarCur.LookupParameter(PARAM_MARK).Set("RADIAL ANCHORS")
            twobarRebarCur.LookupParameter(PARAM_SCHEDULE_MARK).Set(bar_mark)     
            twobarRebarCur.LookupParameter("Rebar r Custom").Set(r)
            #twobarRebarCur.LookupParameter("Rebar Spacing").Set((360.0/no_bars)/304.8)
            twobarRebarCur.LookupParameter("Rebar Quantity").Set(no_bars)     

        
        spliceRot = spliceRot + (2*lap_length/radius)
        debug_print(spliceRot)
excel.Quit()

# supress warnings ####################################################################
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SuppressWarnings())
t.SetFailureHandlingOptions(failHandler)

t.Commit()










