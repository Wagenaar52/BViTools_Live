from Autodesk.Revit.DB.Structure import * 
from Autodesk.Revit.DB.Structure import RebarShape
import math, clr
from Autodesk.Revit.DB import Transaction, Structure, FilteredElementCollector, RadialArray, ArrayAnchorMember
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, Line, XYZ
from Autodesk.Revit.DB import FailureSeverity, FailureProcessingResult,IFailuresPreprocessor
from pyrevit import forms
from Autodesk.Revit.DB import *
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

FPath =  "C:\Users\Wagner.Human\Desktop\Copy of Copy of Wolf_RebarData_RevD_V162.xlsx" 
#FPath = forms.pick_file(file_ext='xlsx', multi_file=False, unc_paths=False)

excel = Excel.ApplicationClass()
excel.Visible = False
workbook = excel.Workbooks.Open(FPath)
xl = workbook.Worksheets['A']


print(str("RADIUS") + "  ---   " + "BAR MARK" + "  ---   " + str('NO. BARS') + "  ---  " + 'BAR SIZE')
print("*"*45)

t = Transaction(doc, 'Reinforce')
t.Start()
for i in range(46,100):
    if xl.Cells(i, 4).Value2 != None:
        radius = float(xl.Cells(i,2).Value2)/304.8
        bar_mark = str(xl.Cells(i,1).Value2)
        #no_bars = int(xl.Cells(i,6).Value2)/2
        bar_size = "Y" + str(xl.Cells(i,4).Value2)[1:3]
        bar_dia = int(xl.Cells(i,4).Value2[1:3])
        i += 1


        #### Calculate from input parameters ####################################################################


        lap_length = 45*bar_dia/304.8
        print(str(radius*304.8) + "  ---  \t " + bar_mark + "  --- \t \t " + str('##') + "  ---\t \t " + bar_size)
        print("*"*45)
            
        #####Rebar Shape ####################################################################
                                                                #shape_list = ['20','32','33','34','35','37','38','39','41','42','43','45','48','49','51','52','53','54','55','60','62','65','72','73','74','75','81','85','86','99h','99j',"99z"]
                                                                # i = 0
                                                                # shape_dict = {}
                                                                # for shape in rebar_shape:
                                                                #     shape_dict[shape_list[i]] = shape.Id
                                                                #     print(shape_list[i])
                                                                #     print(shape.Id)
                                                                #     i += 1  
                                                                # # Use shape code in shape list to get shapeFamilyId (the .Name property cannot be accessed)
                                                                # sc_65 = doc.GetElement(shape_dict['65'])
                                                                # sc_65 = doc.GetElement('380762')
                                                                # sc_65 = doc.GetElement('381017')

        rebar_shape = FilteredElementCollector(doc).OfClass(RebarShape).WhereElementIsElementType().ToElements()  
        for r_shape in rebar_shape:
            if str(r_shape.ShapeFamilyId) == '410183':
                sc_65 = r_shape
                break

        # Rebar type ####################################################################

        all_rebar_types = FilteredElementCollector(doc) \
            .OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsElementType() \
            .ToElements()

        for  rebar_type in all_rebar_types:
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
        type = doc.GetElement(type_id)
        r_base = WTF.LookupParameter("rBase").AsDouble()
        h_base = WTF.LookupParameter("hBase").AsDouble()
        r_plinth = WTF.LookupParameter("rPlinth").AsDouble()
        h_plinth = WTF.LookupParameter("hPlinth").AsDouble()

        WTF.Location.Point = XYZ(0,0,0)

        locPoint = WTF.Location.Point
        p_1 = locPoint
        p_2 = locPoint + XYZ.BasisY
        p_3 = locPoint + XYZ.BasisZ
        # #### Radial Anchor Properties ####################################################################

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
                print("while loop killed")
                break

        print("Number of bars: " + str(no_bars))
        print("A: " + str(A*304.8/1000))
        print("x1: " + str(round(x1*304.8)/1000))

        ##### Build ############################################################################

        # draw a  line  to place the bar 
        plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, locPoint)
        #plane = Plane.CreateByThreePoints(p_1, XYZ(), p_3)
        curve = [Arc.Create(plane, radius, 0, A/radius)]

        adjValue = (barDia)*(1.45*lap_length/A)
        totAdjValue = (adjValue + barDia)/2
        p1 = curve[0].GetEndPoint(0) + XYZ(0,0,totAdjValue)
        p2 = curve[0].GetEndPoint(1) - XYZ(0,0,totAdjValue)
        plane = Plane.CreateByThreePoints(p1,p2, locPoint)
        # get normal of plane
        normal = plane.Normal
        plane = Plane.CreateByNormalAndOrigin(normal, locPoint)
        curve = [Arc.Create(plane, radius, 0, A/radius)]

        #build construction bar
        rebarCur = Structure.Rebar.CreateFromCurvesAndShape(doc, sc_65, bar_type, None, None, WTF, XYZ.BasisZ, curve, RebarHookOrientation.Left, RebarHookOrientation.Left)

        # set construction bar properties
        rebarCur.LookupParameter("A").Set(A)
        rebarCur.LookupParameter("r").Set(r)
        rebarCur.LookupParameter("Mark").Set(bar_mark)


        #build radial array
        RotAngle = 360*math.pi/180
        elem = RadialArray.ArrayElementWithoutAssociation(doc, view, rebarCur.Id, no_bars, Line.CreateBound(p_1,p_2), RotAngle, ArrayAnchorMember.Last)

        #get hold of all elements in radial array to change barmarks
        for elem in elem:
            doc.GetElement(elem).LookupParameter("Mark").Set(bar_mark)
    else:
        break

excel.Quit()

# supress warnings ####################################################################
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SupressWarnings())
t.SetFailureHandlingOptions(failHandler)

t.Commit()


