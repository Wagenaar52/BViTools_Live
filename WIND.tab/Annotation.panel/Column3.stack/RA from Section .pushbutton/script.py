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
import Functions as func

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
view = doc.ActiveView
__doc__ = """
Version = 2.0
Date = 04.05.2026
_____________________________________________________________________
1. Create instances of BVI-DET-SectionBar family in the section view on Rebar dwg 2 
2. set corect bar diameter type
3. Fill bar mark parameter (only one Barmark will be constructed per unique bar mark)
4. Fill lapping orientation parameter (H or V) in the section view on Rebar dwg 2
"""


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

def polar_to_car(radius, angle_radians):
    x = radius * math.cos(angle_radians)
    y = radius * math.sin(angle_radians)
    return x, y


BarDict = {}

FEC = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_DetailComponents).OfCategory(BuiltInCategory.OST_DetailComponents).WhereElementIsNotElementType().ToElements()

for elem in FEC:
    try:
        if elem.Symbol.FamilyName == "BVI-DET-SectionBar":
            BarDict[elem.LookupParameter("Bar Mark").AsString()] = elem
            print("Section Bar Mark: " + elem.LookupParameter("Bar Mark").AsString())
            print(elem.Name)
            print(elem.Location.Point)
            print("#########################")
    except:
        pass
    
print('Dictionary of Section Bars:')
for key, value in BarDict.items():
    print(str(key) + " : " + str(value))
print("*"*45)
print("Number of Bars: " + str(len(BarDict)))
print("*"*45)
print("*"*45)


for bar in BarDict:
    print("Bar Mark: " + str(bar))
    print("Bar Element: " + str(BarDict[bar]))
    print("Bar Location: " + str(BarDict[bar].Location.Point))
    

# Rotate construction bar to prevent splice in same location
spliceRot = 0 


t = Transaction(doc, 'Reinforce')
t.Start()
for bar in BarDict:
    radius = (abs(BarDict[bar].Location.Point.X)**2 + abs(BarDict[bar].Location.Point.Y)**2)**0.5
    bar_mark = bar
    Yoffset = BarDict[bar].Location.Point.Z
    bar_size = BarDict[bar].Name
    bar_dia = int(str(bar_size)[1:3])
    lapOrientation = BarDict[bar].LookupParameter("Lap Orientation").AsString().upper()



    #### Calculate from input parameters ####################################################################


    lap_length = 45*bar_dia/304.8
    print(str(radius*304.8) + "  ---  \t " + bar_mark + "  --- \t \t " + str('##') + "  ---\t \t " + bar_size)
    print("*"*45)  
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
            print("while loop killed")
            break

    print("Number of bars: " + str(no_bars))
    print("A: " + str(A*304.8/1000))
    print("x1: " + str(round(x1*304.8)/1000))
    print("#"*45)
    ##### Build ############################################################################
    if lapOrientation == "H":
        adjValue = barDia*0.45
        plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, locPoint + XYZ(0,0,Yoffset))
        p1 = XYZ(polar_to_car(radius-adjValue,0)[0],polar_to_car(radius-adjValue,0)[1],Yoffset)
        p2 = XYZ(polar_to_car(radius,(A/(2*radius)))[0],polar_to_car(radius,(A/(2*radius)))[1],Yoffset)
        p3 = XYZ(polar_to_car(radius+adjValue,A/radius)[0],polar_to_car(radius+adjValue,A/radius)[1],Yoffset)
        curve = [Arc.Create(p1,p3,p2)]
        #build construction bar
        rebarCur = Structure.Rebar.CreateFromCurvesAndShape(doc, sc_65, bar_type, None, None, WTF, XYZ.BasisZ, curve, RebarHookOrientation.Left, RebarHookOrientation.Left)
        # set construction bar properties
        rebarCur.LookupParameter("A").Set(A)
        rebarCur.LookupParameter("Rebar r Custom").Set(r)
        rebarCur.LookupParameter("Mark").Set("PLINTH CONCENTRIC")
        
        #build radial array
        RotAngle = 360*math.pi/180
        if no_bars > 2:
            elem = RadialArray.ArrayElementWithoutAssociation(doc, view, rebarCur.Id, no_bars, Line.CreateBound(p_0,XYZ.BasisZ), RotAngle, ArrayAnchorMember.Last)
            for elem in elem:
                doc.GetElement(elem).LookupParameter("Mark").Set("PLINTH CONCENTRIC")
                doc.GetElement(elem).LookupParameter("Schedule Mark").Set(bar_mark)
                doc.GetElement(elem).LookupParameter("Rebar r Custom").Set(r)
                rebarCur_rotate = ElementTransformUtils.RotateElement(doc, elem, Line.CreateBound(p_0,p_0+XYZ.BasisZ), spliceRot)
        else:
            doc.Delete(rebarCur.Id)
            twobarRebarCur = Structure.Rebar.CreateFromCurvesAndShape(doc, sc_65, bar_type, None, None, WTF, XYZ.BasisZ, curve, RebarHookOrientation.Left, RebarHookOrientation.Left)
            rebarCopy = ElementTransformUtils.CopyElement(doc, twobarRebarCur.Id, XYZ(0, 0, 0))
            elem = ElementTransformUtils.RotateElement(doc, rebarCopy[0], Line.CreateBound(p_0,XYZ.BasisZ), RotAngle/2)
            doc.GetElement(rebarCopy[0]).LookupParameter("Mark").Set("PLINTH CONCENTRIC")
            doc.GetElement(rebarCopy[0]).LookupParameter("Schedule Mark").Set(bar_mark) 
            doc.GetElement(rebarCopy[0]).LookupParameter("Rebar r Custom").Set(r) 
            twobarRebarCur.LookupParameter("Mark").Set("PLINTH CONCENTRIC")
            twobarRebarCur.LookupParameter("Schedule Mark").Set(bar_mark) 
            twobarRebarCur.LookupParameter("Rebar r Custom").Set(r) 
            
    elif lapOrientation == "V":
        # draw a  line  to place the bar 
        preplane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, locPoint + XYZ(0,0,Yoffset))
        #plane = Plane.CreateByThreePoints(p_1, XYZ(), p_3)
        precurve = [Arc.Create(preplane, radius, 0, A/radius)]
        adjValue = (barDia)*(0.45*lap_length/A)
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
        rebarCur.LookupParameter("A").Set(A)
        rebarCur.LookupParameter("Rebar r Custom").Set(r)
        rebarCur.LookupParameter("Mark").Set("PLINTH CONCENTRIC")
        #build radial array
        RotAngle = 360*math.pi/180
        if no_bars > 2:
            elem = RadialArray.ArrayElementWithoutAssociation(doc, view, rebarCur.Id, no_bars, Line.CreateBound(p_0,XYZ.BasisZ), RotAngle, ArrayAnchorMember.Last)
            for elem in elem:
                doc.GetElement(elem).LookupParameter("Mark").Set("PLINTH CONCENTRIC")
                doc.GetElement(elem).LookupParameter("Schedule Mark").Set(bar_mark) 
                doc.GetElement(elem).LookupParameter("Rebar r Custom").Set(r) 
                rebarCur_rotate = ElementTransformUtils.RotateElement(doc, elem, Line.CreateBound(p_0,p_0+XYZ.BasisZ), spliceRot)
        else:
            doc.Delete(rebarCur.Id)
            twobarRebarCur = Structure.Rebar.CreateFromCurvesAndShape(doc, sc_65, bar_type, None, None, WTF, XYZ.BasisZ, precurve, RebarHookOrientation.Left, RebarHookOrientation.Left)
            rebarCopy = ElementTransformUtils.CopyElement(doc, twobarRebarCur.Id, XYZ(0, 0, barDia))
            elem = ElementTransformUtils.RotateElement(doc, rebarCopy[0], Line.CreateBound(p_0,XYZ.BasisZ), RotAngle/2)
            doc.GetElement(rebarCopy[0]).LookupParameter("Mark").Set("PLINTH CONCENTRIC")
            doc.GetElement(rebarCopy[0]).LookupParameter("Schedule Mark").Set(bar_mark) 
            doc.GetElement(rebarCopy[0]).LookupParameter("Rebar r Custom").Set(r) 
            twobarRebarCur.LookupParameter("Mark").Set("PLINTH CONCENTRIC")
            twobarRebarCur.LookupParameter("Schedule Mark").Set(bar_mark) 
            twobarRebarCur.LookupParameter("Rebar r Custom").Set(r) 
        
    spliceRot = spliceRot + (2*lap_length/radius)
    print("#"*50)


# supress warnings ####################################################################
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SupressWarnings())
t.SetFailureHandlingOptions(failHandler)

t.Commit()


