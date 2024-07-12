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

def polar_to_car(radius, angle_radians):
    x = radius * math.cos(angle_radians)
    y = radius * math.sin(angle_radians)
    return x, y

#FPath =  "C:\Users\Wagner.Human\Desktop\Wolf_RebarData_RevD_V162r5.xlsx" 
FPath = forms.pick_file(file_ext='xlsx', multi_file=False, unc_paths=False)

excel = Excel.ApplicationClass()
excel.Visible = False
workbook = excel.Workbooks.Open(FPath)
xl = workbook.Worksheets['A']


# get all inputs form excel for BC in a dictionary
barDict = {}
for i in range(1,200):
    i += 1
    if "BC" in str(xl.Cells(i, 1).Value2).replace(" ","") :
        bar_mark = str(xl.Cells(i,1).Value2)
        startRadius = float(xl.Cells(i,2).Value2)/304.8
        endRadius = float(xl.Cells(i,3).Value2)/304.8
        bar_size = "Y" + str(xl.Cells(i,4).Value2)[1:3]
        bar_dia = int(xl.Cells(i,4).Value2[1:3])/304.8
        spacing = int(xl.Cells(i,7).Value2)/304.8
        
        bar_parameters = {
            "bar_mark": bar_mark,
            "startRadius": startRadius,
            "endRadius": endRadius,
            "bar_size": bar_size,
            "bar_dia": bar_dia,
            "spacing": spacing
        }
        
        barDict[bar_mark] = bar_parameters


LapConst = 45

#####Rebar Shape ################################################################

rebar_shape = FilteredElementCollector(doc).OfClass(RebarShape).WhereElementIsElementType().ToElements()  
for r_shape in rebar_shape:
    if r_shape.LookupParameter("Type Name").AsString() == '65':
        sc_65 = r_shape

##### Element Host ###############################################################

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
hCone = WTF.LookupParameter("hCone").AsDouble()
slabSlope = (h_plinth - h_base)/(r_base - r_plinth)


#Start a transaction ############################################################



t = Transaction(doc, 'Reinforce')
t.Start()
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

# set bar WTF location to origin
WTF.Location.Point = XYZ(0,0,0)
locPoint = WTF.Location.Point
##### Build single concentric bar at start radius############################################################################
startRadius = barDict["BC100"]["startRadius"]
bar_size = barDict["BC100"]["bar_size"]
radius = startRadius

def concentric_bar(con_barmark, radius, bar_size, top_cover = 40/304.8, bot_cover = 50/304.8, lap_length = 45*barDict["BC100"]["bar_dia"]):

    for  rebar_type in all_rebar_types:
        rebar_name = rebar_type.get_Parameter(BuiltInParameter \
            .SYMBOL_NAME_PARAM).AsString()
        if rebar_name == bar_size:
            bar_type = rebar_type
            break
    
    #work out range of bars

    Yoffset = bot_cover + barDia + 32/304.8

    # start number of bars 
    no_barsStart = 2
    Acon = ((math.pi*(radius -barDia/2)*2)/no_barsStart)+lap_length    
    x1con = radius - (radius*math.cos(Acon/(2*radius)))
    whilekill = 0
    while Acon > 13000/304.8 or x1con > 2500/304.8:
        no_barsStart += 1
        no_bars = no_barsStart
        Acon = ((math.pi*(radius -barDia/2)*2)/no_bars)+lap_length
        Acon = (round((Acon*304.8)/100)*100)/304.8
        x1con = radius - (radius*math.cos(Acon/(2*radius)))
        whilekill += 1
        print('No_barsStart: ' + str(no_barsStart))
        if whilekill > 100:
            print("while loop killed")
            break
    print("Acon: " + str(Acon*304.8/1000))
    print("x1con: " + str(round(x1con*304.8)/1000))
    print("no_bars: " + str(no_bars))
    print("#"*45)
    # draw a  line  to place the bar 
    planeOrigin = locPoint + XYZ(0,0,Yoffset)
    plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, planeOrigin)
    #plane = Plane.CreateByThreePoints(p_1, XYZ(), p_3)
    precurve = [Arc.Create(plane, radius, 0, Acon/radius)]
    adjValue = barDia*0.5

    p1 = precurve[0].GetEndPoint(0) 
    p2 = XYZ(math.cos(Acon/(radius*2))*radius, math.sin(Acon/(radius*2))*radius, Yoffset)
    p3 = precurve[0].GetEndPoint(1)

    p1Vec = p1 - planeOrigin
    p3Vec = p3 - planeOrigin
    p1adj = p1 + p1Vec.Normalize()*adjValue
    p3adj = p3 - p3Vec.Normalize()*adjValue

    curve = [Arc.Create(p1adj , p2 ,  p3adj)]

    #build construction bar
    rebarCur = Structure.Rebar.CreateFromCurvesAndShape(doc, sc_65, bar_type, None, None, WTF, XYZ.BasisZ, curve, RebarHookOrientation.Left, RebarHookOrientation.Left)

    #get barmark
    tot_ss = 1
    if float(endRadius-startRadius) > float(barDict["BC100"]["spacing"]*5) :
        tot_ss = math.floor((endRadius-startRadius)/(spacing*5))

    if con_barmark == None:
        if Acon > 13000/304.8 or x1con > 2500/304.8 or tot_ss > 1:
            con_barmark = "BC101"
        else:
            con_barmark = "BC100"


    # set construction bar properties
    rebarCur.LookupParameter("A").Set(Acon)
    rebarCur.LookupParameter("r").Set(radius)
    rebarCur.LookupParameter("Mark").Set("BOTTOM CONCENTRIC")

    #build radial array
    RotAngle = 360*math.pi/180
    if no_bars > 2:
        elem = RadialArray.ArrayElementWithoutAssociation(doc, view, rebarCur.Id, no_bars, Line.CreateBound(locPoint,XYZ.BasisZ), RotAngle, ArrayAnchorMember.Last)
        for elem in elem:
            doc.GetElement(elem).LookupParameter("Mark").Set('BOTTOM CONCENTRIC')
            doc.GetElement(elem).LookupParameter("Schedule Mark").Set(con_barmark) 
            # rebarCur_rotate = ElementTransformUtils.RotateElement(doc, elem, Line.CreateBound(locPoint,locPoint+XYZ.BasisZ), spliceRot)
    else:
        doc.Delete(rebarCur.Id)
        twobarRebarCur = Structure.Rebar.CreateFromCurvesAndShape(doc, sc_65, bar_type, None, None, WTF, XYZ.BasisZ, precurve, RebarHookOrientation.Left, RebarHookOrientation.Left)
        rebarCopy = ElementTransformUtils.CopyElement(doc, twobarRebarCur.Id, XYZ(0, 0, barDia))
        elem = ElementTransformUtils.RotateElement(doc, rebarCopy[0], Line.CreateBound(locPoint,XYZ.BasisZ), RotAngle/2)
        doc.GetElement(rebarCopy[0]).LookupParameter("Mark").Set('BOTTOM CONCENTRIC')
        doc.GetElement(rebarCopy[0]).LookupParameter("Schedule Mark").Set(con_barmark)                
    radius -= spacing

    failHandler = t.GetFailureHandlingOptions()
    failHandler.SetFailuresPreprocessor(SupressWarnings())
    t.SetFailureHandlingOptions(failHandler)

#######################################################################################################################################################
startRadius = barDict["BC100"]["startRadius"] 
lap_length = barDict["BC100"]["bar_dia"]*LapConst
top_cover =  40/304.8
bot_cover =  50/304.8
Yoffset  =  bot_cover + barDict["BC100"]["bar_dia"]/2 + 32/304.8

A = 13500/304.8  
x1 = startRadius - (startRadius*math.cos(A/(2*startRadius)))
while A > 13000/304.8 or x1 > 2500/304.8:
    A = A - 100/304.8
    x1 = startRadius - (startRadius*math.cos(A/(2*startRadius)))
print(" start A: " + str(A*304.8/1000))
print("start x1: " + str(round(x1*304.8)/1000))
print("startRad:  " + str(startRadius*304.8))
r1 = startRadius +barDict["BC100"]["bar_dia"]
r3 = ((startRadius + math.sqrt(startRadius**2 + 4*(spacing/(2*math.pi))*A))/2)-barDict["BC100"]["bar_dia"]
r2 = (r1+r3)/2
r3spl = (startRadius + math.sqrt(startRadius**2 + 4*(spacing/(2*math.pi))*(A-lap_length)))/2
theta1 = 0
theta3 = A/r2  + theta1
theta2 = (theta3-theta1)/2 + theta1
theta3spl = (A-lap_length)/r2

p1 = XYZ(polar_to_car(r1,theta1)[0],polar_to_car(r1,theta1)[1],Yoffset)
p2 = XYZ(polar_to_car(r2,theta2)[0],polar_to_car(r2,theta2)[1],Yoffset)
p3 = XYZ(polar_to_car(r3,theta3)[0],polar_to_car(r3,theta3)[1],Yoffset)
preCurve = Arc.Create(p1,p3,p2)
bar_size = barDict["BC100"]["bar_size"]
for  rebar_type in all_rebar_types:
    rebar_name = rebar_type.get_Parameter(BuiltInParameter \
        .SYMBOL_NAME_PARAM).AsString()
    if rebar_name == bar_size:
        bar_type = rebar_type
        break
rebarCur1 = Structure.Rebar.CreateFromCurvesAndShape(doc, sc_65, bar_type, None, None, WTF, XYZ.BasisZ, [preCurve], RebarHookOrientation.Left, RebarHookOrientation.Left)
            
rebarCur1.LookupParameter("Mark").Set('BOTTOM CONCENTRIC')


for i in range(len(barDict.items())-1):
    i += 1
    bar_mark = barDict["BC"+str(i)+"00"]["bar_mark"]
    startRadius = barDict["BC"+str(i)+"00"]["startRadius"]
    endRadius = barDict["BC"+str(i)+"00"]["endRadius"]
    bar_size = barDict["BC"+str(i)+"00"]["bar_size"]
    bar_dia = barDict["BC"+str(i)+"00"]["bar_dia"]
    spacing = barDict["BC"+str(i)+"00"]["spacing"]
    
    if endRadius > r_base:
        endRadius = r_base-bot_cover-barDia*2.5
        print("Adjusted endRadius: " + str(endRadius*304.8))
        
    for  rebar_type in all_rebar_types:
        rebar_name = rebar_type.get_Parameter(BuiltInParameter \
            .SYMBOL_NAME_PARAM).AsString()
        if rebar_name == bar_size:
            bar_type = rebar_type
            break

    barDia = bar_type.LookupParameter("Bar Diameter").AsDouble()

    #### Calculate from input parameters ############################################
    lap_length = LapConst*bar_dia
    Yoffset  =  bot_cover + barDia/2 + 32/304.8
    # print(str('radius') + "  ---  \t " + bar_mark + "  --- \t \t " + str('##') + "  ---\t \t " + bar_size + "  ---\t \t " + str('##') + "  ---\t \t " + str('spacing') + "  ---\t \t " + str(spacing))
    print("*"*45)  

    print("bar_mark: " + str(bar_mark))
    print("startRadius: " + str(startRadius*304.8))
    print("endRadius: " + str(endRadius*304.8))
    print("bar_size: " + str(bar_size))
    print("bar_dia: " + str(bar_dia*304.8))
    print("spacing: " + str(spacing*304.8))
    print("lap_length: " + str(lap_length*304.8))
    print("*"*45)


    if startRadius > 6000/304.8 and barDia < 32/304.8:
        A = 13000/304.8

    tot_subset = 1
    if endRadius-startRadius > spacing*3:
        tot_subset = math.floor((endRadius-startRadius)/(spacing*3))
        print("subset: " + str(tot_subset))
    subset = 1
    subsetRange = (endRadius - startRadius)/tot_subset

    if A < 13000/304.8:
        while r2 < (endRadius-bot_cover-(barDia*1.5)):
            while r2 < startRadius+subsetRange*subset:
                r1 = r3spl-(bar_dia/2)
                r3 = (r1 + math.sqrt(r1**2 + 4*(spacing/(2*math.pi))*(A-lap_length)))/2
                r2 = (r1+r3)/2
                r3spl = ((r1 + math.sqrt(r1**2 + 4*(spacing/(2*math.pi))*(A-lap_length)))+bar_dia)/2
                theta1 = theta3spl
                theta3 = theta1 + A/r2
                theta2 = (theta3-theta1)/2 +theta1 #+ A/(r2*2)
                theta3spl = (A-lap_length)/r2 + theta1
                #build curve list
                p1 = XYZ(polar_to_car(r1,theta1)[0],polar_to_car(r1,theta1)[1],Yoffset)
                p2 = XYZ(polar_to_car(r2,theta2)[0],polar_to_car(r2,theta2)[1],Yoffset)
                p3 = XYZ(polar_to_car(r3,theta3)[0],polar_to_car(r3,theta3)[1],Yoffset)
                preCurve = Arc.Create(p1,p3,p2)
                #build bar
                rebarCur = Structure.Rebar.CreateFromCurvesAndShape(doc, sc_65, bar_type, None, None, WTF, XYZ.BasisZ, [preCurve], RebarHookOrientation.Left, RebarHookOrientation.Left)
                # set mark and schedule mark
                rebarCur.LookupParameter("Mark").Set('BOTTOM CONCENTRIC')
                rebarCur.LookupParameter("Schedule Mark").Set(str(bar_mark)[:-1]+str(subset))
            
            subset += 1
            A_live = (theta3 - theta1)*r2
            x1_live = r2 - (r2*math.cos(A/(2*r2)))
            while A_live < 13000/304.8 and x1_live < 2500/304.8:
                A = A + 100/304.8
                x1 = r2 - (r2*math.cos(A/(2*r2)))
                A_live = A
                x1_live = x1
            if r2 > 6000/304.8 and barDia < 25/304.8:
                A = 13000/304.8

            # print(" start A: " + str(A*304.8/1000))
            # print("start x1: " + str(round(x1*304.8)/1000))
            # print("subset: " + str(subset))

    elif A > 12999/304.8 and A < 13001/304.8:
        print(" in 000 loop " + str(bar_mark))
        while r2 < (endRadius-bot_cover-(barDia*1.5)):
            r1 = r3spl-(bar_dia/2)
            r3 = (r1 + math.sqrt(r1**2 + 4*(spacing/(2*math.pi))*(A-lap_length)))/2
            r2 = (r1+r3)/2
            r3spl = ((r1 + math.sqrt(r1**2 + 4*(spacing/(2*math.pi))*(A-lap_length)))+bar_dia)/2
            theta1 = theta3spl
            theta3 = theta1 + A/r2
            theta2 = (theta3-theta1)/2 +theta1 #+ A/(r2*2)
            theta3spl = (A-lap_length)/r2 + theta1
            #build curve list
            p1 = XYZ(polar_to_car(r1,theta1)[0],polar_to_car(r1,theta1)[1],Yoffset)
            p2 = XYZ(polar_to_car(r2,theta2)[0],polar_to_car(r2,theta2)[1],Yoffset)
            p3 = XYZ(polar_to_car(r3,theta3)[0],polar_to_car(r3,theta3)[1],Yoffset)
            preCurve = Arc.Create(p1,p3,p2)

            #build bar
            rebarCur = Structure.Rebar.CreateFromCurvesAndShape(doc, sc_65, bar_type, None, None, WTF, XYZ.BasisZ, [preCurve], RebarHookOrientation.Left, RebarHookOrientation.Left)
            # set mark and schedule mark
            rebarCur.LookupParameter("Mark").Set('BOTTOM CONCENTRIC')
            rebarCur.LookupParameter("Schedule Mark").Set(bar_mark)
            
        
        subset += 1
        A_live = (theta3 - theta1)*r2
        x1_live = r2 - (r2*math.cos(A/(2*r2)))
        while A < 13000/304.8 and x1 < 2500/304.8:
            A = A + 100/304.8
            x1 = r2 - (r2*math.cos(A/(2*r2)))
        print(" start A: " + str(A*304.8/1000))
        print("start x1: " + str(round(x1*304.8)/1000))
        print("subset: " + str(subset))

outConDia = []
for bar_mark in barDict:
    outConDia.append(str(bar_mark)[2:])
outConDia = max(outConDia)
barDia = barDict["BC"+str(outConDia)]["bar_dia"]
outConRadius = r_base-bot_cover-barDia*0.5
print('#'*100)
print("outConDia: " + str(outConDia))
print("outConRadius: " + str(outConRadius*304.8))
print('#'*100)
startRadius = outConRadius

outCon = concentric_bar("BC"+str(outConDia), outConRadius, barDict["BC"+str(outConDia)]["bar_size"], top_cover = 40/304.8, bot_cover = 50/304.8, lap_length = 45*barDia)

conbars = concentric_bar(None ,barDict["BC100"]["startRadius"], barDict["BC100"]["bar_size"])
#conbars.LookupParameter("A").Set(rebarCur.LookupParameter("A").AsDouble())
doc.Delete(rebarCur1.Id)



t.Commit()        



FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
scheduleMarkList = []
for elem in FEC:
    if elem.LookupParameter("Schedule Mark").AsString() not in scheduleMarkList:
        scheduleMarkList.append(elem.LookupParameter("Schedule Mark").AsString())



t = Transaction(doc, "Update r")
t.Start()
for i in range(len(scheduleMarkList)):
    if "BC" in scheduleMarkList[i]:
        sum_r = 0
        count_r = 0
        for elem in FEC:
            if elem.LookupParameter("Schedule Mark").AsString() == scheduleMarkList[i]:
                sum_r += elem.LookupParameter("r").AsDouble()
                count_r += 1
        r = round((sum_r/count_r)*100)/100
        r = round(r*304.8)/304.8
        print(scheduleMarkList[i])
        print(r*304.8)
        for elem in FEC:
            if elem.LookupParameter("Schedule Mark").AsString() == scheduleMarkList[i]:
                elem.LookupParameter("Rebar r Custom").Set(r)

        print("updated r")


t.Commit()

t = Transaction(doc, "Update A")    
t.Start()
for i in range(len(scheduleMarkList)):
    if "BC" in scheduleMarkList[i]:
        sum_A = 0
        count_A = 0
        A_max = 0
        for elem in FEC:
            if elem.LookupParameter("Schedule Mark").AsString() == scheduleMarkList[i]:
                sum_A += elem.LookupParameter("A").AsDouble()
                count_A += 1
                if elem.LookupParameter("A").AsDouble() > A_max:
                    A_max = elem.LookupParameter("A").AsDouble()
        A = round((sum_A/count_A)*100)/100
        A = round(A*304.8)/304.8
        print(scheduleMarkList[i])
        print(A*304.8)
        print(A_max*304.8)

        for elem in FEC:
            if elem.LookupParameter("Schedule Mark").AsString() == scheduleMarkList[i]:
                elem.LookupParameter("A").Set(((round((A_max*304.8)/100)*100)/304.8))



t.Commit()

print(barDict)

# supress warnings ####################################################################
failHandler = t.GetFailureHandlingOptions()
failHandler.SetFailuresPreprocessor(SupressWarnings())
t.SetFailureHandlingOptions(failHandler)





