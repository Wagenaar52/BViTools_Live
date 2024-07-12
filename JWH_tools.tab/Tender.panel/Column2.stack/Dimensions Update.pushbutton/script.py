import clr
import math
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import forms

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document


# FPath = forms.pick_file(file_ext='xlsx', multi_file=False, unc_paths=False)


# #get hould of excel file using ironpython
# clr.AddReference("Microsoft.Office.Interop.Excel")
# import Microsoft.Office.Interop.Excel as Excel
# excel = Excel.ApplicationClass()
# excel.Visible = False
# workbook = excel.Workbooks.Open(FPath)
# xl = workbook.Worksheets['SheetDATA']

# Get hold of sheets and tables in document to populate
notes = doc.GetElement(ElementId(1056082))


# Get Elements from document
wtf_steel = doc.GetElement(ElementId(1087568))
grout = doc.GetElement(ElementId(1804084))
blinding = doc.GetElement(ElementId(1097087))
backfill = doc.GetElement(ElementId(2342235))
anchor_cage = doc.GetElement(ElementId(1322206))

rTower = wtf_steel.LookupParameter('rTower').AsDouble()
hPlinth = wtf_steel.LookupParameter('hPlinth').AsDouble()
rPlinth = wtf_steel.LookupParameter('rPlinth').AsDouble()
hPit = wtf_steel.LookupParameter('hBottomVoid').AsDouble()
rBoltIN = anchor_cage.LookupParameter('rBoltInner').AsDouble()
rBoltOUT = anchor_cage.LookupParameter('rBoltOuter').AsDouble()

towerFlangeIN  = anchor_cage.LookupParameter('rTower').AsDouble() - anchor_cage.LookupParameter('wFlange').AsDouble()/2
towerFlangeOUT = anchor_cage.LookupParameter('rTower').AsDouble() + anchor_cage.LookupParameter('wFlange').AsDouble()/2
bearingPlateIN = anchor_cage.LookupParameter('rTower').AsDouble() - anchor_cage.LookupParameter('WBearingPlate').AsDouble()/2
bearingPlateOUT = anchor_cage.LookupParameter('rTower').AsDouble() + anchor_cage.LookupParameter('WBearingPlate').AsDouble()/2
groutIN = anchor_cage.LookupParameter('rTower').AsDouble() - grout.LookupParameter('wGroutTop').AsDouble()/2
groutOUT = anchor_cage.LookupParameter('rTower').AsDouble() + grout.LookupParameter('wGroutTop').AsDouble()/2


t = Transaction(doc)
t.Start("Dimension Center Line")

#Top line end points
#Left line end points
topLplinth = XYZ(-rPlinth, 0, hPlinth+3)
topL_towerFlangeIN = XYZ(-towerFlangeIN, 0, hPlinth+1)
topL_towerFlangeOUT = XYZ(-towerFlangeOUT, 0, hPlinth+3)
topL_bearingPlateIN = XYZ(-bearingPlateIN, 0, hPlinth+1)
topL_bearingPlateOUT = XYZ(-bearingPlateOUT, 0, hPlinth+3)
topL_groutIN = XYZ(-groutIN, 0, hPlinth+1)
topL_groutOUT = XYZ(-groutOUT, 0, hPlinth+3)
#Right line end points
topRplinth = XYZ(rPlinth, 0, hPlinth+3)
topR_towerFlangeIN = XYZ(towerFlangeIN, 0, hPlinth+1)
topR_towerFlangeOUT = XYZ(towerFlangeOUT, 0, hPlinth+3)
topR_bearingPlateIN = XYZ(bearingPlateIN, 0, hPlinth+1)
topR_bearingPlateOUT = XYZ(bearingPlateOUT, 0, hPlinth+3)
topR_groutIN = XYZ(groutIN, 0, hPlinth+1)
topR_groutOUT = XYZ(groutOUT, 0, hPlinth+3)

dim_plinth = Line.CreateBound(topLplinth+XYZ(0,0,4.2177), topRplinth+XYZ(0,0,4.2177))
dim_towerFlangeIN = Line.CreateBound(topL_towerFlangeIN+XYZ(0,0,1.6246), topR_towerFlangeIN+XYZ(0,0,1.6246))
dim_bearingPlateIN = Line.CreateBound(topL_bearingPlateIN+XYZ(0,0,0.9685), topR_bearingPlateIN+XYZ(0,0,0.9685))
dim_groutIN = Line.CreateBound(topL_groutIN+XYZ(0,0,0.3123), topR_groutIN+XYZ(0,0,0.3123))
dim_towerFlangeOUT = Line.CreateBound(topL_towerFlangeOUT+XYZ(0,0,2.2493), topR_towerFlangeOUT+XYZ(0,0,2.2493))
dim_bearingPlateOUT = Line.CreateBound(topL_bearingPlateOUT+XYZ(0,0,2.9055), topR_bearingPlateOUT+XYZ(0,0,2.9055))
dim_groutOUT = Line.CreateBound(topL_groutOUT+XYZ(0,0,3.5616), topR_groutOUT+XYZ(0,0,3.5616))

#Bottom line end points
#Left line end points
botLplinth = XYZ(-rPlinth, 0, hPlinth+2.9)
botL_towerFlangeIN = XYZ(-towerFlangeIN, 0, hPlinth+0.9)
botL_towerFlangeOUT = XYZ(-towerFlangeOUT, 0, hPlinth+2.9)
botL_bearingPlateIN = XYZ(-bearingPlateIN, 0, hPlinth+0.9)
botL_bearingPlateOUT = XYZ(-bearingPlateOUT, 0, hPlinth+2.9)
botL_groutIN = XYZ(-groutIN, 0, hPlinth+0.9)
botL_groutOUT = XYZ(-groutOUT, 0, hPlinth+2.9)
#Right line end points
botRplinth = XYZ(rPlinth, 0, hPlinth+2.9)
botR_towerFlangeIN = XYZ(towerFlangeIN, 0, hPlinth+0.9)
botR_towerFlangeOUT = XYZ(towerFlangeOUT, 0, hPlinth+2.9)
botR_bearingPlateIN = XYZ(bearingPlateIN, 0, hPlinth+0.9)
botR_bearingPlateOUT = XYZ(bearingPlateOUT, 0, hPlinth+2.9)
botR_groutIN = XYZ(groutIN, 0, hPlinth+0.9)
botR_groutOUT = XYZ(groutOUT, 0, hPlinth+2.9)



#Center line end points
topLCL_BoltIN =     XYZ(-rBoltIN,  0, hPlinth+1.3)
topLCL_BoltOUT =    XYZ(-rBoltOUT, 0, hPlinth+1.3)
topLCL_Tower =      XYZ(-rTower,   0, hPlinth+2)
topRCL_BoltIN =     XYZ(rBoltIN, 0, hPlinth+1.3)
topRCL_BoltOUT =    XYZ(rBoltOUT, 0, hPlinth+1.3)
topRCL_Tower =      XYZ(rTower,  0, hPlinth+2)
botLCL_BoltIN =     XYZ(-rBoltIN,  0, -hPit+100/304.8)
botLCL_BoltOUT =    XYZ(-rBoltOUT, 0, -hPit+100/304.8)
botLCL_Tower =      XYZ(-rTower,   0, -hPit+50/304.8)
botRCL_BoltIN =     XYZ(rBoltIN, 0, -hPit+100/304.8)
botRCL_BoltOUT =    XYZ(rBoltOUT, 0, -hPit+100/304.8)
botRCL_Tower =      XYZ(rTower, 0, -hPit+50/304.8)


#Center line curves
line1 = Line.CreateBound(topLCL_BoltIN,     botLCL_BoltIN)
line2 = Line.CreateBound(topLCL_BoltOUT,    botLCL_BoltOUT)
line3 = Line.CreateBound(topLCL_Tower,      botLCL_Tower)
line4 = Line.CreateBound(topRCL_BoltIN,     botRCL_BoltIN)
line5 = Line.CreateBound(topRCL_BoltOUT,    botRCL_BoltOUT)
line6 = Line.CreateBound(topRCL_Tower,      botRCL_Tower)

# dimention line curves
L_Plinth = Line.CreateBound(topLplinth, botLplinth)
L_towerFlangeIN = Line.CreateBound(topL_towerFlangeIN, botL_towerFlangeIN)
L_towerFlangeOUT = Line.CreateBound(topL_towerFlangeOUT, botL_towerFlangeOUT)
L_bearingPlateIN = Line.CreateBound(topL_bearingPlateIN, botL_bearingPlateIN)
L_bearingPlateOUT = Line.CreateBound(topL_bearingPlateOUT, botL_bearingPlateOUT)
L_groutIN = Line.CreateBound(topL_groutIN, botL_groutIN)
L_groutOUT = Line.CreateBound(topL_groutOUT, botL_groutOUT)
R_Plinth = Line.CreateBound(topRplinth, botRplinth)
R_towerFlangeIN = Line.CreateBound(topR_towerFlangeIN, botR_towerFlangeIN)
R_towerFlangeOUT = Line.CreateBound(topR_towerFlangeOUT, botR_towerFlangeOUT)
R_bearingPlateIN = Line.CreateBound(topR_bearingPlateIN, botR_bearingPlateIN)
R_bearingPlateOUT = Line.CreateBound(topR_bearingPlateOUT, botR_bearingPlateOUT)
R_groutIN = Line.CreateBound(topR_groutIN, botR_groutIN)
R_groutOUT = Line.CreateBound(topR_groutOUT, botR_groutOUT)
# Create dimensions between top LCL lines and  top RCL lines
dimCL_BoltIN = Line.CreateBound(XYZ(topLCL_BoltIN.X,0,hPlinth+1000/304.8),       XYZ(topRCL_BoltIN.X,0,hPlinth+1000/304.8))
dimCL_Tower = Line.CreateBound(XYZ(topLCL_BoltOUT.X,0,hPlinth+1200/304.8),      XYZ(topRCL_BoltOUT.X,0,hPlinth+1200/304.8))
dimCL_BoltOUT = Line.CreateBound(XYZ(topLCL_Tower.X,0,hPlinth+1400/304.8),        XYZ(topRCL_Tower.X,0,hPlinth+1400/304.8))

LeftlineList = [L_Plinth, L_towerFlangeIN, L_towerFlangeOUT, L_bearingPlateIN, L_bearingPlateOUT, L_groutIN, L_groutOUT,line1, line3, line2]
RightlineList = [R_Plinth, R_towerFlangeIN, R_towerFlangeOUT, R_bearingPlateIN, R_bearingPlateOUT, R_groutIN, R_groutOUT, line4, line6, line5]
dimLineList = [dim_plinth, dim_towerFlangeIN, dim_towerFlangeOUT, dim_bearingPlateIN, dim_bearingPlateOUT, dim_groutIN, dim_groutOUT, dimCL_BoltIN, dimCL_Tower, dimCL_BoltOUT]

lineList = [line1, line2, line3, line4, line5, line6]
SectionAA = doc.GetElement(ElementId(453188))
DetailView1 = doc.GetElement(ElementId(522437))
DetailView2 = doc.GetElement(ElementId(522453))
viewList = [SectionAA, DetailView1, DetailView2]
dimList =[]

for i in range(len(LeftlineList)):
    LeftLine = doc.Create.NewDetailCurve(SectionAA, LeftlineList[i])
    RightLine = doc.Create.NewDetailCurve(SectionAA, RightlineList[i])
    refArray = ReferenceArray()
    refArray.Append(LeftLine.GeometryCurve.Reference)
    refArray.Append(RightLine.GeometryCurve.Reference)
    dim = doc.Create.NewDimension(SectionAA, dimLineList[i], refArray, doc.GetElement(ElementId(1018370)))
    dim.TextPosition = dim.TextPosition + XYZ(0.5, 0, -0.1)


for view in viewList:
    for line in lineList:
        dline = doc.Create.NewDetailCurve(view, line)
        dline.LineStyle = doc.GetElement(ElementId(1018897))



t.Commit()

# t = Transaction(doc)
# t.Start("Update leaders in plan view")


# txt1 = doc.GetElement(ElementId(1131183))       
# txt2 = doc.GetElement(ElementId(1131197))

# dim1 = doc.GetElement(ElementId(1116067))
# dim2 = doc.GetElement(ElementId(1116072))





# txt1.Coord = XYZ(dim1.Origin.X -4 , dim1.Origin.Y +2 , 0)
# txt2.Coord = XYZ(dim1.Origin.X -4 , dim2.Origin.Y +2 , 0)

# txt1.GetLeaders()[0].Elbow = XYZ(dim1.Origin.X -3 , txt1.Coord.Y , 0)
# txt2.GetLeaders()[0].Elbow = XYZ(dim2.Origin.X -3 , txt2.Coord.Y , 0)

# txt1.GetLeaders()[0].End = dim1.Origin
# txt2.GetLeaders()[0].End = dim2.Origin

# t.Commit()





















# t = Transaction(doc)
# t.Start("move line boundingbox")


# rTower = wtf_steel.LookupParameter('rTower').AsDouble()
# hPlinth = wtf_steel.LookupParameter('hPlinth').AsDouble()
# hPit = wtf_steel.LookupParameter('hBottomVoid').AsDouble()
# rBoltIN = anchor_cage.LookupParameter('rBoltInner').AsDouble()
# rBoltOUT = anchor_cage.LookupParameter('rBoltOuter').AsDouble()


# #Center line end points
# topLCL_BoltIN =     XYZ(-rBoltIN,  0, hPlinth+1.3)
# topLCL_BoltOUT =    XYZ(-rBoltOUT, 0, hPlinth+1.3)
# topLCL_Tower =      XYZ(-rTower,   0, hPlinth+2)
# topRCL_BoltIN =     XYZ(rBoltIN, 0, hPlinth+1.3)
# topRCL_BoltOUT =    XYZ(rBoltOUT, 0, hPlinth+1.3)
# topRCL_Tower =      XYZ(rTower,  0, hPlinth+2)
# botLCL_BoltIN =     XYZ(-rBoltIN,  0, -hPit+100/304.8)
# botLCL_BoltOUT =    XYZ(-rBoltOUT, 0, -hPit+100/304.8)
# botLCL_Tower =      XYZ(-rTower,   0, -hPit+50/304.8)
# botRCL_BoltIN =     XYZ(rBoltIN, 0, -hPit+100/304.8)
# botRCL_BoltOUT =    XYZ(rBoltOUT, 0, -hPit+100/304.8)
# botRCL_Tower =      XYZ(rTower, 0, -hPit+50/304.8)
# #Center line curves
# line1 = Line.CreateBound(topLCL_BoltIN,     botLCL_BoltIN)
# line2 = Line.CreateBound(topLCL_BoltOUT,    botLCL_BoltOUT)
# line3 = Line.CreateBound(topLCL_Tower,      botLCL_Tower)
# line4 = Line.CreateBound(topRCL_BoltIN,     botRCL_BoltIN)
# line5 = Line.CreateBound(topRCL_BoltOUT,    botRCL_BoltOUT)
# line6 = Line.CreateBound(topRCL_Tower,      botRCL_Tower)

# lineList = [ line4, line5, line6,line1, line2, line3]
# Detail1= doc.GetElement(ElementId(2751417))
# Detail2= doc.GetElement(ElementId(522437))
# Detail3= doc.GetElement(ElementId(522453))
# DetailList = [Detail1, Detail2, Detail3]

# for view in DetailList:
#     for line in lineList:
#         doc.Create.NewDetailCurve(view, line).LineStyle = doc.GetElement(ElementId(1018897))

#     # Create dimensions between top LCL lines and  top RCL lines
#     dimCL_BoltIN = Line.CreateBound(XYZ(topLCL_BoltIN.X,0,hPlinth+1000/304.8),       XYZ(topRCL_BoltIN.X,0,hPlinth+1000/304.8))
#     dimCL_Tower = Line.CreateBound(XYZ(topLCL_BoltOUT.X,0,hPlinth+1200/304.8),      XYZ(topRCL_BoltOUT.X,0,hPlinth+1200/304.8))
#     dimCL_BoltOUT = Line.CreateBound(XYZ(topLCL_Tower.X,0,hPlinth+1400/304.8),        XYZ(topRCL_Tower.X,0,hPlinth+1400/304.8))


#     t.Commit()
