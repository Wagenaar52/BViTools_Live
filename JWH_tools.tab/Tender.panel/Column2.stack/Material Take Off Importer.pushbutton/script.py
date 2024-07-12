import clr
import math
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import forms


uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

FPath = forms.pick_file(file_ext='xlsx', multi_file=False, unc_paths=False)

#get hould of excel file using ironpython
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
excel = Excel.ApplicationClass()
excel.Visible = False
workbook = excel.Workbooks.Open(FPath)
xl = workbook.Worksheets['SheetDATA']

# Get hold of sheets an tables in document to populate
MTO = doc.GetElement(ElementId(2257920))
# Get Elements from document
wtf_steel = doc.GetElement(ElementId(1087568))
grout = doc.GetElement(ElementId(1804084))
blinding = doc.GetElement(ElementId(1097087))
backfill = doc.GetElement(ElementId(2342235))


# Create variables from excel sheet
GroutClassXL = (xl.Cells(9,16).Value2)
PlinthClassXL =(xl.Cells(10,16).Value2)
PitClassXL = (xl.Cells(11,16).Value2)
BaseClassXL =(xl.Cells(12,16).Value2)
BlindingClassXL =(xl.Cells(13,16).Value2)
RebarRateXL = (xl.Cells(14,16).Value2)
BackfillClassXL =(xl.Cells(15,16).Value2)
ExcavationClassXL = (xl.Cells(16,16).Value2)

# Volume Caculations
meanPitRad = ((wtf_steel.LookupParameter('rVoidOuter').AsDouble()*0.3048)+(wtf_steel.LookupParameter('rVoidInner').AsDouble()*0.3048))/2

concVol = wtf_steel.LookupParameter("Volume").AsDouble()*0.3048**3
GroutVol = grout.LookupParameter("Volume").AsDouble()*0.3048**3
if ((wtf_steel.LookupParameter('hPlinth').AsDouble())-(wtf_steel.LookupParameter('hCone').AsDouble())) < 900/304.8:
    PlinthVol = ((wtf_steel.LookupParameter('rPlinth').AsDouble())*0.3048)**2*math.pi*((wtf_steel.LookupParameter('hPlinth').AsDouble())-(wtf_steel.LookupParameter('hCone').AsDouble()))*0.3048
else:
    PlinthVol = ((wtf_steel.LookupParameter('rPlinth').AsDouble())*0.3048)**2*math.pi*(900/304.8)*0.3048

PitVol = (meanPitRad)**2*math.pi*(wtf_steel.LookupParameter('hBottomVoid').AsDouble()*0.3048)
BaseVol = concVol - PlinthVol - PitVol
BlVol = blinding.LookupParameter("Volume").AsDouble()*0.3048**3
RebarTonage = wtf_steel.LookupParameter("Volume").AsDouble()*0.3048**3*RebarRateXL/1000
BFVol = backfill.LookupParameter("Volume").AsDouble()*0.3048**3
ExVol = BFVol + BlVol + concVol

#Area Calculations

f1_plinth = str(round(wtf_steel.GetMaterialArea(ElementId(1017985), usePaintMaterial = False)*0.092903,0))[:-2]
f1_base = str(round(wtf_steel.GetMaterialArea(ElementId(2096835), usePaintMaterial = False)*0.092903,0))[:-2]


U2_plinth = str(round(((wtf_steel.LookupParameter('rPlinth').AsDouble())*0.3048)**2*math.pi,0))[:-2]   #wtf_steel.GetMaterialArea(ElementId(2096793), usePaintMaterial = False)*0.092903,0

U2_Base = str(round(wtf_steel.GetMaterialArea(ElementId(2096792), usePaintMaterial = False)*0.092903,0))[:-2]


print(f1_base)
print(f1_plinth)
print(U2_plinth)
print(U2_Base)
print("*"*50)

# Set floats to string with two decimal places for parameter assignment
GroutVol = str(round(GroutVol,1))   
PlinthVol = str(round(PlinthVol,1))
PitVol = str(round(PitVol,1))
BaseVol = str(round(BaseVol,1))
BlVol = str(round(BlVol,1))
RebarTonage = str(round(RebarTonage,2))
BFVol = str((round(BFVol,0)))[:-2]
ExVol = str(round(ExVol,0))[:-2]


#  MTO Parameters

t = Transaction(doc)
t.Start("Apply parameter values")

#RebarClass = MTO.LookupParameter("RebarClass")
#BFDens = MTO.LookupParameter("BackfillDens")
# F1 = MTO.LookupParameter("F1 BaseV")
# F2 = MTO.LookupParameter("F2 PlinthV")
# U1 = MTO.LookupParameter("U1 SlabS")
# U2 = MTO.LookupParameter("U2 PlinthH")

MTO.LookupParameter("GroutClass").Set(GroutClassXL)
MTO.LookupParameter("PlinthClass").Set(PlinthClassXL)
MTO.LookupParameter("PitClass").Set(PitClassXL)
MTO.LookupParameter("BaseClass").Set(BaseClassXL)
MTO.LookupParameter("BlindingClass").Set(BlindingClassXL)
MTO.LookupParameter("RebarRate").Set(RebarRateXL)
MTO.LookupParameter("BackfillClass").Set(BackfillClassXL)
MTO.LookupParameter("ExcavationClass").Set(ExcavationClassXL)

MTO.LookupParameter("F1 BaseV").Set(f1_base)
MTO.LookupParameter("F2 PlinthV").Set(f1_plinth)
MTO.LookupParameter("U1 SlabS").Set(U2_Base)
MTO.LookupParameter("U2 PlinthH").Set(U2_plinth)

print(concVol)
print(GroutVol)
print(PlinthVol)
print(PitVol)
print(BaseVol)
print(BlVol)
print(RebarTonage)
print(BFVol)
print(ExVol)

MTO.LookupParameter("GroutVol").Set(GroutVol)
MTO.LookupParameter("PlinthVol").Set(PlinthVol)
MTO.LookupParameter("PitVol").Set(PitVol)
MTO.LookupParameter("BaseVol").Set(BaseVol)
MTO.LookupParameter("BlindingVol").Set(BlVol)
MTO.LookupParameter("RebarTonage").Set(RebarTonage)
MTO.LookupParameter("BackFillVol").Set(BFVol)
MTO.LookupParameter("ExcavationVol").Set(ExVol)

t.Commit()

print("Material Take Off Parameters Updated")
# Close Excel application object
workbook.Close(False)
excel.Quit()
print("Excel Closed")














# sheet = doc.GetElement(ElementId(2106182))
# anchor_cage = doc.GetElement(ElementId(1322206))
# duct = doc.GetElement(ElementId(1897290))



# #Values to be assigned to parameters

# rebar_rate = float(int(xl.Cells(10,2).Value2)/1000)

# SlabH = float(wtf_steel.LookupParameter("hBase").AsValueString())/1000
# ConeH = float(wtf_steel.LookupParameter("hCone").AsValueString())/1000
# PlinthH = float(wtf_steel.LookupParameter("hPlinth").AsValueString())/1000
#RebarTON = float(wtf_steel.LookupParameter("Volume").AsValueString()[:-3])*rebar_rate
# rPlinth = float(wtf_steel.LookupParameter("rPlinth").AsValueString())*0.002
# rBase = float(wtf_steel.LookupParameter("rBase").AsValueString())*0.002
#FoundVol = float(wtf_steel.LookupParameter("Volume").AsValueString()[:-3])
# pitH = float(wtf_steel.LookupParameter("hBottomVoid").AsValueString())/1000
# pitIN = float(wtf_steel.LookupParameter("rVoidInner").AsValueString())/1000
# pitOUT = float(wtf_steel.LookupParameter("rVoidOuter").AsValueString())/1000

# BlindDia = ((float(blind.LookupParameter("rBlindingExtension").AsValueString())/1000)*2)+rBase
#BlindVol = float(blinding.LookupParameter("Volume").AsValueString()[:-3])
# Blindh = float(blind.LookupParameter("hBlinding").AsValueString())/1000

#backfVol = float(backfill.LookupParameter("Volume").AsValueString()[:-3])

#Excavation = backfVol + BlindVol + FoundVol

#Grout_Vol = float(groutPock.LookupParameter("Volume").AsValueString()[:-3])

# plinthVol = 3.1415*(PlinthH-ConeH)*(rPlinth/2)**2
# pitVol = 3.1415*((pitIN + pitOUT)/2)**2
# slabVol = FoundVol - plinthVol - pitVol








# fam_list = [wtf_steel, grout, anchor_cage, blinding, duct, backfill]

# #  Get hold of parameters for each family
# PA_WTF_Blinding =['Volume', 'Area', 'hBottomVoid', 'rBase', 'hBlinding', 'rBlindingExtension', 'rBottomVoidInner', 'rBottomVoidOuter']
# WTF_SteelTower = ['Volume', 'Area', 'hPlinth', 'hBase', 'hBottomVoid', 'hCone', 'rBase', 'rPlinth', 'rTower', 'dChamfer', 'wGroutTop', 'wGroutBottom', 'dGroutMiddle', 'dGroutSides', 'rSlope', 'rVoidOuter', 'rVoidInner']
# PA_Vestas_Duct_Assembly = ['rBase']
# PA_WTF_Backfill = ['Volume', 'Area', 'hBase', 'hNGL', 'hCone', 'hPlinth', 'orExcavationBottom', 'ovPlinthToOverburden', 'rBase', 'rPlinth', 'sExcavation', 'sOverburdenFill', 'sOverburdenTop']
# PA_WTF_Grout = ['Volume', 'Area', 'rTower', 'dGroutMiddle', 'dGroutSide', 'wGroutTop', 'wGroutBottom']
# PA_AnchorCage_Assembly = ['placehold1','placehold2','ovBoltTop', 'ovBoltBot', 'lBolt', 'tFlangeTop', 'ovTOCtoBOTtopFlange', 'rBoltInner', 'rBoltOuter', 'nBolts', 'hShell', 'rChamfer', 'rTower', 'tShell', 'wFlange', 'vBearingPlate', 'nBolt', 'nSupStud', 'nSupStuds', 'dSupStudHole', 'rSupStudPath', 'tFlangeBot', 'ovSleaveTop', 'ovSleaveBot', 'dBoltHoleTop', 'dBoltHoleBot', 'vSupportStudHoles', 'wFlangeBot', 'tBearingPlate', 'WBearingPlate']

# par_list = [WTF_SteelTower, PA_WTF_Grout, PA_AnchorCage_Assembly, PA_WTF_Blinding, PA_Vestas_Duct_Assembly, PA_WTF_Backfill]



# # Transaction

# t = Transaction(doc)
# t.Start("Apply param val")

# i = 5   
# for item in paraList:
#     p = loadsTable.LookupParameter(item)
#     print(p.Definition.Name)
#     pVal = xl.Cells(i,2).Value2
#     print(pVal)
#     if "_M" in item:
#         p.Set(float(pVal)*10763.9104167097)
#     elif "_F" in item:
#          p.Set(float(pVal)*3280.83989501312)
#     i = i+1
#     print("*"*50)

# t.Commit()


# #AsDouble()/3280.839895