import clr
import math
import datetime
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import forms

#get date
date = datetime.datetime.now().strftime("%d/%m/%Y")



FPath = forms.pick_file(file_ext='xlsx', multi_file=False, unc_paths=False)

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

#get hould of excel file using ironpython
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
excel = Excel.ApplicationClass()
excel.Visible = False
workbook = excel.Workbooks.Open(FPath)
xl = workbook.Worksheets['SheetDATA']

# Get hold of sheets an tables in document to populate
sheet = doc.GetElement(ElementId(2106182))


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
RebarRateXL = (float(xl.Cells(14,16).Value2))
BackfillClassXL =(xl.Cells(15,16).Value2)
ExcavationClassXL = (xl.Cells(16,16).Value2)
Wind_Farm_Name = str((xl.Cells(29,16).Value2)).upper()
Tower_Supplier = (xl.Cells(30,16).Value2)
Locationval = (xl.Cells(31,16).Value2)
Turbine_Type = (xl.Cells(32,16).Value2)
Tower_Type = (xl.Cells(33,16).Value2)
Loading_Doc = (xl.Cells(5,16).Value2)
Geotechnical_Investigation = (xl.Cells(36,16).Value2)
Subgrade_Reaction_Modulus = (str(float(xl.Cells(23,16).Value2)))
Plastic_Bearing_Pressure = (str(float(xl.Cells(24,16).Value2)))
Average_Founding_Depth = (str((float(xl.Cells(5,4).Value2)-float(xl.Cells(10,12).Value2))/1000))
Allowance_for_Buoyancy = (str(float(xl.Cells(25,16).Value2)/1000))
Developer = (xl.Cells(34,16).Value2)
Foundation_Type = (xl.Cells(35,16).Value2)
Geotechnical_Notes = (xl.Cells(37,16).Value2)
Hub_Height = (str(float(xl.Cells(6,16).Value2)))
Number_of_Turbines = (str(int(xl.Cells(8,16).Value2)))
Project_Name = (xl.Cells(29,16).Value2)



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

f1_plinth = str(round(wtf_steel.GetMaterialArea(ElementId(1017985), usePaintMaterial = False)*0.092903,0))
f1_base = str(round(wtf_steel.GetMaterialArea(ElementId(2096835), usePaintMaterial = False)*0.092903,0))
U2 = ((round(wtf_steel.LookupParameter('rPlinth').AsDouble())*0.3048)**2*math.pi)+round(wtf_steel.GetMaterialArea(ElementId(2096792), usePaintMaterial = False)*0.092903,0)
# U2_Base = str(round(wtf_steel.GetMaterialArea(ElementId(2096792), usePaintMaterial = False)*0.092903,0))


print(30*'#')
print('diagonal slope area U2 :   ' + str(float(round(wtf_steel.GetMaterialArea(ElementId(2096792), usePaintMaterial = False)*0.092903))))
print('hor plint area U2 :    ' + str(round(float(round(wtf_steel.LookupParameter('rPlinth').AsDouble()))**2*math.pi*0.092903)))
print(30*'#')

# Set floats to string with two decimal places for parameter assignment
GroutVol = str(round(GroutVol,2))   
PlinthVol = str(round(PlinthVol,1))
PitVol = str(round(PitVol,1))
BaseVol = str(round(BaseVol,1))
BlVol = str(round(BlVol,1))
RebarTonage = str(round(RebarTonage,2))
BFVol = str(round(BFVol,1))[:-2]
ExVol = str(round(ExVol,1))[:-2]



#  MTO Parameters

t = Transaction(doc)
t.Start("Apply parameter values")

#RebarClass = MTO.LookupParameter("RebarClass")
#BFDens = MTO.LookupParameter("BackfillDens")
# F1 = MTO.LookupParameter("F1 BaseV")
# F2 = MTO.LookupParameter("F2 PlinthV")
# U1 = MTO.LookupParameter("U1 SlabS")
# U2 = MTO.LookupParameter("U2 PlinthH")


sheet.LookupParameter("Sheet Issue Date").Set(date)
sheet.LookupParameter("Grout Class").Set(str(GroutClassXL))
sheet.LookupParameter("Plinth Class").Set(PlinthClassXL)
sheet.LookupParameter("Pit Class").Set(PitClassXL)
sheet.LookupParameter("Foundation Class").Set(BaseClassXL)
sheet.LookupParameter("Blinding Class").Set(BlindingClassXL)
sheet.LookupParameter("Reinforcement Rate").Set(str(RebarRateXL))
sheet.LookupParameter("Backfill").Set(BackfillClassXL)
sheet.LookupParameter("Excavation").Set(ExcavationClassXL)
sheet.LookupParameter("Slab Height").Set(str(float(wtf_steel.LookupParameter("hBase").AsValueString())/1000))
sheet.LookupParameter("Cone Height").Set(str(float(wtf_steel.LookupParameter("hCone").AsValueString())/1000))
sheet.LookupParameter("Plinth Height").Set(str(float(wtf_steel.LookupParameter("hPlinth").AsValueString())/1000))
sheet.LookupParameter("Reinforcement t").Set(RebarTonage)
sheet.LookupParameter("Plinth Diameter").Set(str(float(wtf_steel.LookupParameter("rPlinth").AsValueString())*2/1000))
sheet.LookupParameter("Foundation Diameter").Set(str(float(wtf_steel.LookupParameter("rBase").AsValueString())*2/1000))

blindingDiameter = (float(blinding.LookupParameter("rBase").AsValueString()) + float(blinding.LookupParameter("rBlindingExtension").AsValueString()))*2/1000

sheet.LookupParameter("Blinding Diameter").Set(str(blindingDiameter))
sheet.LookupParameter("Blinding Vol").Set(BlVol)
sheet.LookupParameter("Average Blinding Depth").Set(str(float(blinding.LookupParameter("hBlinding").AsValueString())/1000))
sheet.LookupParameter("Backfill").Set(BFVol)
sheet.LookupParameter("Excavation Vol").Set(ExVol)

sheet.LookupParameter("Grout Vol").Set(GroutVol)
sheet.LookupParameter("Plinth Vol").Set(PlinthVol)
sheet.LookupParameter("Pitt Vol").Set(PitVol)
sheet.LookupParameter("Foundation Vol").Set(BaseVol)
sheet.LookupParameter("Reinforcement t").Set(RebarTonage)

sheet.LookupParameter("Tower Type").Set(Tower_Type)
sheet.LookupParameter("Loading Doc").Set(Loading_Doc)
sheet.LookupParameter("Plastic Bearing Pressure").Set(Plastic_Bearing_Pressure)
sheet.LookupParameter("Average Founding Depth").Set(Average_Founding_Depth)
sheet.LookupParameter("Allowance for Buoyancy").Set(Allowance_for_Buoyancy)
sheet.LookupParameter("Developer").Set(Developer)
sheet.LookupParameter("Foundation Type").Set(Foundation_Type)
sheet.LookupParameter("Geotechnical Notes").Set(Geotechnical_Notes)
sheet.LookupParameter("Wind Farm Name").Set(Wind_Farm_Name)
sheet.LookupParameter("Turbine Type").Set(Turbine_Type)
sheet.LookupParameter("Subgrade Reaction Modulus").Set(Subgrade_Reaction_Modulus)
sheet.LookupParameter("Geotechnical Investigation").Set(Geotechnical_Investigation)
sheet.LookupParameter("Turbine Type").Set(Turbine_Type)
sheet.LookupParameter("Location").Set(Locationval)
sheet.LookupParameter("Tower Supplier").Set(Tower_Supplier)
sheet.LookupParameter("Number of towers").Set(Number_of_Turbines)
sheet.LookupParameter("Hub Height").Set(Hub_Height)


sheet.LookupParameter("U1 and U2 Finish").Set(str(round(U2,0)))
sheet.LookupParameter("Vertical Plinth Formwork").Set(f1_plinth)
sheet.LookupParameter("Vertical Slab Formwork").Set(f1_base)

# doc.Project.SetProjectName(Wind_Farm_Name)
doc.GetElement(ElementId(1250563)).Text = date

t.Commit()

print("Material Take Off Parameters Updated in A4 sheet")
# Close Excel application object
workbook.Close(False)
excel.Quit()
print("Excel Closed")






