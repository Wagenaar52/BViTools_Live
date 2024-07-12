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
notes = doc.GetElement(ElementId(1056082))


# Get Elements from document
wtf_steel = doc.GetElement(ElementId(1087568))
grout = doc.GetElement(ElementId(1804084))
blinding = doc.GetElement(ElementId(1097087))
backfill = doc.GetElement(ElementId(2342235))
anchor_cage = doc.GetElement(ElementId(1322206))

# Create variables from excel sheet
GroutClassXL = (xl.Cells(9,16).Value2)
PlinthClassXL =(xl.Cells(10,16).Value2)
PitClassXL = (xl.Cells(11,16).Value2)
BaseClassXL =(xl.Cells(12,16).Value2)
BlindingClassXL =(xl.Cells(13,16).Value2)
RebarRateXL = (float(xl.Cells(14,16).Value2))
BackfillClassXL =(xl.Cells(15,16).Value2)
ExcavationClassXL = (xl.Cells(16,16).Value2)
Wind_Farm_Name = (xl.Cells(29,16).Value2)
Tower_Supplier = (xl.Cells(30,16).Value2)
Locationval = (xl.Cells(31,16).Value2)
Turbine_Type = (xl.Cells(32,16).Value2)
Tower_Type = (xl.Cells(33,16).Value2)
Loading_Doc = (xl.Cells(5,16).Value2)
Geotechnical_Investigation = (xl.Cells(36,16).Value2)
Subgrade_Reaction_Modulus = (str(float(xl.Cells(23,16).Value2)))
Plastic_Bearing_Pressure = (str(float(xl.Cells(24,16).Value2)))
Average_Founding_Depth = (str(float(xl.Cells(6,12).Value2)/1000))
Allowance_for_Buoyancy = (str(float(xl.Cells(25,16).Value2)/1000))
Developer = (xl.Cells(34,16).Value2)
Foundation_Type = (xl.Cells(35,16).Value2)
Geotechnical_Notes = (xl.Cells(37,16).Value2)
Hub_Height = (str(float(xl.Cells(6,16).Value2)))
Number_of_Turbines = (str(float(xl.Cells(7,16).Value2)))
DyRotSitff = (str(float(xl.Cells(17,16).Value2)))

# Volume Caculations
meanPitRad = ((wtf_steel.LookupParameter('rVoidOuter').AsDouble()*0.3048)+(wtf_steel.LookupParameter('rVoidInner').AsDouble()*0.3048))/2

concVol = wtf_steel.LookupParameter("Volume").AsDouble()*0.3048**3
GroutVol = grout.LookupParameter("Volume").AsDouble()*0.3048**3
PlinthVol = (wtf_steel.LookupParameter('rPlinth').AsDouble()*0.3048)**2*math.pi*(wtf_steel.LookupParameter('hPlinth').AsDouble()*0.3048)
PitVol = (meanPitRad)**2*math.pi*(wtf_steel.LookupParameter('hBottomVoid').AsDouble()*0.3048)
BaseVol = concVol - PlinthVol - PitVol
BlVol = blinding.LookupParameter("Volume").AsDouble()*0.3048**3
RebarTonage = wtf_steel.LookupParameter("Volume").AsDouble()*0.3048**3*RebarRateXL/1000
BFVol = backfill.LookupParameter("Volume").AsDouble()*0.3048**3
ExVol = BFVol + BlVol + concVol

# Set floats to string with two decimal places for parameter assignment
GroutVol = str(round(GroutVol,2))   
PlinthVol = str(round(PlinthVol,2))
PitVol = str(round(PitVol,2))
BaseVol = str(round(BaseVol,2))
BlVol = str(round(BlVol,2))
RebarTonage = str(round(RebarTonage,2))
BFVol = str(round(BFVol,2))
ExVol = str(round(ExVol,2))
ksh = float(Subgrade_Reaction_Modulus)/2 




# Create string for parameter assignment

D1tag1 = doc.GetElement(ElementId(602345))
D1tag2 = doc.GetElement(ElementId(1132204))
D1tag3 = doc.GetElement(ElementId(1120004))
D2tag1 = doc.GetElement(ElementId(1131474))
D2tag2 = doc.GetElement(ElementId(1131644))
D2tag3 = doc.GetElement(ElementId(1942706))
S1tag1 = doc.GetElement(ElementId(1930425))
S1tag2 = doc.GetElement(ElementId(679023))
S1tag3 = doc.GetElement(ElementId(679489))
P1tag1 = doc.GetElement(ElementId(1166586))
D1tag4 = doc.GetElement(ElementId(679489))

t = Transaction(doc)
t.Start("Apply parameter values")

D1tag1.Text = "2 x %d pcs. M%d GRADE 10.9 ANCHOR BOLTS" % (float(xl.Cells(12,8).Value2), float(xl.Cells(33,8).Value2))
D1tag2.Text = "2 x %d pcs. M%d GRADE 10.9 NUTS WITH WASHERS" % (float(xl.Cells(12,8).Value2), float(xl.Cells(33,8).Value2))
D1tag3.Text = "HIGH STRENGTH NON-SHRINK GROUT %s" % str(xl.Cells(9,16).Value2)
D2tag1.Text = "2 x %d pcs. M%d GRADE 10.9 ANCHOR BOLTS" % (float(xl.Cells(12,8).Value2), float(xl.Cells(33,8).Value2))
D2tag2.Text = "2 x %d pcs. M%d GRADE 10.9 NUTS WITH WASHERS" % (float(xl.Cells(12,8).Value2), float(xl.Cells(33,8).Value2))   
D2tag3.Text = "%s BLINDING MIN %d THK" % (str(xl.Cells(13,16).Value2), float(xl.Cells(7,10).Value2))
S1tag1.Text = "CLASS %s CONCRETE" % (str(xl.Cells(10,16).Value2))
S1tag2.Text = "CLASS %s CONCRETE (NO COLD JOINTS ALLOWED)" % (str(xl.Cells(12,16).Value2))
S1tag3.Text = "MINIMUM 200mm THICK %s CONCRETE BLINDING LAYER" % (str(xl.Cells(13,16).Value2))
P1tag1.Text = "2 x %d pcs. M%d GRADE 10.9 ANCHOR BOLTS COMPLETE WITH 3 x 2 x %d pcs. GRADE 10.9 NUTS AND WASHERS" % (float(xl.Cells(12,8).Value2), float(xl.Cells(33,8).Value2), float(xl.Cells(12,8).Value2))
D1tag4.Text = "MINIMUM %smm THICK %s CONCRETE BLINDING LAYER" % (str(float(xl.Cells(7,10).Value2))[:3], str(xl.Cells(13,16).Value2))

t.Commit()

print("condeRan ")
# Close Excel application object
workbook.Close(False)
excel.Quit()



