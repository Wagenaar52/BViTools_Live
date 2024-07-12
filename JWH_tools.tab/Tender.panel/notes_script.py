import clr
import math
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

#get hould of excel file using ironpython
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
excel = Excel.ApplicationClass()
excel.Visible = False
workbook = excel.Workbooks.Open(r"""c:\Users\Wagner.Human\Desktop\Tender Drawing Template\IMPORT_FILEREV4.xlsx""")
xl = workbook.Worksheets['SheetDATA']

# Get hold of sheets an tables in document to populate
notes = doc.GetElement(ElementId(1056082))


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

# filledString = "NOTES: \n" \
# "\n" \
# "1.	DESIGN BASIS\n" \
# "1.1	THE DESIGN IS BASED ON A 2D FINITE ELEMENT PLATE MODEL USING 	SOFiSTiK 	SOFTWARE\n" \
# "1.2	THE FOLLOWING CODES AND GUIDELINES ARE ADHERED TO:\n" \
# "	 -IEC 61400-1 THIRD EDITION (2005-08)\n" \
# "  	 -IEC 61400-6 FIRST EDITION (2020-04)\n" \
# "    -DNV-GL GUIDELINES FOR THE CERTIFICATION OF WIND TURBINES\n" \
# "    -BS EN  1992-1-1:2004 - GENERAL RULES AND RULES FOR BUILDINGS\n" \
# "    -SABS 0100-1 - STRUCTURAL USE OF CONCRETE (LOCAL CODE)\n" \
# "    -fib MODEL CODE 2010 (FATIGUE ANALYSIS)\n" \
# "\n" \
# "2.	MATERIAL SPECIFICATIONS\n" \
# "2.1	CONCRETE (AS PER EN 1992-1-1-2004):\n" \
# "       BLINDING:				{} MPa\n" \
# "       BASE:					{} MPa, 	E = 34 GPa\n" \
# "       PLINTH:			    	{} MPa, 	E = 37 GPa\n" \
# "       HIGH STRENGTH GROUT:	{} MPa,    E = 44 GPa\n" \
# "       POISSON RATIO:			0.2\n" \
# "2.2	STEEL REINFORCEMENT:\n" \
# "        HIGH TENSILE STEEL:\n" \
# "	    500 MPa YIELD STRESS			\n" \
# "       (GR B500B AS PER BS 4449:2005)\n" \
# "2.3	CABLE DUCT:\n" \
# "       DUCT MATERIAL:			 UNPLASTICISED POLYVINYL CHLORIDE						\n" \
# "                            DUCTS IN ACCORDANCE WITH SANS 1061:2017\n" \
# "					         WITH SMOOTH INNER LINING.\n" \
# "2.4	MINIMUM DENSITIES:\n" \
# "    REINFORCED CONCRETE:	24 kN/m3\n" \
# "    DRY BACKFILL:			18 kN/m3\n" \
# "    SATURATED SOIL:			21 kN/m3\n" \
# "\n" \
# "3.	APPLIED LOADS\n" \
# "	3.1	Foundation design inputs SG 5.0-145 CIIB HH127.5M (T127.5.43) and BC T.SP. SG\n" \
# "		5.0-145 T127.5.43 IIB M39\n" \
# "3.2	PRESTRESSING OF ANCHOR BOLTS: 515 kN/BOLT\n" \
# "\n" \
# "4.	GEOTECHNICAL (TO BE CONFIRMED AFTER COMPLETION OF INVESTIGATION)\n" \
# "4.1	DESIGN MINIMUM BEARING PRESSURE:\n" \
# "    OPERATIONAL LOADS (SLS):			\t\t\t{} kPa\n" \
# "    EXTREME LOADS (EQU):				{} kPa\n" \
# "    EXTREME LOADS TRANSVERSE(EQU):		50 kPa	\n" \
# "4.2	DESIGN SUBGRADE REACTION MODULUS\n" \
# "    SOIL SPRING STIFFNESS 		ksv:		{} MPa/m\n" \
# "			            		ksh:		{} MPa/m	\n" \
# "4.3	BEARING CAPACITY AND FOUNDING STIFFNESS TO BE CONFIRMED ON SITE\n" \
# " 	BY 	GEOTECHNICAL ENGINEER\n" \
# "4.4	DYNAMIC ROTATIONAL STIFFNESS OF FOUNDATION TO BE EQUAL TO OR 	\n" \
# "    LARGER THAN {} GNm/rad AT S3 LOAD LEVEL AS PER IEC61400-6 8.5.3.2\n" \
# "4.5	EXCAVATION PROFILE TO BE CONFIRMED BY COMPETENT PERSON AT EACH 	LOCATION\n" \
# "\n" \
# "5.	CONSTRUCTION\n" \
# "5.1	CONCRETE COVER TO ALL FACES:	50 mm\n" \
# "5.2	CONCRETE FINISH:\n" \
# "    HIDDEN HORIZONTAL:			    U1\n" \
# "    HIDDEN VERTICAL:			    F1\n" \
# "    EXPOSED HORIZONTAL:			    U2\n" \
# "    EXPOSED VERTICAL:			    F2\n" \
# "    CHAMFERS:					    40 mm\n" \
# "5.3	ALL CONCRETE WORKS SHALL BE STRICTLY IN ACCORDANCE WITH COTO\n" \
# " 	STANDARD SPECIFICATIONS FOR ROAD AND BRIDGE WORKS FOR SOUTH\n" \
# "	AFRICAN AUTHORITIES\n"\

# MyString = filledString.format(BlindingClassXL, BaseClassXL, PlinthClassXL, GroutClassXL,Plastic_Bearing_Pressure ,Subgrade_Reaction_Modulus , ksh, DyRotSitff)


MyString = """
NOTES: 

1.	DESIGN BASIS
1.1	THE DESIGN IS BASED ON A 2D FINITE ELEMENT PLATE MODEL USING 	SOFiSTiK 	SOFTWARE
1.2	THE FOLLOWING CODES AND GUIDELINES ARE ADHERED TO:
	-IEC 61400-1 THIRD EDITION (2005-08)
	-IEC 61400-6 FIRST EDITION (2020-04)
    -DNV-GL GUIDELINES FOR THE CERTIFICATION OF WIND TURBINES
    -BS EN  1992-1-1:2004 - GENERAL RULES AND RULES FOR BUILDINGS
    -SABS 0100-1 - STRUCTURAL USE OF CONCRETE (LOCAL CODE)
    -fib MODEL CODE 2010 (FATIGUE ANALYSIS)

2.	MATERIAL SPECIFICATIONS
2.1	CONCRETE (AS PER EN 1992-1-1-2004):
    BLINDING:				"""+ BlindingClassXL 
"""} MPa
    BASE:					{BaseClassXL} MPa, 	E = 34 GPa
    PLINTH:			    	{PlinthClassXL} MPa, 	E = 37 GPa
    HIGH STRENGTH GROUT:	{GroutClassXL} MPa,    E = 44 GPa
    POISSON RATIO:			0.2
2.2	STEEL REINFORCEMENT:
    HIGH TENSILE STEEL:
	500 MPa YIELD STRESS			
    (GR B500B AS PER BS 4449:2005)
2.3	CABLE DUCT:
    DUCT MATERIAL:			UNPLASTICISED POLYVINYL CHLORIDE						
                            DUCTS IN ACCORDANCE WITH SANS 1061:2017
					        WITH SMOOTH INNER LINING.
2.4	MINIMUM DENSITIES:
    REINFORCED CONCRETE:	24 kN/m3
    DRY BACKFILL:			18 kN/m3
    SATURATED SOIL:			21 kN/m3

3.	APPLIED LOADS
	3.1	Foundation design inputs SG 5.0-145 CIIB HH127.5M (T127.5.43) and BC T.SP. SG
		5.0-145 T127.5.43 IIB M39
3.2	PRESTRESSING OF ANCHOR BOLTS: 515 kN/BOLT

4.	GEOTECHNICAL (TO BE CONFIRMED AFTER COMPLETION OF INVESTIGATION)
4.1	DESIGN MINIMUM BEARING PRESSURE:
    OPERATIONAL LOADS (SLS):			{Plastic_Bearing_Pressure} kPa
    EXTREME LOADS (EQU):				{Plastic_Bearing_Pressure} kPa
    EXTREME LOADS TRANSVERSE(EQU):		50 kPa	
4.2	DESIGN SUBGRADE REACTION MODULUS
    SOIL SPRING STIFFNESS 		ksv:		{Subgrade_Reaction_Modulus} MPa/m
			            		ksh:		{ksh} MPa/m	
4.3	BEARING CAPACITY AND FOUNDING STIFFNESS TO BE CONFIRMED ON SITE
 	BY 	GEOTECHNICAL ENGINEER
4.4	DYNAMIC ROTATIONAL STIFFNESS OF FOUNDATION TO BE EQUAL TO OR 	
    LARGER THAN {DyRotSitff} GNm/rad AT S3 LOAD LEVEL AS PER IEC61400-6 8.5.3.2
4.5	EXCAVATION PROFILE TO BE CONFIRMED BY COMPETENT PERSON AT EACH 	LOCATION

5.	CONSTRUCTION
5.1	CONCRETE COVER TO ALL FACES:	50 mm
5.2	CONCRETE FINISH:
    HIDDEN HORIZONTAL:			    U1
    HIDDEN VERTICAL:			    F1
    EXPOSED HORIZONTAL:			    U2
    EXPOSED VERTICAL:			    F2
    CHAMFERS:					    40 mm
5.3	ALL CONCRETE WORKS SHALL BE STRICTLY IN ACCORDANCE WITH COTO
 	STANDARD SPECIFICATIONS FOR ROAD AND BRIDGE WORKS FOR SOUTH
	AFRICAN AUTHORITIES"""


t = Transaction(doc)
t.Start("Apply parameter values")

notes.Text = MyString

t.Commit()

print("Notes Updated in General Arangement Drawing")
# Close Excel application object
workbook.Close(False)
excel.Quit()
print("Excel Closed")


