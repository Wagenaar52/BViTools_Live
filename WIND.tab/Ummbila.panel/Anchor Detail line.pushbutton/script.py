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




# Get Elements from document
#wtf_steel = doc.GetElement(ElementId(1087568))
#grout = doc.GetElement(ElementId(1804084))
#blinding = doc.GetElement(ElementId(1097087))
#backfill = doc.GetElement(ElementId(2342235))
#anchor_cage = doc.GetElement(ElementId(1322206))

WTF = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
for element in WTF:
    if element.Name == "1PA_AnchorCage_Assembly 2":
        anchor_cage = element
    elif element.Name == "1PA_WTF_Grout":
        grout = element
    elif element.Name == "1PA_WTF_SteelTower":
        wtf_steel = element
    elif element.Name == "1PA_WTF_Blinding":
        blinding = element
    elif element.Name == "1PA_WTF_Backfill":
        backfill = element



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


planViewOfAnchorCage = doc.GetElement(ElementId(792567))
planViewOfBottomAnchorCage = doc.GetElement(ElementId(794699))    

centreLineStyle = doc.GetElement(ElementId(1018897))
hiddenLineStyle = doc.GetElement(ElementId(792574))

viewList = [planViewOfAnchorCage, planViewOfBottomAnchorCage]













# Start the transaction
t = Transaction(doc, "Dimension Center Line")
t.Start()

# Get the plan view element
planViewOfAnchorCage = doc.GetElement(ElementId(792567))

# Define the center point, radius, and the normal vectors for the arc (circle)
centerPoint = wtf_steel.Location.Point
radius = towerFlangeIN

# Create a plane for the arc
plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, centerPoint)

# Create the arc (full circle)
cl_towerFlangeIN = Arc.Create(plane, radius, 0, 2 * math.pi)

# Create the detail curve (circle) in the plan view
cl = doc.Create.NewDetailCurve(planViewOfAnchorCage, cl_towerFlangeIN)

# Assign the line style (make sure centreLineStyle is defined and valid)
cl.LineStyle = centreLineStyle

# Get the dimension type
dimType = doc.GetElement(ElementId(1018440))

# Create a reference array and add the reference of the circle's geometry
refArray = ReferenceArray()
refArray.Append(cl.GeometryCurve.Reference)

# Define the radial dimension arc line location
dimensionArcLine = Line.CreateBound(centerPoint, XYZ(centerPoint.X + radius, centerPoint.Y, centerPoint.Z))

# Create the radial dimension using the NewRadialDimension method from the Document class
dim = doc.Create.NewRadialDimension(planViewOfAnchorCage, cl.GeometryCurve.Reference, centerPoint, dimType)

# Commit the transaction
t.Commit()

print("Detail line circle and radial dimension created.")


























# t = Transaction(doc)
# t.Start("Dimension Center Line")


# cl_towerFlangeIN = Arc.Create(wtf_steel.Location.Point, towerFlangeIN, 0.1, 2*math.pi, XYZ(1,0,0), XYZ(0,1,0))
# cl = doc.Create.NewDetailCurve(planViewOfAnchorCage, cl_towerFlangeIN)
# cl.LineStyle = centreLineStyle

# dimType = doc.GetElement(ElementId(1018440))



# refArray = ReferenceArray()
# refArray.Append(cl.GeometryCurve.Reference)
# doc.Create.NewRadialDimension(planViewOfAnchorCage, refArray, XYZ(0,0,0),dimType)





# t.Commit()
