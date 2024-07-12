__title__ = "Place SOPs from Excel"
__doc__ = """This script will place SOPs from an excel file into the model. The excel file should have the following columns: |Name| X| Y| Elevation| in colums A|B|C|D respektively. The script will place the SOPs at the X and Y coordinates and set the elevation parameters. The script will also coordinate the SOPs in the model. \n ======\n Make sure to have the correct family (1PA_BVi_Setting Out Point) loaded in the model. \n Make sure to that the excel file is saved in .xlsx format \n Make sure to have the correct columns in the excel file. \n Make sure to name the sheet 'A'in your excel file"""

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Architecture import *
from Autodesk.Revit.DB.Analysis import *
import math
import uuid
import clr
from Autodesk.Revit.DB import UnitTypeId
from pyrevit import revit, DB
from pyrevit import forms, output
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel


uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

#read excel file
ExcelfamilyPath = forms.pick_file(file_ext='xlsx', multi_file=False, unc_paths=False)
excel = Excel.ApplicationClass()
excel.Visible = False
workbook = excel.Workbooks.Open(ExcelfamilyPath)
xl = workbook.Worksheets['A']
SOPdict = {}
for i in range(1,200):
    i += 1
    if "BH " in str(xl.Cells(i, 1).Value2):
        name = str(xl.Cells(i,1).Value2)
        xCord = str(xl.Cells(i,2).Value2)
        yCord = str(xl.Cells(i,3).Value2)
        elevation = str(xl.Cells(i,4).Value2)
        comment = str(xl.Cells(i,5).Value2)
        SOPdict[name] = [xCord,yCord,elevation,comment]




class FamilyLoadOptions(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = True
        return True

    def OnSharedFamilyFound(self, sharedFamily, overwriteParameterValues):
        overwriteParameterValues.Value = True
        return True




FEC = FilteredElementCollector(doc).OfClass(Family).ToElements()

for family in FEC:
    if family.Name == "1PA_BVi_Setting Out Point":
        family = family

print(family)


def LengthToMM(val):
		return UnitUtils.ConvertFromInternalUnits(val,UnitTypeId.Millimeters)

# familyPath = forms.pick_file(file_ext='rfa', multi_file=False, unc_paths=False)

def placeSOPfromExcel(name, familyPath, SOPdict, doc):
    t=Transaction(doc, "Place SOPs")
    t.Start()

    
    # Get the family symbol
    family_symbol = None
    for symbol_id in family.GetFamilySymbolIds():
        family_symbol = doc.GetElement(symbol_id)
        break

    if not family_symbol:
        print("Failed to get family symbol.")
        raise Exception("Family symbol not found.")    

    # Ensure the symbol is active
    if not family_symbol.IsActive:
        family_symbol.Activate()
        doc.Regenerate()

    adaptivePoint = XYZ(float(SOPdict[name][0]),float(SOPdict[name][1]),0)



    # Create the family instance
    family_instance = AdaptiveComponentInstanceUtils.CreateAdaptiveComponentInstance(doc, family_symbol)

    family_instance.LookupParameter("X").Set(SOPdict[name][1])
    family_instance.LookupParameter("Y").Set(SOPdict[name][2])
    family_instance.LookupParameter("Elevation").Set(SOPdict[name][3])


    # Get the placement points
    placement_points = AdaptiveComponentInstanceUtils.GetInstancePlacementPointElementRefIds(family_instance)

    # Set the placement points
    for i, point_id in enumerate(placement_points):
        point = doc.GetElement(point_id)
        point.Position = adaptivePoint

    # Commit the transaction

    print("Adaptive family placed successfully.")
    t.Commit()

for name in SOPdict:
    placeSOPfromExcel(name, ExcelfamilyPath, SOPdict, doc)



print ("###########################################All SOPs placed successfully.")



t=Transaction(doc, "Coordinate Update")
t.Start()

projLoc = doc.ActiveProjectLocation

outVar = []

# Define the family name you want to filter by
family_name = "1PA_BVi_Setting Out Point"

# Create a parameter filter for the Family Name parameter
param_prov = ParameterValueProvider(ElementId(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM))
string_rule = FilterStringRule(param_prov, FilterStringEquals(), family_name)

# Create an element filter with the parameter filter
element_filter = ElementParameterFilter(string_rule)

# Use the element filter in your FilteredElementCollector                                                       .WherePasses(element_filter)
FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()

count = 0
for elem in FEC:
    if elem.Name == family_name:
        print(elem.Name)
        count += 1
        elemFam=elem.Category.Name
        location = elem.Location
        locPnt = elem.Location.Point
        position = projLoc.GetProjectPosition(locPnt)
        x = round(LengthToMM(position.EastWest))
        y = round(LengthToMM(position.NorthSouth))
        z = round(LengthToMM(position.Elevation))
        xParam = elem.LookupParameter("E/W Coordinate")
        yParam = elem.LookupParameter("N/S Coordinate")
        zParam = elem.LookupParameter("Elevation")
        xSAParam = elem.LookupParameter("X")
        ySAParam = elem.LookupParameter("Y")
        #if not xParam or not yParam or not zParam: continue
        xParam.Set(x/1000)
        yParam.Set(y/1000)
        zParam.Set(str(z/1000))
        xSAParam.Set(y/-1000)
        ySAParam.Set(x/-1000)
        # outVar.append([elem,elemFam,location,locPnt])
        print(z/1000)

t.Commit()
#close excel
workbook.Close(False)
excel.Quit()