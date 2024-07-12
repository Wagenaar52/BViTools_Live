from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Architecture import *
from Autodesk.Revit.DB.Analysis import *
import math
import uuid
from Autodesk.Revit.DB import UnitTypeId
from pyrevit import revit, DB
from pyrevit import forms, output
from pyrevit import revit, DB

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

projLoc = doc.ActiveProjectLocation

def LengthToMM(val):
		return UnitUtils.ConvertFromInternalUnits(val,UnitTypeId.Millimeters)


tr = Transaction(doc, "Coordinates")
tr.Start()

# Define the family name you want to filter by
family_name =  "LHS" #"1PA_BVi_Setting Out Point_1"

# Create a parameter filter for the Family Name parameter
param_prov = ParameterValueProvider(ElementId(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM))
string_rule = FilterStringRule(param_prov, FilterStringEquals(), family_name)

# Create an element filter with the parameter filter
element_filter = ElementParameterFilter(string_rule)

# Use the element filter in your FilteredElementCollector                                                       .WherePasses(element_filter)
FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()

for elem in FEC:
    if elem.Name == "BVi_2PA_RetainingWall_MedianBarrier_Assembly":
        for subcom in elem.GetDependentElements(None):
               if doc.GetElement(subcom).Name == "1PA_BVi_Setting Out Point_1" :#and doc.GetElement(subcom).LookupParameter("Setting Out Point Reference").AsString() == "SOP_FOOTING_START_F":     
                    locPnt = doc.GetElement(subcom).Location.Point
                    position = projLoc.GetProjectPosition(locPnt)
                    locX = -LengthToMM(position.EastWest)/1000
                    locY = -LengthToMM(position.NorthSouth)/1000
                    locZ = LengthToMM(position.Elevation)/1000
                    # assNr = elem.LookupParameter("Assembly Number").AsInteger()
                    # elem.LookupParameter("Mark").Set("SOP L  PANEL %d" %assNr) 
                    elem.LookupParameter("X").Set(locX)
                    elem.LookupParameter("Y").Set(locY)
                    elem.LookupParameter("Elevation").Set(locZ)


    # elif elem.Name == "RHS":
    #     for subcom in elem.GetDependentElements(None):
    #            if doc.GetElement(subcom).Name == "1PA_BVi_Setting Out Point_1" and doc.GetElement(subcom).LookupParameter("Setting Out Point Reference").AsString() == "SOP_FOOTING_START_B":     
    #                 locPnt = doc.GetElement(subcom).Location.Point
    #                 position = projLoc.GetProjectPosition(locPnt)
    #                 locX = -LengthToMM(position.EastWest)/1000
    #                 locY = -LengthToMM(position.NorthSouth)/1000
    #                 locZ = LengthToMM(position.Elevation)/1000
    #                 assNr = elem.LookupParameter("Assembly Number").AsInteger()
    #                 elem.LookupParameter("Mark").Set("SOP R  PANEL %d" % assNr) 
    #                 elem.LookupParameter("X").Set(locX)
    #                 elem.LookupParameter("Y").Set(locY)
    #                 elem.LookupParameter("Elevation").Set(locZ)

tr.Commit()