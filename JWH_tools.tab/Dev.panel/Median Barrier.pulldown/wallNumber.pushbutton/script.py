from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Architecture import *
from Autodesk.Revit.DB.Analysis import *
import math
import uuid
from Autodesk.Revit.DB import UnitTypeId
from pyrevit import revit, DB
from pyrevit import forms, output

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

	
def LengthToMM(val):
		return UnitUtils.ConvertFromInternalUnits(val,UnitTypeId.Millimeters)


projLoc = doc.ActiveProjectLocation

# T = Transaction(doc, "Coordinates")
# T.Start()


# Define the family name you want to filter by
family_name = "BVi_2PA_RetainingWall_MedianBarrier_Assembly"

# Create a parameter filter for the Family Name parameter
param_prov = ParameterValueProvider(ElementId(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM))
string_rule = FilterStringRule(param_prov, FilterStringEquals(), family_name)

# Create an element filter with the parameter filter
element_filter = ElementParameterFilter(string_rule)

# Use the element filter in your FilteredElementCollector                                                       .WherePasses(element_filter)
FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()

Count = 1

T = Transaction(doc, "Coordinates")
T.Start()

for elem in FEC:
    if elem.Name == family_name:
        print(elem.Name)
        print(elem.LookupParameter("Assembly Number").AsString())
        OneToTwo = elem.LookupParameter("Trans T1-T2").AsInteger()
        TwoToOne = elem.LookupParameter("Trans T2-T1").AsInteger()
        T1 = elem.LookupParameter("Type 1").AsInteger()
        T2 = elem.LookupParameter("Type 2").AsInteger()
        LM = elem.LookupParameter("RightLightPost").AsInteger()

        #set assembly number
        elem.LookupParameter("Comments").Set(str(Count))
        print(elem.LookupParameter("Comments").AsInteger())
        #set mark to type panel
        if OneToTwo == 1 or TwoToOne == 1:
            elem.LookupParameter("Mark").Set("Transition Panel".upper())
        elif LM == 1 and T1 == 1:
            elem.LookupParameter("Mark").Set("Type 1 with Light Mast".upper())
        elif LM == 1 and T2 == 1:
            elem.LookupParameter("Mark").Set("Type 2 with Light Mast".upper())
        elif T1 == 1:
            elem.LookupParameter("Mark").Set("Type 1".upper())
        elif T2 == 1:
            elem.LookupParameter("Mark").Set("Type 2".upper())
        Count += 1

T.Commit()
        