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

tr = Transaction(doc, "Coordinates")
tr.Start()


# Use the element filter in your FilteredElementCollector                                                       .WherePasses(element_filter)
FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()



sopFam = "1PA_BVi_Setting Out Point_RW_1"
pannelNum = "BVi_SRW_Global Element Number"
family_name = "Bottom Polyhedron"

for elem in FEC:
    if elem.Name == sopFam:
        # print(elem.Name)
        # print(elem.LookupParameter("BVI_SRW_SOP_Ref").AsString())    
        # print(elem.LookupParameter("Setting Out Point Number").AsInteger())    
        for wall in FEC:
            if wall.Name == family_name:
                print("####################")    
                print(str(elem.LookupParameter("Setting Out Point Number").AsValueString()))
                print(str(wall.LookupParameter(pannelNum).AsInteger()))
                if str(elem.LookupParameter("Setting Out Point Number").AsValueString()) == str(wall.LookupParameter(pannelNum).AsInteger()):
                    elem.LookupParameter("Mark").Set(float(elem.LookupParameter("BVI_SRW_SOP_Ref").AsString()))
                    print("Update")

tr.Commit()
