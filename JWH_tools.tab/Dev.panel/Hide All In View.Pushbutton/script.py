
import clr
import math
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *


uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
selection = uidoc.Selection
view = doc.ActiveView

typeof = clr.GetClrType(FamilyInstance)

FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()

ElementCategoryFilter.FamilyFilter
typeof('BVi_1PA_Pre-Cast_Barrier')


print(FEC)