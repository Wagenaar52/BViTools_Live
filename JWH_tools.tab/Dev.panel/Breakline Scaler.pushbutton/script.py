from Autodesk.Revit.DB import Transaction, Structure, FilteredElementCollector, ElementId, ElementParameterFilter, ParameterValueProvider, FilterStringRule, FilterStringEquals
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, Line, XYZ
from pyrevit import forms
from Autodesk.Revit.DB import ElementId
from Autodesk.Revit.DB import ElementParameterFilter, FilterStringRule, FilterStringEquals

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document


# Define the family name you want to filter by
family_name = "BVI-Breakline"

# Create a parameter filter for the Family Name parameter
param_prov = ParameterValueProvider(ElementId(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM))
string_rule = FilterStringRule(param_prov, FilterStringEquals(), family_name)

# Create an element filter with the parameter filter
element_filter = ElementParameterFilter(string_rule)

# Use the element filter in your FilteredElementCollector
FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DetailComponents).WhereElementIsNotElementType().WherePasses(element_filter).ToElements()

i = 0

for element in FEC:
    owner_view = doc.GetElement(element.OwnerViewId)
    t = Transaction(doc, "Change Scale")
    t.Start()
    Param = element.LookupParameter("Scale")
    Param.Set(int(owner_view.Scale)/304.8)
    t.Commit()
    i += 1
