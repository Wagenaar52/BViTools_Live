from Autodesk.Revit.DB import Transaction, Structure, FilteredElementCollector, ElementId, ElementParameterFilter, ParameterValueProvider, FilterStringRule, FilterStringEquals
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, Line, XYZ
from pyrevit import forms, output
import math
from Autodesk.Revit.DB import ElementId

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
selection = uidoc.Selection
view = doc.ActiveView

# get all rebar elements with mark "DOWEL"
f_param = ParameterValueProvider(ElementId(BuiltInParameter.DOOR_NUMBER))

evaluator = FilterStringEquals()

f_param_value = "DOWEL"

f_rule = FilterStringRule(f_param, evaluator, f_param_value)

# apply filter to elements in active view
FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().WherePasses(ElementParameterFilter(f_rule)).ToElements()

for element in FEC:
    print(element.Name)
    print(element.LookupParameter("Mark").AsString())

print('see attached list {x} with mark {y}'.format(x=FEC, y=f_param_value))


output.PyRevitOutputWindow.Close()
# # create a filter

# filter = ElementParameterFilter(f_rule)

# # apply filter to elements in active view

# FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().WherePasses(filter).ToElements()

# for element in FEC:
#     print(element.Name)
#     print(element.LookupParameter("Mark").AsString())

# # collect all structural rebar elements in the project in a filtered element collector

# # selected_elements = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralRebar).WhereElementIsNotElementType().ToElements()

# # print(FEC)
# # print(selected_elements)
# # print("code ran")



# for element in selected_elements:
#     if element.Category.Name == 'Structural Rebar':


