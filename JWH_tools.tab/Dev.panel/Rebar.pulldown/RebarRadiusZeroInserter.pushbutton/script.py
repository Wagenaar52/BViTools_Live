# IMPORTS
from pyrevit import forms


from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import *

#VARIABLES

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document



#FUNCTIONS

def get_selected_elements(uidoc):
    """This function returns a list of selected elements in Revit UI.
    :param uidoc:   uidoc where elements are selected.
    :return:        list of selected elements.
    """
    return [uidoc.Document.GetElement(elem_id) for elem_id in uidoc.Selection.GetElementIds()]

   #retrieve all rebat with an empthy Rebar r Custom Parameter


#  .get_Parameter("Rebar r Custom")


rebar = FilteredElementCollector(doc).OfClass(typeof(Rebar))

for r in rebar:
    if r.LookupParameter("Rebar r Custom").HasValue() == True:
        print(r.Id)
        print(r.LookupParameter("Rebar r Custom").AsString())

for r in rebar:
    if r.LookupParameter("Rebar r Custom").










# IMPORTS

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import *

#VARIABLES

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document



#  .get_Parameter("Rebar r Custom")


#rebar = FilteredElementCollector(doc).OfClass(typeof(Rebar))



rebar = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar)

for r in rebar:
#	if r.LookupParameters("Rebar r Custom").ToString() == "50.0":
#		print("IN LOOP ")
#	if r.get_Parameter("Rebar r Custom") == "50.0":
#		print(r.Id)
#print(r.LookupParameter("Rebar r Custom").AsString())