# IMPORTS

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

   