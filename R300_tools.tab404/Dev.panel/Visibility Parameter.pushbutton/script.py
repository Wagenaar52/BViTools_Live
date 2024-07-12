import clr
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import forms
from rpw import revit, db, ui, DB, UI  
import clr
from Autodesk.Revit.DB import *

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

selection = uidoc.Selection
view = doc.ActiveView

selected_elementsID = selection.GetElementIds()
selected_elements = [doc.GetElement(eid) for eid in selected_elementsID]

# get parameters of selected element

ElementParameters = []

for element in selected_elements:
    for param in element.Parameters:
        if param.StorageType == StorageType.Integer:
            ElementParameters.append(param.Definition.Name)

ElementParameters = list(set(ElementParameters))

EXCL_list = ['Flip', 'Export to IFC', 'Assembly Number', 'Workset','Site Plan Lines' ]

for param in ElementParameters:
    if param in EXCL_list:
        ElementParameters.remove(param)

# create form to select multiple parameters to update

ON_form = forms.SelectFromList.show(ElementParameters, button_name='Select Parameters to turn on', multiselect=True)

# remove all duplicate values from the list

T = Transaction(doc)
T.Start("Apply param val for each family type")
for element in selected_elements:
    for param in element.Parameters:
        if str(param.Definition.Name) in ON_form:
            param.Set(1)
        # elif str(param.Definition.Name) in ElementParameters:
        #     param.Set(0)
T.Commit()

