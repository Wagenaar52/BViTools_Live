import clr
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import forms
from rpw import revit, db, ui, DB, UI  
import clr
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import *

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

selection = uidoc.Selection
view = doc.ActiveView

selected_elementsID = selection.GetElementIds()
selected_elements = [doc.GetElement(eid) for eid in selected_elementsID]

ParameterList = [ 'ovRow7.12', 'ovRow7.11', 'ovRow7.10', 'ovRow7.9', 'ovRow7.8', 'ovRow7.7', 'ovRow7.6', 'ovRow7.5', 'ovRow7.4', 'ovRow7.3', 'ovRow7.2', 'ovRow7.1']

#  Get hold of parameters for each selected family type

#var =  [0,1,2,3,4,5,6,7,8,9,10,11,12]
start = 0
end = 13

# val = UnitUtils.ConvertToInternalUnits(300, DisplayUnitType.DUT_MILLIMETERS)


T = Transaction(doc)
T.Start("Apply param val for each family type")
for element in selected_elements:
    for param in element.Parameters:
        if param.Definition.Name in ParameterList:
            if param.AsDouble() < 0.984252:
                element.LookupParameter(param.Definition.Name).Set(0.984252)
                element.LookupParameter("ovSoilNailBotMax").Set(float(element.LookupParameter("ovSoilNailBotMax").AsDouble())+0.984252)
                # print(float(element.LookupParameter("ovSoilNailBotMax").AsDouble())+0.984252)
        

T.Commit()
