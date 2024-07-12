from Autodesk.Revit.DB import Transaction, Structure, FilteredElementCollector, ElementId, ElementParameterFilter, ParameterValueProvider, FilterStringRule, FilterStringEquals
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, Line, XYZ
from pyrevit import forms

#  import revit python wrapper module 
from rpw import doc, uidoc, DB, UI

import math
from Autodesk.Revit.DB import ElementId

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
selection = uidoc.Selection
view = doc.ActiveView

changeList = ['ovBottom','ovTop','ohTop','ohBottom', 'sBarrier',	'lBarrier',	'sPanel',	'sPanel2',	'hCoping',	'hCopingLip',	'hCopingKink',	'wCopingLip',	'wCopingKink',	'wCopingSeat',	'wCopingTop',	'dChannel',	'ovChannel',	'wFooting',	'ovFooting',	'ohFooting',	'dFooting',	'wMassConcrete',	'ohPannel',	'wPanel',	'ovPanelTop',	'ovPanelBottom',	'dPanel',	'ovBarrier',	'ohBarrier',	'wChannel',	'wChannelWall1',	'wChannelWall2',	'hChannelWall1',	'hChannelWall2',	'ovCoping',		'ohChannelEnd',	'ohChannelStart',	'ohCopingEnd',	'ohCopingStart',	'olCopingKinkTermination']
print(changeList)
fromParam = {}
#get all parameters from selected element

element = doc.GetElement(selection.GetElementIds()[0])
parameters = element.Parameters

print(parameters)



for parameter in parameters: 
    if parameter.Definition.Name in changeList:
        #append dictory fromParam with parameter name and value
        fromParam[parameter.Definition.Name] = parameter.AsValueString()


print(fromParam)

# prompt user to select element to copy parameters to

toElement = selection.PickObject(UI.Selection.ObjectType.Element, "Select element to copy parameters to").ElementId

print(str(toElement) + "_______code done")