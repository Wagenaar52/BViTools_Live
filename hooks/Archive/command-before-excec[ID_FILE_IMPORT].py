from pyrevit import revit, EXEC_PARAMS
from Autodesk.Revit.UI import TaskDialog


#Variables
sender =__eventsender__ 
args =__eventargs__

doc = revit.doc
doc = args.Document

if not doc.IsFamilyDocument:
    TaskDialog.Show("dont do that", "This is not a family document.")


