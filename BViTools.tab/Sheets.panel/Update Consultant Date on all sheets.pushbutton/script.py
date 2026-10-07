# -*- coding: utf-8 -*-
__title__ = "title"
__doc__ = """
_____________________________________________________________________
Description:
This is a template file for pyRevit Scripts.
_____________________________________________________________________"""

# IMPORTS 📚
# ==================================================

from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script
import clr
clr.AddReference("System")
from System.Collections.Generic import List

#  VARIABLES 📄
# ==================================================
doc = revit.doc
uidoc = revit.uidoc
app = revit.HOST_APP



# collect all the sheets in the project and check whether they have a parameter called "BVi_TB_Consult_Date"  , prompt the user to select a date and then populate the parameter with the selected date on all the sheets that have the parameter.
sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
sheets_with_param = []      
for sheet in sheets:
    param = sheet.LookupParameter("BVi_TB_Consult_Date")
    if param:
        sheets_with_param.append(sheet)
if not sheets_with_param:
    forms.alert("No sheets with the parameter 'BVi_TB_Consult_Date' were found.", title="Parameter Not Found")


date = forms.ask_for_string("Enter the consultation date (e.g., 21/11/2025):", title="Input BVi_TB_Consult_Date")
if date:
    with Transaction(doc, "Set Consultation Date") as t:
        t.Start()
        for sheet in sheets_with_param:
            param = sheet.LookupParameter("BVi_TB_Consult_Date")
            if param:
                param.Set(date)
        t.Commit()
    forms.alert("Consultation date has been set for all applicable sheets.", title="Success")
else:
    forms.alert("No date entered. Operation cancelled.", title="Cancelled")

