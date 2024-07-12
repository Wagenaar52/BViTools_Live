from Autodesk.Revit.DB.Structure import * 
from Autodesk.Revit.DB.Structure import RebarShape
import math, clr
from Autodesk.Revit.DB import Transaction, Structure, FilteredElementCollector, RadialArray, ArrayAnchorMember
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, Line, XYZ
from Autodesk.Revit.DB import FailureSeverity, FailureProcessingResult,IFailuresPreprocessor
from pyrevit import forms
from Autodesk.Revit.DB import *
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel





doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
view = doc.ActiveView


FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
scheduleMarkList = []
for elem in FEC:
    if elem.LookupParameter("Schedule Mark").AsString() not in scheduleMarkList:
        scheduleMarkList.append(elem.LookupParameter("Schedule Mark").AsString())


# t = Transaction(doc, "Update A")    
# t.Start()
# for i in range(len(scheduleMarkList)):
#     if "BC" in scheduleMarkList[i]:
#         sum_A = 0
#         count_A = 0
#         A_max = 0
#         for elem in FEC:
#             if elem.LookupParameter("Schedule Mark").AsString() == scheduleMarkList[i]:
#                 sum_A += elem.LookupParameter("A").AsDouble()
#                 count_A += 1
#                 if elem.LookupParameter("A").AsDouble() > A_max:
#                     A_max = elem.LookupParameter("A").AsDouble()
#         A = round((sum_A/count_A)*100)/100
#         A = round(A*304.8)/304.8
#         print(scheduleMarkList[i])
#         print(A*304.8)
#         print(A_max*304.8)

#         for elem in FEC:
#             if elem.LookupParameter("Schedule Mark").AsString() == scheduleMarkList[i]:
#                 elem.LookupParameter("A").Set(((round((A_max*304.8)/100)*100)/304.8))



# t.Commit()

t = Transaction(doc, "Update r")
t.Start()
for i in range(len(scheduleMarkList)):
    if "TC" in scheduleMarkList[i]:
        sum_r = 0
        count_r = 0
        for elem in FEC:
            if elem.LookupParameter("Schedule Mark").AsString() == scheduleMarkList[i]:
                sum_r += elem.LookupParameter("r").AsDouble()
                count_r += 1
        r = round((sum_r/count_r)*100)/100
        r = round(r*304.8)/304.8
        print(scheduleMarkList[i])
        print(r*304.8)
        for elem in FEC:
            if elem.LookupParameter("Schedule Mark").AsString() == scheduleMarkList[i]:
                elem.LookupParameter("Rebar r Custom").Set(r)
        print("updated r")
    if "TC" in scheduleMarkList[i]:
        sum_r = 0
        count_r = 0
        for elem in FEC:
            if elem.LookupParameter("Schedule Mark").AsString() == scheduleMarkList[i]:
                sum_r += elem.LookupParameter("r").AsDouble()
                count_r += 1
        r = round((sum_r/count_r)*100)/100
        r = round(r*304.8)/304.8
        print(scheduleMarkList[i])
        print(r*304.8)
        for elem in FEC:
            if elem.LookupParameter("Schedule Mark").AsString() == scheduleMarkList[i]:
                elem.LookupParameter("Rebar r Custom").Set(r)
        print("updated r")


t.Commit()

