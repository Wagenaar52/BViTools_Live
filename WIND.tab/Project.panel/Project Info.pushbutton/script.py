import clr
import math
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import forms


uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

FPath = forms.pick_file(file_ext='xlsx', multi_file=False, unc_paths=False)

#get hould of excel file using ironpython
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
excel = Excel.ApplicationClass()
excel.Visible = False
workbook = excel.Workbooks.Open(FPath)
xl = workbook.Worksheets['ProjectInfo']


Organization_Name           =    str(xl.Cells(2,2).Value2)
Organization_Description    =    str(xl.Cells(3,2).Value2)
Building_Name               =    str(xl.Cells(4,2).Value2)
Project_Status              =    str(xl.Cells(5,2).Value2)
Author                      =    str(xl.Cells(6,2).Value2)
Project_Issue_Date          =    str(xl.Cells(7,2).Value2)
Client_Name                 =    str(xl.Cells(8,2).Value2)
Project_Address             =    str(xl.Cells(9,2).Value2)
Project_Name                =    str(xl.Cells(10,2).Value2)
Project_Number              =    str(xl.Cells(11,2).Value2)


t = Transaction(doc, "Update Project Information")
t.Start()

prInfo = doc.ProjectInformation
prInfo.OrganizationName = Organization_Name
prInfo.OrganizationDescription = Organization_Description
prInfo.BuildingName = Building_Name
prInfo.Status = Project_Status
prInfo.Author = Author
prInfo.IssueDate = Project_Issue_Date
prInfo.ClientName = Client_Name
prInfo.Address = Project_Address
prInfo.Name = Project_Name
prInfo.Number = Project_Number

t.Commit()

print("Project Information Updated")
# Close Excel application object
workbook.Close(False)
excel.Quit()
print("Excel Closed")






