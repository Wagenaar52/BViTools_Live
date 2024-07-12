import clr
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import forms


uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

#get hould of excel file using ironpython


FPath = forms.pick_file(file_ext='xlsx', multi_file=False, unc_paths=False)

# r"""c:\Users\Wagner.Human\Desktop\Tender Drawing Template\IMPORT_FILEREV4.xlsx"""
clr.AddReference("Microsoft.Office.Interop.Excel")

import Microsoft.Office.Interop.Excel as Excel
excel = Excel.ApplicationClass()
excel.Visible = False
workbook = excel.Workbooks.Open(FPath)
xl = workbook.Worksheets['SheetDATA']

FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericAnnotation).WhereElementIsNotElementType().ToElements()

for elem in FEC:
    if elem.Name == "WTF_LoadsTable":
        loadsTable = elem
        break
    # else:
    #     print("No WTF_LoadsTable found")

    # loadsTable = doc.GetElement(ElementId(1229727))
paraList = ['cS1_Fz', 'cS1_Fxy', 'cS1_Mxy','cS1_Mz', 'cS2_Fz', 'cS2_Fxy', 'cS2_Mxy', 'cS2_Mz', 'cS3_Fxy', 'cS3_Fz', 'cS3_Mxy', 'cS3_Mz', 'cN_Fxy', 'cN_Fz', 'cN_Mxy', 'cN_Mz', 'cA_Fz', 'cA_Fxy', 'cA_Mxy', 'cA_Mz', 'cMz_Fz', 'cMz_Fxy', 'cMz_Mz', 'cMz_Mxy']


# Transaction

t = Transaction(doc)
t.Start("Apply param val")

   
for i in range(5, 28):

    for para in loadsTable.Parameters:
        if str(xl.Cells(i,1).Value2) == para.Definition.Name:
            p = loadsTable.LookupParameter(str(xl.Cells(i,1).Value2))
            pVal = xl.Cells(i,2).Value2
            if "_M" in para.Definition.Name:
                p.Set(float(pVal)*10763.9104167097)
            elif "_F" in para.Definition.Name:
                p.Set(float(pVal)*3280.83989501312)
            i = i+1


t.Commit()
print("Loads Table Imported")
excel.Quit()
#AsDouble()/3280.839895