import clr
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import forms

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

#get hould of excel file using ironpython

FPath = forms.pick_file(file_ext='xlsx', multi_file=False, unc_paths=False)

clr.AddReference("Microsoft.Office.Interop.Excel")

import Microsoft.Office.Interop.Excel as Excel
excel = Excel.ApplicationClass()
excel.Visible = False
workbook = excel.Workbooks.Open(FPath)
xl = workbook.Worksheets['SheetDATA']


#FOR EACH FAMILY TYPE IN THE MODEL, GET THE PARAMETERS AND UPDATE THEM WITH THE VALUES FROM THE EXCEL FILE



#  Get hold of parameters for each family

#  Get hold of elements in the model

WTFac = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
for element in WTFac:
    if element.Name == "1PA_AnchorCage_Assembly 2":
        WTFac = element
        break

WTFgr = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
for element in WTFgr:
    if element.Name == "1PA_WTF_Grout":
        WTFgr = element
        break

wtf_steel = doc.GetElement(ElementId(1087568))
grout = WTFgr
anchor_cage = WTFac
blinding = doc.GetElement(ElementId(1097087))
duct = doc.GetElement(ElementId(1897290))
backfill = doc.GetElement(ElementId(2342235))



#  Get hold of parameters for each family
PA_WTF_Blinding =['Volume', 'Area', 'hBottomVoid', 'rBase', 'hBlinding', 'rBlindingExtension', 'rBottomVoidInner', 'rBottomVoidOuter']
WTF_SteelTower = ['Volume', 'Area', 'hPlinth', 'hBase', 'hBottomVoid', 'hCone', 'rBase', 'rPlinth', 'rTower', 'dChamfer', 'wGroutTop', 'wGroutBottom', 'dGroutMiddle', 'dGroutSides', 'rSlope', 'rVoidOuter', 'rVoidInner']
PA_Vestas_Duct_Assembly = ['rBase']
PA_WTF_Backfill = ['Volume', 'Area', 'hBase', 'hNGL', 'hCone', 'hPlinth', 'orExcavationBottom', 'ovPlinthToOverburden', 'rBase', 'rPlinth']#, 'sExcavation', 'sOverburdenFill', 'sOverburdenTop']
PA_WTF_Grout = ['Volume', 'Area', 'rTower', 'dGroutMiddle', 'dGroutSide', 'wGroutTop', 'wGroutBottom']
PA_AnchorCage_Assembly = ['placehold1','placehold2','ovBoltTop', 'ovBoltBot', 'lBolt', 'tFlangeTop', 'ovTOCtoBOTtopFlange', 'rBoltInner', 'rBoltOuter', 'nBolts', 'hShell', 'rChamfer', 'rTower', 'tShell', 'wFlange', 'vBearingPlate', 'nBolt', 'nSupStud', 'nSupStuds', 'dSupStudHole', 'rSupStudPath', 'tFlangeBot', 'ovSleaveTop', 'ovSleaveBot', 'dBoltHoleTop', 'dBoltHoleBot', 'vSupportStudHoles', 'wFlangeBot', 'tBearingPlate', 'WBearingPlate']
integerParameter = ['nBolts', 'nSupStud', 'nSupStuds', 'nBolt', 'vBearingPlate', 'vSupportStudHoles']
fam_list = [wtf_steel, grout, anchor_cage, blinding, duct, backfill]
par_list = [WTF_SteelTower, PA_WTF_Grout, PA_AnchorCage_Assembly, PA_WTF_Blinding, PA_Vestas_Duct_Assembly, PA_WTF_Backfill]
col_list = [4, 6, 8, 10, 14, 12]

######change code to update parameters for each family type in the model

fam_list = [grout, anchor_cage]
par_list = [PA_WTF_Grout, PA_AnchorCage_Assembly]
col_list = [6, 8]

########################################################################

t = Transaction(doc)

t.Start("Apply param val for each family type")

j = 0
for fam in fam_list:
    i = 5   
    plist = par_list[j] 

    for item in plist[2:]:
        p = fam.LookupParameter(item)
        pVal = xl.Cells(i,col_list[j]).Value2
        i = i+1
        if  item in integerParameter:
            p.Set(int(pVal))
            
        else:

            p.Set(float(pVal)/304.8)

    j = j+1




print("All Family Parameters Updated in General Arangement Drawing")
# Close Excel application object
workbook.Close(False)
excel.Quit()
print("Excel Closed")
t.Commit()