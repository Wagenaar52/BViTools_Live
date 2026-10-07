import clr
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import forms
import Functions as func
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
wtf_steel    = func.getWTFElements(xl)[0]
grout        = func.getWTFElements(xl)[1]
anchor_cage  = func.getWTFElements(xl)[2]
blinding     = func.getWTFElements(xl)[3]
# duct         = func.getWTFElements(xl)[5]
backfill     = func.getWTFElements(xl)[4]

print(wtf_steel.Name+"  IMPORTED")
print(grout.Name+"  IMPORTED")
print(anchor_cage.Name+"  IMPORTED")
print(blinding.Name+"  IMPORTED")
# print(duct.Name+"  IMPORTED")
print(backfill.Name+"  IMPORTED")
print("###################################################")


fam_list = [wtf_steel, grout, anchor_cage, blinding,  backfill]#,duct,]

#  Get hold of parameters for each family
PA_WTF_Blinding =['Volume', 'Area', 'hBottomVoid', 'rBase', 'hBlinding', 'rBlindingExtension', 'rBottomVoidInner', 'rBottomVoidOuter']
WTF_SteelTower = ['Volume', 'Area', 'hPlinth', 'hBase', 'hBottomVoid', 'hCone', 'rBase', 'rPlinth', 'rTower', 'dChamfer', 'wGroutTop', 'wGroutBottom', 'dGroutMiddle', 'dGroutSides', 'rSlope', 'rVoidOuter', 'rVoidInner']
PA_Vestas_Duct_Assembly = ['rBase']
PA_WTF_Backfill = ['Volume', 'Area', 'hBase', 'hNGL', 'hCone', 'hPlinth', 'orExcavationBottom', 'ovPlinthToOverburden', 'rBase', 'rPlinth']#, 'sExcavation', 'sOverburdenFill', 'sOverburdenTop']
PA_WTF_Grout = ['Volume', 'Area', 'rTower', 'dGroutMiddle', 'dGroutSide', 'wGroutTop', 'wGroutBottom']
PA_AnchorCage_Assembly = ['placehold1','placehold2','ovBoltTop', 'ovBoltBot', 'lBolt', 'tFlangeTop', 'ovTOCtoBOTtopFlange', 'rBoltInner', 'rBoltOuter', 'nBolts', 'hShell', 'rChamfer', 'rTower', 'tShell', 'wFlange', 'vBearingPlate', 'nBolt', 'nSupStud', 'nSupStuds', 'dSupStudHole', 'rSupStudPath', 'tFlangeBot', 'ovSleaveTop', 'ovSleaveBot', 'dBoltHoleTop', 'dBoltHoleBot', 'vSupportStudHoles', 'wFlangeBot', 'tBearingPlate', 'WBearingPlate']
integerParameter = ['nBolts', 'nSupStud', 'nSupStuds', 'nBolt', 'vBearingPlate', 'vSupportStudHoles']
par_list = [WTF_SteelTower, PA_WTF_Grout, PA_AnchorCage_Assembly, PA_WTF_Blinding, PA_WTF_Backfill] #PA_Vestas_Duct_Assembly, 
col_list = [4, 6, 8, 10, 12] #, 12]

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
            try:
                p.Set(int(pVal))
                print(str(fam.Name) + "| Parameter: "+ str(item) + "   |   "+str(pVal) + " ---- INTEGER " )
            except:
                print("************************************************************************************")
                print("**********************NOT SET*******************************************************------- integer parameter :  " + str(fam.Name) + "| Parameter: "+ str(item) + "   |   "+str(pVal))
                print("************************************************************************************")
        else:
            try:
                p.Set(float(pVal)/304.8)
                print(str(fam.Name) + "| Parameter: "+ str(item) + "   |   "+str(pVal))
            except:
                print("************************************************************************************")
                print("**********************NOT SET*******************************************************------- float parameter :  " + str(fam.Name) + "| Parameter: "+ str(item) + "   |   "+str(pVal))
                print("************************************************************************************")
    j = j+1
    #print("*******{}parameters updated*******".format(fam.Name))
t.Commit()

#Move Grout in place   
t = Transaction(doc)
t.Start("Move Grout in place")

grout.Location.Point = wtf_steel.Location.Point + XYZ(0,0,wtf_steel.LookupParameter('hPlinth').AsDouble())
print("Grout Moved in place")

t.Commit()  



print("All Family Parameters Updated in General Arangement Drawing")
# Close Excel application object
workbook.Close(False)
excel.Quit()
print("Excel Closed")
