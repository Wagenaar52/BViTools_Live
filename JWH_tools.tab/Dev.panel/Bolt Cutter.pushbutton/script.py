from Autodesk.Revit.DB import Transaction, Structure, FilteredElementCollector,BuiltInCategory, SolidSolidCutUtils
from pyrevit import forms


doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
view = doc.ActiveView

#import SolidSolidCutUtils class


#select the bolts
WTFgrout = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
for element in WTFgrout:
    if element.Name == "1PA_WTF_Grout":
        WTFgrout = element
        break

#select the bolts
WTFTower = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
for element in WTFTower:
    if element.Name == "1PA_WTF_SteelTower":
        WTFsteeltower = element
        break

# print("Grout selected: " + str(WTFgrout.Name))

boltList = []

WTFbolts = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
for element in WTFbolts:
    if element.Name == "2PA_WTF_AnchorBolt_sandbox":
        boltList.append(element)

stoolList = []

WTFstools = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
for element in WTFstools:
    if element.Name == "1PA_AnchorCage_Stool-BearingPlate":
        stoolList.append(element)

# print("Bolts selected: " + str(len(boltList)))
# print("Bolts: " + str(boltList))

t = Transaction(doc, "Cut bolts in grout")
t.Start()

#remove cut between the bolts and the grout
for bolt in boltList:
    SolidSolidCutUtils.RemoveCutBetweenSolids(doc, WTFgrout, bolt)
    SolidSolidCutUtils.RemoveCutBetweenSolids(doc, WTFsteeltower, bolt)

    SolidSolidCutUtils.AddCutBetweenSolids(doc, WTFgrout, bolt)
    SolidSolidCutUtils.AddCutBetweenSolids(doc, WTFsteeltower, bolt)

for stool in stoolList:
    SolidSolidCutUtils.RemoveCutBetweenSolids(doc, WTFsteeltower, stool)

    SolidSolidCutUtils.AddCutBetweenSolids(doc, WTFsteeltower, stool)

t.Commit()
