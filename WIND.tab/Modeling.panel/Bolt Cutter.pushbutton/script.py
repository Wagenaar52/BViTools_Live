from Autodesk.Revit.DB import Transaction, Structure, FilteredElementCollector,BuiltInCategory, SolidSolidCutUtils, FamilyInstance
from pyrevit import forms


doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
view = doc.ActiveView

#import SolidSolidCutUtils class


def is_valid_cutter(element):
    """Check if element can be used as a cutting solid"""
    try:
        if not isinstance(element, FamilyInstance):
            return False
        if element.Symbol is None:
            return False
        return True
    except:
        return False


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

WTFlange = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
for element in WTFlange:
    if element.Name == "1PA_AnchorCage_BotFlange 2":
        botFlange = element
        elId = botFlange.Id
        break

WTFBearingPlate = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
for element in WTFBearingPlate:
    if element.Name == "1PA_AnchorCage_BotFlange 2" and element.Id != elId:
        bearingPlate = element
        break

WTFlange = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
for element in WTFlange:
    if element.Name == "1PA_TowerFlangeSweep":
        towerFlange = element
        break

voidList = []

WTFvoid = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
for element in WTFvoid:
    if element.Name == "2PA_AnchorCage_SupBoltVoid":
        voidList.append(element)
        break

t = Transaction(doc, "Cut bolts in grout")
t.Start()

#remove cut between the bolts and the grout
for bolt in boltList:
    if not is_valid_cutter(bolt):
        continue
    try:
        SolidSolidCutUtils.RemoveCutBetweenSolids(doc, WTFgrout, bolt)
        SolidSolidCutUtils.RemoveCutBetweenSolids(doc, WTFsteeltower, bolt)
        SolidSolidCutUtils.RemoveCutBetweenSolids(doc, botFlange, bolt)
        SolidSolidCutUtils.RemoveCutBetweenSolids(doc, bearingPlate, bolt)
        SolidSolidCutUtils.RemoveCutBetweenSolids(doc, towerFlange, bolt)

        SolidSolidCutUtils.AddCutBetweenSolids(doc, WTFgrout, bolt)
        SolidSolidCutUtils.AddCutBetweenSolids(doc, WTFsteeltower, bolt)
        SolidSolidCutUtils.AddCutBetweenSolids(doc, botFlange, bolt)
        SolidSolidCutUtils.AddCutBetweenSolids(doc, bearingPlate, bolt)
        SolidSolidCutUtils.AddCutBetweenSolids(doc, towerFlange, bolt)
    except Exception as e:
        print("Error cutting bolt {}: {}".format(bolt.Name if hasattr(bolt, 'Name') else 'Unknown', str(e)))

for stool in stoolList:
    if not is_valid_cutter(stool):
        continue
    try:
        SolidSolidCutUtils.RemoveCutBetweenSolids(doc, WTFsteeltower, stool)
        SolidSolidCutUtils.RemoveCutBetweenSolids(doc, botFlange, stool)
        SolidSolidCutUtils.RemoveCutBetweenSolids(doc, botFlange, stool)

        SolidSolidCutUtils.AddCutBetweenSolids(doc, WTFsteeltower, stool)
        SolidSolidCutUtils.AddCutBetweenSolids(doc, botFlange, stool)
        SolidSolidCutUtils.AddCutBetweenSolids(doc, botFlange, stool)
    except Exception as e:
        print("Error cutting stool {}: {}".format(stool.Name if hasattr(stool, 'Name') else 'Unknown', str(e)))

for void in voidList:
    if not is_valid_cutter(void):
        continue
    try:
        SolidSolidCutUtils.RemoveCutBetweenSolids(doc, botFlange, void)

        SolidSolidCutUtils.AddCutBetweenSolids(doc, botFlange, void)
    except Exception as e:
        print("Error cutting void {}: {}".format(void.Name if hasattr(void, 'Name') else 'Unknown', str(e)))

t.Commit()
