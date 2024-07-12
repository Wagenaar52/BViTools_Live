from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document


# Get Elements from document


found = doc.GetElement(ElementId(1087568))


MTO = doc.GetElement(ElementId(2257920))

# Define parameters and values to be inserted


f1_base = found.GetMaterialArea(ElementId(1017985), usePaintMaterial = False)
f1_plinth = found.GetMaterialArea(ElementId(2096835), usePaintMaterial = False)



print(f1_base)
print(f1_plinth)


# get hold of MTO parameters


# F1 = MTO.LookupParameter("F1 BaseV")
# F2 = MTO.LookupParameter("F2 PlinthV")
# U1 = MTO.LookupParameter("U1 SlabS")
# U2 = MTO.LookupParameter("U2 PlinthH")



# t = Transaction(doc)

# t.Start("Apply param val")





# t.Commit()



