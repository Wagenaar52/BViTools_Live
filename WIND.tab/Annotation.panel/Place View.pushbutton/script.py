from Autodesk.Revit.DB import  *
from pyrevit.forms import alert
uidoc   = __revit__.ActiveUIDocument
app     = __revit__.Application
doc     = __revit__.ActiveUIDocument.Document
rvt_year = int(app.VersionNumber)

def convert_cm_to_internal(length):
    """Function to convert cm to internal units."""
    # RVT >= 2022
    if rvt_year < 2022:
        from Autodesk.Revit.DB import DisplayUnitType
        return UnitUtils.Convert(length,
                                DisplayUnitType.DUT_CENTIMETERS,
                                DisplayUnitType.DUT_DECIMAL_FEET)
    # RVT >= 2022
    else:
        from Autodesk.Revit.DB import UnitTypeId
        return UnitUtils.ConvertToInternalUnits(length, UnitTypeId.Centimeters)

def get_selected_elements():
    """Property that retrieves selected views or promt user to select some from the dialog box."""
    # GET SELECTED ELEMENTS IN UI
    selected_elements = [doc.GetElement(el_id) for el_id in uidoc.Selection.GetElementIds()]
    return selected_elements

__controls__ = """
ADJUST THESE VALUES AS YOU WANT. 
You can use Positive, Negative and Zero values"""
TOP    = convert_cm_to_internal(0)
BOTTOM = convert_cm_to_internal(-10)
RIGHT  = convert_cm_to_internal(0)
LEFT   = convert_cm_to_internal(0)

WTF = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
for element in WTF:
    if element.Name == "1PA_WTF_SteelTower":
        WTF = element


radius = WTF.LookupParameter("rBase").AsDouble()

if __name__ == '__main__':
    # GET SELECTED VIEWPORTS
    selected           = get_selected_elements()
    selected_viewports = [i for i in selected if type(i) == Viewport]

    # VERIFY THAT VIEWPORTS WERE SELECTED
    if not selected_viewports:
        alert("No ViewPorts were selected.\nPlease, try again.", exitscript=True)


    # START TRANSACTION
    with Transaction(doc,'change viewport') as t:
        t.Start()

        # LOOP THROUGH SELECTED VIEWPORTS
        for vp in selected_viewports:
            # vp.CropBoxActive = True #FIXME This might give an error if View has ScopeBox

            # GET VIEW CROPBOX
            view_id = vp.ViewId
            view    = doc.GetElement(view_id)
            view_bb = view.CropBox

            # CREATE NEW BOUNDING BOX
            BB = BoundingBoxXYZ()
            BB.Min = XYZ(0   , 0   , 0)#6000/304.8)
            BB.Max = XYZ(radius +2  , radius +2    , 0)# -1000/304.8)

            # APPLY NEW BOUNDING BOX
            view.CropBox = BB

        t.Commit()