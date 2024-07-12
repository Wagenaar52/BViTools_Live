from Autodesk.Revit.DB.Structure import *
from Autodesk.Revit.DB import Transaction, Structure, FilteredElementCollector
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, Line, XYZ
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


def new_point(p_1, x, y, z, x_dir, y_dir):
    new_point_1 = p_1 + x*x_dir
    new_point_2 = new_point_1 + y*y_dir
    return(XYZ(new_point_2.X, new_point_2.Y, new_point_2.Z + z))


all_rebar_types = FilteredElementCollector(doc) \
    .OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsElementType() \
    .ToElements()

for  rebar_type in all_rebar_types:
    rebar_name = rebar_type.get_Parameter(BuiltInParameter \
        .SYMBOL_NAME_PARAM).AsString()
    if rebar_name == 'Y32':
        bar_type = rebar_type
        break

WTF = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
for element in WTF:
    if element.Name == "1PA_WTF_SteelTower":
        WTF = element
        break


type_id = WTF.GetTypeId()
type = doc.GetElement(type_id)
w_width = WTF.LookupParameter("rBase").AsDouble()
w_height = WTF.LookupParameter("rTower").AsDouble()
locPoint = WTF.Location.Point


p_1 = locPoint
p_2 = locPoint + XYZ.BasisY

# get the curve form the two points
curve = Line.CreateBound(p_1, p_2)

cc_ext = WTF.LookupParameter("Rebar Cover").AsElementId()
#print(doc.GetElement(WTF.get_Parameter(BuiltInParameter \
#    .CLEAR_COVER_EXTERIOR).AsElementId()))

#print(cc_ext)

# cover = doc.GetElement(cc_ext).Name()
cc_ext = 50/304.8

# print(cover)

direction_y = curve.Direction
direction_x = direction_y.CrossProduct(XYZ.BasisZ).Normalize()
direction_z = XYZ.BasisZ


x_offset = w_width/2 - cc_ext - 10 / 2
y_offset = 50/304.8

rebar_p_1 = new_point(p_1, x_offset+1000/304.8, y_offset , 2000/304.8, direction_x, direction_y)
rebar_p_2 = new_point(rebar_p_1, 0, 0,w_height + 1000/304.8, direction_x, direction_y)
lines = [Line.CreateBound(rebar_p_1, rebar_p_2)]


#get rebar Shape in project in FilteredElementCollector

rebar_shape = FilteredElementCollector(doc).OfClass(RebarShape).WhereElementIsElementType().ToElements()   


for r_shape in rebar_shape:
    if str(r_shape.ShapeFamilyId) == '388762':
        sc_37 = r_shape
        #print(sc_37)
        break

A = 1000/304.8
B = 1000/304.8
C = 1000/304.8
D = 1000/304.8

step = 200/304.8
length = p_1.DistanceTo(p_2)
count = length / step
with Transaction(doc, "Reinforce") as t:
    t.Start()
    # rebar = Structure.Rebar.CreateFromCurves(doc,
    #     Structure.RebarStyle.Standard, bar_type, None, None, WTF, direction_y,
    #     lines, Structure.RebarHookOrientation.Right,
    #     Structure.RebarHookOrientation.Left, True, True)

    rebar = Structure.Rebar.CreateFromRebarShape(doc, sc_37, bar_type, WTF, rebar_p_1,-direction_y, -direction_z)

    rebar.LookupParameter("A").Set(A)
    rebar.LookupParameter("B").Set(B)
    rebar.LookupParameter("C").Set(C)
    rebar.LookupParameter("D").Set(D)
   # rebar.get_Parameter(BuiltInParameter.REBAR_ELEM_LAYOUT_RULE).Set(3)
    # rebar.get_Parameter(BuiltInParameter.REBAR_ELEM_BAR_SPACING).Set(step)
    # rebar.get_Parameter(BuiltInParameter.REBAR_ELEM_QUANTITY_OF_BARS).Set(count)
    # rebar.GetShapeDrivenAccessor().BarsOnNormalSide = True

    t.Commit()