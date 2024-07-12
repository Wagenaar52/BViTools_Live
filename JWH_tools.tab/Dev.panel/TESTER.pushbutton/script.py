
import clr
import math
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB.Structure import RebarShape

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
selection = uidoc.Selection
view = doc.ActiveView


##snipet##### how to get shape family ############################


# rebar_shape = FilteredElementCollector(doc).OfClass(RebarShape).WhereElementIsElementType().ToElements()  
# for r_shape in rebar_shape:
#     if r_shape.LookupParameter("Type Name").AsString() == '20':
#         sc_20 = r_shape

#         print("shap shap")
##snipet############################################################################################################



##draw model line between two points

# p1 = XYZ(21.929133858, 1.331328908, 6.817116610)
# p2 = XYZ(20.236220472, -1.331328908, 0.164041995)

# line = Line.CreateBound(p1, p2)

# t = Transaction(doc, 'Create Model Line')
# t.Start()
# model_line = doc.Create.NewModelCurve(line, SketchPlane.Create(doc, Plane.CreateByThreePoints(p1, p2, XYZ(0,0,0))))
# t.Commit()


##snipet################################################################################################################

# FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
# t = Transaction(doc, 'Center Elements')
# t.Start('Center Elements')
# for elem in FEC:
#     if elem.Name == "1PA_AnchorCage_Assembly 2":
#         print("1PA_AnchorCage_Assembly 2")
#         print(elem.Location.Point)
#         elem.Pinned = False
#         elem.Location.Point = XYZ(0,0,elem.Location.Point.Z)
#         elem.Pinned = True
#         print(elem.Location.Point)
#     elif elem.Name == "1PA_WTF_Grout":
#         print("1PA_WTF_Grout")
#         print(elem.Location.Point)
#         elem.Pinned = False
#         elem.Location.Point = XYZ(0,0,elem.Location.Point.Z)
#         elem.Pinned = True
#         print(elem.Location.Point)
#     elif elem.Name == "1PA_WTF_Blinding":
#         print("1PA_WTF_Blinding")
#         print(elem.Location.Point)
#         elem.Pinned = False
#         elem.Location.Point = XYZ(0,0,elem.Location.Point.Z)
#         elem.Pinned = True
#         print(elem.Location.Point)
#     elif elem.Name == "1PA_WTF_SteelTower":
#         print("1PA_WTF_SteelTower")
#         print(elem.Location.Point)
#         elem.Pinned = False
#         elem.Location.Point = XYZ(0,0,elem.Location.Point.Z)
#         elem.Pinned = True
#         print(elem.Location.Point)


# t.Commit()




FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
scheduleMarkList = []
for elem in FEC:
    if elem.LookupParameter("Schedule Mark").AsString() not in scheduleMarkList:
        scheduleMarkList.append(elem.LookupParameter("Schedule Mark").AsString())

t = Transaction(doc, "Update A")    
t.Start()
for i in range(len(scheduleMarkList)):
    if "TC" in scheduleMarkList[i]:
        sum_A = 0
        count_A = 0
        A_max = 0
        for elem in FEC:
            if elem.LookupParameter("Schedule Mark").AsString() == scheduleMarkList[i]:
                sum_A += elem.LookupParameter("A").AsDouble()
                count_A += 1
                if elem.LookupParameter("A").AsDouble() > A_max:
                    A_max = elem.LookupParameter("A").AsDouble()
        A = round((sum_A/count_A)*100)/100
        A = round(A*304.8)/304.8
        print(scheduleMarkList[i])
        print(A*304.8)
        print(A_max*304.8)

        for elem in FEC:
            if elem.LookupParameter("Schedule Mark").AsString() == scheduleMarkList[i]:
                elem.LookupParameter("A").Set(((round((A_max*304.8)/100)*100)/304.8))



t.Commit()



































# # Rebar script for review

# # -*- coding: utf-8 -*-
# import Autodesk
# from Autodesk.Revit.DB import Transaction, Structure, FilteredElementCollector
# from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, Line, XYZ
# doc = __revit__.ActiveUIDocument.Document
# uidoc = __revit__.ActiveUIDocument


# def new_point(p_1, x, y, z, x_dir, y_dir):
#     new_point_1 = p_1 + x*x_dir
#     new_point_2 = new_point_1 + y*y_dir
#     return(XYZ(new_point_2.X, new_point_2.Y, new_point_2.Z + z))


# all_rebar_types = FilteredElementCollector(doc) \
#     .OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsElementType() \
#     .ToElements()

# for  rebar_type in all_rebar_types:
#     rebar_name = rebar_type.get_Parameter(BuiltInParameter \
#         .SYMBOL_NAME_PARAM).AsString()
#     if rebar_name == '10 A500':
#         bar_type = rebar_type
#         break

# wall = [doc.GetElement( elId ) for elId in uidoc.Selection.GetElementIds()][0]

# type_id = wall.GetTypeId()
# type = doc.GetElement(type_id)
# w_width = type.Width
# w_height = wall.get_Parameter(BuiltInParameter \
#     .WALL_USER_HEIGHT_PARAM).AsDouble()


# curve = wall.Location.Curve
# p_1 = curve.GetEndPoint(0)
# p_2 = curve.GetEndPoint(1)

# cc_ext = doc.GetElement(wall.get_Parameter(BuiltInParameter \
#     .CLEAR_COVER_EXTERIOR).AsElementId()).CoverDistance
# print(doc.GetElement(wall.get_Parameter(BuiltInParameter \
#     .CLEAR_COVER_EXTERIOR).AsElementId()))

# cc_int = doc.GetElement(wall.get_Parameter(BuiltInParameter \
#     .CLEAR_COVER_INTERIOR).AsElementId()).CoverDistance

# direction_y = curve.Direction
# direction_x = direction_y.CrossProduct(XYZ.BasisZ).Normalize()

# x_offset = w_width/2 - cc_ext - bar_type.BarDiameter / 2
# y_offset = 50/304.8

# rebar_p_1 = new_point(p_1, x_offset, y_offset , 0, direction_x, direction_y)
# rebar_p_2 = new_point(rebar_p_1, 0, 0,w_height + 1000/304.8, direction_x, direction_y)
# lines = [Line.CreateBound(rebar_p_1, rebar_p_2)]

# step = 200/304.8
# length = p_1.DistanceTo(p_2)
# count = length / step
# with Transaction(doc, "Reinforce") as t:
#     t.Start()
#     rebar = Structure.Rebar.CreateFromCurves(doc,
#         Structure.RebarStyle.Standard, bar_type, None, None, wall, direction_y,
#         lines, Structure.RebarHookOrientation.Right,
#         Structure.RebarHookOrientation.Left, True, True)

#     rebar.get_Parameter(BuiltInParameter.REBAR_ELEM_LAYOUT_RULE).Set(3)
#     rebar.get_Parameter(BuiltInParameter.REBAR_ELEM_BAR_SPACING).Set(step)
#     rebar.get_Parameter(BuiltInParameter.REBAR_ELEM_QUANTITY_OF_BARS).Set(count)
#     rebar.GetShapeDrivenAccessor().BarsOnNormalSide = True

#     t.Commit()