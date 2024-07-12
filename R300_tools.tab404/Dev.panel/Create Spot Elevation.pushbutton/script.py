import clr
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
import math

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
selection = uidoc.Selection
view = doc.ActiveView

selected_elementsID = selection.GetElementIds()
selected_elements = [doc.GetElement(eid) for eid in selected_elementsID]


FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Grids).WhereElementIsNotElementType().ToElements()


t = Transaction(doc, "Create Spot Elevations")
t.Start()

# Loop through the family instances
for instance in selected_elements: #FEC:
    # Check if the instance is of the desired family
    # if instance.Name == "BVi_2PA_Pre-Cast_Coping":
        # Get the location of the instance
        location = instance.Location

        # Define the origin, bend, and end points for the spot elevation
        origin = location
        bend =  location
        end =  location
                
        # Get the GeometryElement of the element
        geometry_element = instance.get_Geometry(Options())

        # Loop through the geometry objects in the GeometryElement
        for geometry_object in geometry_element:
            # Check if the geometry object is a solid
            if isinstance(geometry_object, Solid):
                # Loop through the faces of the solid
                print("Solid Found")
                for face in geometry_object.Faces:
                    # Create a reference to the face
                    #face_reference = Reference.Create(geometry_element.document,geometry_element.ID) #face.Reference
                    ref = Reference(instance)

                    point = XYZ(-220.095040761, 10.287303403, -0.699106506)

        # Create the spot elevation
                    print("Spot Elevation Create atempted")
                    spot_elevation = doc.Create.NewSpotElevation(view, ref, point, point, point, point , True)

# Commit the transaction
                    
t.Commit()
print("Spot Elevations Created")

#################################

# Import necessary pyRevit and Revit API modules
# from pyrevit import revit, DB
# from pyrevit import forms

# # Function to place a spot elevation for a given element
# def place_spot_elevation(element):
#     # Get the active Revit document
#     doc = revit.doc

#     # Get the active view
#     active_view = doc.ActiveView

#     # Get the location of the element (assuming it has a Location property)
#     location = element.Location.Point

#     # Create a new Spot Elevation without a leader line
#     spot_elevation = doc.Create.NewSpotElevation(active_view, DB.Reference(), location, location, False)

#     # Display a success message
#     forms.alert("Spot Elevation placed successfully!", title="Success")

# # Get the selected element(s) using pyRevit's selection prompt
# selection = revit.get_selection()

# # Check if an element is selected
# if selection:
#     # Iterate through selected elements and place spot elevation for each
#     for element_id in selection:
#         element = revit.doc.GetElement(element_id)
#         place_spot_elevation(element)
# else:
#     # Display a message if no element is selected
#     forms.alert("Please select a generic element to place a spot elevation.", title="Error")
