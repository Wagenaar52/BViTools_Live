import clr

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
selection = uidoc.Selection
view = doc.ActiveView


selected_elementsID = selection.GetElementIds()
selected_elements = [doc.GetElement(eid) for eid in selected_elementsID]

viewT = view.ViewType
leaderLenght = 0.85*view.Scale/100

# get view direction
viewDirection = view.ViewDirection.Normalize()


# yNor = view.UpDirection.Normalize()
# xNor = -viewDirection.CrossProduct(view.UpDirection).Normalize()
# zNor = viewDirection.Normalize()





# t = Transaction(doc)

# t.Start("test1")

# for element in selected_elements:
#     if str(element.HorizontalAlignment) == "Left":
#         element.GetLeaders()[0].Elbow = XYZ(((element.GetLeaders()[0].Anchor.X-leaderLenght)*xNor.X),(element.GetLeaders()[0].Anchor.Y)*yNor.Y,0*zNor.Z)
#     elif str(element.HorizontalAlignment) == "Right":
#         element.GetLeaders()[0].Elbow = XYZ(((element.GetLeaders()[0].Anchor.X+leaderLenght)*xNor.X),(element.GetLeaders()[0].Anchor.Y)*yNor.Y,0*zNor.Z)
            
# t.Commit()   

t = Transaction(doc)

t.Start("Set leader end")

if viewDirection == XYZ(0,0,1) or viewDirection == XYZ(0,0,-1):
    for element in selected_elements:
        if str(element.HorizontalAlignment) == "Left":
            element.GetLeaders()[0].Elbow = XYZ(element.GetLeaders()[0].Anchor.X-leaderLenght,element.GetLeaders()[0].Anchor.Y,0)
        elif str(element.HorizontalAlignment) == "Right":
            element.GetLeaders()[0].Elbow = XYZ((element.GetLeaders()[0].Anchor.X+leaderLenght),element.GetLeaders()[0].Anchor.Y,0)
        
        element.LeaderLeftAttachment = element.LeaderLeftAttachment.TopLine
        element.LeaderRightAttachment = element.LeaderRightAttachment.TopLine
elif viewDirection == XYZ(0,1,0) or viewDirection == XYZ(0,-1,0):       
    for element in selected_elements:
        if str(element.HorizontalAlignment) == "Left":
            element.GetLeaders()[0].Elbow = XYZ(element.GetLeaders()[0].Anchor.X-leaderLenght,0,element.GetLeaders()[0].Anchor.Z)
        elif str(element.HorizontalAlignment) == "Right":
            element.GetLeaders()[0].Elbow = XYZ((element.GetLeaders()[0].Anchor.X+leaderLenght),0,element.GetLeaders()[0].Anchor.Z)
        element.LeaderLeftAttachment = element.LeaderLeftAttachment.TopLine
        element.LeaderRightAttachment = element.LeaderRightAttachment.TopLine
elif viewDirection == XYZ(1,0,0) or viewDirection == XYZ(-1,0,0):
    for element in selected_elements:
        if str(element.HorizontalAlignment) == "Left":
            element.GetLeaders()[0].Elbow = XYZ(0,element.GetLeaders()[0].Anchor.Y,element.GetLeaders()[0].Anchor.Z-leaderLenght)
        elif str(element.HorizontalAlignment) == "Right":
            element.GetLeaders()[0].Elbow = XYZ(0,element.GetLeaders()[0].Anchor.Y,element.GetLeaders()[0].Anchor.Z+leaderLenght)
        element.LeaderLeftAttachment = element.LeaderLeftAttachment.TopLine
        element.LeaderRightAttachment = element.LeaderRightAttachment.TopLine



t.Commit()




# if str(viewT) in planViews:
#     # print("This is a floor plan view.")
#     for element in selected_elements:
#         if str(element.HorizontalAlignment) == "Left":
#             element.GetLeaders()[0].Elbow = XYZ(element.GetLeaders()[0].Anchor.X-leaderLenght,element.GetLeaders()[0].Anchor.Y,0)
#         elif str(element.HorizontalAlignment) == "Right":
#             element.GetLeaders()[0].Elbow = XYZ((element.GetLeaders()[0].Anchor.X+leaderLenght),element.GetLeaders()[0].Anchor.Y,0)
#         element.LeaderLeftAttachment = element.LeaderLeftAttachment.TopLine
#         element.LeaderRightAttachment = element.LeaderRightAttachment.TopLine
# elif str(viewT) in elevationVievs:
#     # print("This is a elevation of section plan view.")
#     for element in selected_elements:
#         if str(element.HorizontalAlignment) == "Left":
#             element.GetLeaders()[0].Elbow = XYZ(element.GetLeaders()[0].Anchor.X-leaderLenght,0,element.GetLeaders()[0].Anchor.Z)
#         elif str(element.HorizontalAlignment) == "Right":
#             element.GetLeaders()[0].Elbow = XYZ((element.GetLeaders()[0].Anchor.X+leaderLenght),0,element.GetLeaders()[0].Anchor.Z)
#         element.LeaderLeftAttachment = element.LeaderLeftAttachment.TopLine
#         element.LeaderRightAttachment = element.LeaderRightAttachment.TopLine
# elif str(viewT) in ISOViews:
#     print("This is an iso view.")
# else:
#     print("This is a different type of view. Check Code")

   
# for element in selected_elements:
#     left_attachment = element.LeaderLeftAttachment
#     right_attachment = element.LeaderRightAttachment
#     print("Left attachment: ", left_attachment)
#     print("Right attachment: ", right_attachment)
#     print("Leader shape: ", element.GetLeaders()[0].LeaderShape)
#     print("Leader elbow: ", element.GetLeaders()[0].Elbow)
#     print("Leader end: ", element.GetLeaders()[0].End)
#     print("Leader anchor: ", element.GetLeaders()[0].Anchor)



# for element in selected_elements:
#     print(element.GetLeaders()[0].Anchor.X)
#     print(element.GetLeaders()[0].LeaderShape)
#     print(element.GetLeaders()[0].HorizontalTextAlignment)

#   AddLeader()element.AddLeader(TextNoteLeaderTypes.TNLT_STRAIGHT_L)
# LeaderLeftAttachment and LeaderRightAttachment indicate the attachment position of the leaders
#  on the corresponding side of the TextNote. Options for the LeaderAttachment are TopLine, MidPoint, and BottomLine.