# -*- coding: utf-8 -*-
import clr
import math
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from collections import OrderedDict
from pyrevit import forms
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel


uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
view = doc.GetElement(ElementId(6085764))   

FPath = forms.pick_file(file_ext='xlsx', multi_file=False, unc_paths=False)


excel = Excel.ApplicationClass()
excel.Visible = False
workbook = excel.Workbooks.Open(FPath)
xl = workbook.Worksheets['A']

# get all inputs form excel in a dictionary
BarDict_TR = {}
for i in range(1,200):
    i += 1
    if "TR" in str(xl.Cells(i, 1).Value2).replace(" ","") :
        bar_mark = str(xl.Cells(i,1).Value2)
        rStart = float(xl.Cells(i,2).Value2)/304.8
        rEnd = float(xl.Cells(i,3).Value2)/304.8
        Size = str(xl.Cells(i,4).Value2)
        count = int(xl.Cells(i,8).Value2)
        print(bar_mark)
        BarDict_TR[bar_mark] = [rStart, rEnd, count, Size]
        #print(BarDict_TR[bar_mark])


BarDict_BR = {}
for i in range(1,200):
    i += 1
    if "BR" in str(xl.Cells(i, 1).Value2).replace(" ","") :
        bar_mark = str(xl.Cells(i,1).Value2)
        rStart = float(xl.Cells(i,2).Value2)/304.8
        rEnd = float(xl.Cells(i,3).Value2)/304.8
        Size = str(xl.Cells(i,4).Value2)
        count = int(xl.Cells(i,8).Value2)
        print(bar_mark)
        BarDict_BR[bar_mark] = [rStart, rEnd, count, Size]


BarDict_TC = {}
for i in range(1,200):
    i += 1
    if "TC" in str(xl.Cells(i, 1).Value2).replace(" ","") :
        bar_mark = str(xl.Cells(i,1).Value2)
        rStart = float(xl.Cells(i,2).Value2)/304.8
        rEnd = float(xl.Cells(i,3).Value2)/304.8
        spacing = float(xl.Cells(i,7).Value2)
        Size = str(xl.Cells(i,4).Value2)
        print(bar_mark)
        BarDict_TC[bar_mark] = [rStart, rEnd, spacing, Size]

BarDict_BC = {}
for i in range(1,200):
    i += 1
    if "BC" in str(xl.Cells(i, 1).Value2).replace(" ","") :
        bar_mark = str(xl.Cells(i,1).Value2)
        rStart = float(xl.Cells(i,2).Value2)/304.8
        rEnd = float(xl.Cells(i,3).Value2)/304.8
        Size = str(xl.Cells(i,4).Value2)
        spacing = float(xl.Cells(i,7).Value2)
        print(bar_mark)
        BarDict_BC[bar_mark] = [rStart, rEnd, spacing, Size]


# Get Elements from document

FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
for elem in FEC:
    if elem.Name == '1PA_WTF_SteelTower':
        wtf_steel = elem

rPlinth = wtf_steel.LookupParameter('rPlinth').AsDouble()
rBase = wtf_steel.LookupParameter('rBase').AsDouble()
hPlinth = wtf_steel.LookupParameter('hPlinth').AsDouble()
hBase = wtf_steel.LookupParameter('hBase').AsDouble()
hCone  = wtf_steel.LookupParameter('hCone').AsDouble()
hPit = wtf_steel.LookupParameter('hBottomVoid').AsDouble()
rPitIn = wtf_steel.LookupParameter('rVoidInner').AsDouble()
rPitOut = wtf_steel.LookupParameter('rVoidOuter').AsDouble()
slabSlope = ((hCone-hBase)/(rBase - rPlinth))


FEC_dl = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Lines).WhereElementIsNotElementType().ToElements()
# for elem in FEC_dl:
    # if elem.OwnerViewId == ElementId(6085766):
        # print(elem.Name)
        # doc.Delete(elem.Id)
        # print("Deleted")

# Sort BarDict by 'rEnd' (index 1 in the list)
sorted_BarDict_TR = OrderedDict(sorted(BarDict_TR.items(), key=lambda item: item[1][1]))
sorted_BarDict_BR = OrderedDict(sorted(BarDict_BR.items(), key=lambda item: item[1][1]))
sorted_BarDict_TC = OrderedDict(sorted(BarDict_TC.items(), key=lambda item: item[1][1]))
sorted_BarDict_BC = OrderedDict(sorted(BarDict_BC.items(), key=lambda item: item[1][1]))


FEC_TC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
TC_BM_List = []
for elem in FEC_TC:
    if "TC" in elem.LookupParameter('Schedule Mark').AsString():
        if elem.LookupParameter('Schedule Mark').AsString() not in TC_BM_List:
            TC_BM_List.append(elem.LookupParameter('Schedule Mark').AsString())
TC_BM_List.sort()       

FEC_BC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
BC_BM_List = []
for elem in FEC_BC:
    if "BC" in elem.LookupParameter('Schedule Mark').AsString():
        if elem.LookupParameter('Schedule Mark').AsString() not in BC_BM_List:
            BC_BM_List.append(elem.LookupParameter('Schedule Mark').AsString())
BC_BM_List.sort()




TC_Dict = {}
for i in range(9):
    TC_Dict["TC"+str(i+1)] = []
for bm in TC_BM_List:
        TC_Dict[str(bm)[:3]].append(bm)


BC_Dict = {}
for i in range(9):
    BC_Dict["BC"+str(i+1)] = []
print("**************")
print(BC_Dict)
for bm in BC_BM_List:
        BC_Dict[str(bm)[:3]].append(bm)
print("**************")
print(BC_Dict)
print("**************")





# Get the default text note type
text_note_type_id = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType)

t = Transaction(doc)
t.Start("Section Anotation")

### Draw Center Line
# Define Points for Center Line
cl_top = XYZ(0, 0, hPlinth+3)
cl_bot = XYZ(0, 0, -hPit-2)
# Create Center Line
CL = doc.Create.NewDetailCurve(view, Line.CreateBound(cl_top, cl_bot))
CL.LineStyle = doc.GetElement(ElementId(1018897))
PL = doc.Create.NewDetailCurve(view, Line.CreateBound(XYZ(-rPlinth, 0, hPlinth-1), XYZ(-rPlinth, 0, hPlinth-1.2)))
############################################### TC ###############################################
hcounter = 0
hidLineList = []
for bar_mark, values in sorted_BarDict_TC.items():
    rStart = values[0]
    rEnd = values[1]
    spacing = str(values[2])[:-2]
    Size = values[3]
    # Define Points for Bar
    bar_top_end = XYZ(-rEnd, 0, hCone+1.5)
    if rEnd >= rPlinth:
        bar_bot_end = XYZ(-rEnd, 0, 1 + hBase + (rBase-rEnd)*slabSlope)
    # Create hidden line 
    endHidLine = doc.Create.NewDetailCurve(view, Line.CreateBound(bar_top_end, bar_bot_end))
    endHidLine.LineStyle = doc.GetElement(ElementId(1019189))
    hidLineList.append(endHidLine)
    # Create the dimension
    refArray = ReferenceArray()
    refArray.Append(CL.GeometryCurve.Reference)
    refArray.Append(endHidLine.GeometryCurve.Reference)
    dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(0,0,hPlinth+500/304.8+hcounter),XYZ(rEnd,0,hPlinth+500/304.8+hcounter)), refArray, doc.GetElement(ElementId(1018370)))
    dim.TextPosition = dim.TextPosition + XYZ(0, 0, -0.04750)
    tagBarMark = 0
    subElements = len(TC_Dict[bar_mark[:3]])
    if subElements == 1:
        tagBarMark = bar_mark
    elif subElements > 1:
        tagBarMark = "TC("+str(bar_mark)[2]+"01-"+str(bar_mark)[2]+"0"+str(subElements)+")"
    tagText = str(Size)+"-"+str(tagBarMark)+"-"+str(spacing)
    textXposition = -((rStart+rEnd)/2)-((len(tagText)*35/304.8)/2)
    tag_text_note = TextNote.Create(doc, view.Id, XYZ(textXposition,0,hCone+1.25), tagText, text_note_type_id)
    tag_text_note.TextNoteType = doc.GetElement(ElementId(1018389))
    hcounter += 120/304.8

refArray = ReferenceArray()
refArray.Append(PL.GeometryCurve.Reference)
for line in hidLineList:
    refArray.Append(line.GeometryCurve.Reference)
dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(0,0,hCone+1.5),XYZ(-rBase,0,hCone+500/304.8)), refArray, doc.GetElement(ElementId(1018370)))
# dim.TextPosition = dim.TextPosition + XYZ(0, 0, -0.04750)

############################################### BC ###############################################

hcounter = 0
hidLineList = []
for bar_mark, values in sorted_BarDict_BC.items():
    rStart = values[0]
    rEnd = values[1]
    spacing = str(values[2])[:-2]
    Size = values[3]

    # Define Points for Bar
    bar_bot_end = XYZ(-rEnd, 0, -hPit-0.5)
    if rEnd >= rPlinth:
        bar_top_end = XYZ(-rEnd, 0, -0.25)
    elif rEnd < rPlinth:
        bar_top_end = XYZ(-rEnd, 0, -hPit-0.25)
    # Create hidden line 
    endHidLine = doc.Create.NewDetailCurve(view, Line.CreateBound(bar_top_end, bar_bot_end))
    endHidLine.LineStyle = doc.GetElement(ElementId(1019189))
    hidLineList.append(endHidLine)
    # Create the dimension
    refArray = ReferenceArray()
    refArray.Append(CL.GeometryCurve.Reference)
    refArray.Append(endHidLine.GeometryCurve.Reference)
    dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(0,0,-hPit-1+hcounter),XYZ(rEnd,0,-hPit-1+hcounter)), refArray, doc.GetElement(ElementId(1018370)))
    dim.TextPosition = dim.TextPosition + XYZ(0, 0, -0.04750)
    subElements = len(BC_Dict[bar_mark[:3]])
    tagBarMark
    if subElements == 1:
        tagBarMark = bar_mark
    elif subElements > 1:
        tagBarMark = "BC("+str(bar_mark)[2]+"01-"+str(bar_mark)[2]+"0"+str(subElements)+")"
    tagText = str(Size)+"-"+str(tagBarMark)+"-"+str(spacing)
    textXposition = -((rStart+rEnd)/2)-((len(tagText)*35/304.8)/2)
    tag_text_note = TextNote.Create(doc, view.Id, XYZ(textXposition,0,-hPit-0.75), tagText, text_note_type_id)
    tag_text_note.TextNoteType = doc.GetElement(ElementId(1018389))

    hcounter -= 120/304.8

refArray = ReferenceArray()
refArray.Append(PL.GeometryCurve.Reference)
for line in hidLineList:
    refArray.Append(line.GeometryCurve.Reference)
dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(0,0,-hPit-0.5),XYZ(-rBase,0,-hPit-0.5)), refArray, doc.GetElement(ElementId(1018370)))
# dim.TextPosition = dim.TextPosition + XYZ(0, 0, -0.04750)

############################################### TR ###############################################
hcounter = 0
hScounter = 0
rEndOld = 0
rStartOld = 0
hTagcounter = -120/304.8
hStartTagcounter = -120/304.8

for bar_mark, values in sorted_BarDict_TR.items():
    rStart = values[0]
    rEnd = values[1]
    count = values[2]
    Size = values[3]
    
    angle = str(float(360)/float(count))+"°"

    ####START 
    # Define Points for Bar
    bar_top_start = XYZ(rStart, 0, hPlinth-3)
    if rStart < rPlinth:
        bar_bot_start = XYZ(rStart, 0, hCone-1)
    elif rStart >= rPlinth:
        bar_bot_start = XYZ(rStart, 0, 1 + hBase + (rBase-rStart)*slabSlope)
    # Create hidden line 
    startHidLine = doc.Create.NewDetailCurve(view, Line.CreateBound(bar_top_start, bar_bot_start))
    startHidLine.LineStyle = doc.GetElement(ElementId(1019189))
    # Create the text note
    text_note = TextNote.Create(doc, view.Id, XYZ(bar_top_start.X+0.1,0,hPlinth-3), "START", text_note_type_id)
    text_note.TextNoteType = doc.GetElement(ElementId(1018389))
    # text_note.Coord = XYZ(text_note.Coord.X, text_note.Coord.Y-1, text_note.Coord.Z)
    ####
    if rStart <= rPlinth:
        if rStart != rStartOld:
            hStartTagcounter = hStartTagcounter-120/304.8
        tag_text_note = TextNote.Create(doc, view.Id, XYZ(rPlinth+1,0,hPlinth-3+hStartTagcounter), str(count)+"x"+str(Size)+"-"+str(bar_mark)+"-"+str(angle), text_note_type_id)
        tag_text_note.TextNoteType = doc.GetElement(ElementId(1018389))
        tag_text_note.AddLeader(TextNoteLeaderTypes.TNLT_STRAIGHT_L)
        tag_text_note.GetLeaders()[0].End = XYZ(rStart,0,+hPlinth-3+hStartTagcounter-45/304.8)
        hScounter += 120/304.8
        rStartOld = rStart
    elif rStart > rPlinth:
        tag_text_note = TextNote.Create(doc, view.Id, XYZ(bar_top_start.X+0.25,0,hPlinth-3+hTagcounter), str(count)+"x"+str(Size)+"-"+str(bar_mark)+"-"+str(angle), text_note_type_id)
        tag_text_note.TextNoteType = doc.GetElement(ElementId(1018389))
        tag_text_note.AddLeader(TextNoteLeaderTypes.TNLT_STRAIGHT_L)
        tag_text_note.GetLeaders()[0].End = XYZ(bar_top_start.X,0,hPlinth-3+hTagcounter-45/304.8)
        hScounter += 120/304.8
        rStartOld = rStart

    ####END 
    ####
    # Define Points for Bar
    bar_top_end = XYZ(rEnd, 0, hPlinth-3)
    if rEnd < rPlinth:
        bar_bot_end = XYZ(rEnd, 0, hCone-1)
    elif rEnd >= rPlinth:
        bar_bot_end = XYZ(rEnd, 0, 1 + hBase + (rBase-rEnd)*slabSlope)
    # Create hidden line 
    endHidLine = doc.Create.NewDetailCurve(view, Line.CreateBound(bar_top_end, bar_bot_end))
    endHidLine.LineStyle = doc.GetElement(ElementId(1019189))
    if rEnd != rEndOld:
        # Create the text note
        text_note = TextNote.Create(doc, view.Id, XYZ(bar_top_end.X-0.53,0,hPlinth-3),"END", text_note_type_id)
        text_note.TextNoteType = doc.GetElement(ElementId(1018389))
        # Create the dimension
        refArray = ReferenceArray()
        refArray.Append(CL.GeometryCurve.Reference)
        refArray.Append(endHidLine.GeometryCurve.Reference)
        dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(0,0,hPlinth+500/304.8+hcounter),XYZ(rEnd,0,hPlinth+500/304.8+hcounter)), refArray, doc.GetElement(ElementId(1018370)))
        dim.TextPosition = dim.TextPosition + XYZ(0, 0, -0.04750)
    if rEnd == rEndOld:
        hTagcounter = hTagcounter-120/304.8
    tag_text_note = TextNote.Create(doc, view.Id, XYZ(bar_top_end.X-2.25,0,hPlinth-3+hTagcounter), str(count)+"x"+str(Size)+"-"+str(bar_mark)+"-"+str(angle), text_note_type_id)
    tag_text_note.TextNoteType = doc.GetElement(ElementId(1018389))
    tag_text_note.AddLeader(TextNoteLeaderTypes.TNLT_STRAIGHT_R)
    tag_text_note.GetLeaders()[0].End = XYZ(bar_top_end.X,0,+hPlinth-3+hTagcounter-45/304.8)
    hcounter += 120/304.8
    rEndOld = rEnd

############################################### BR ###############################################
hcounter = -120/304.8
rEndOld = 0
hTagcounter = -120/304.8

for bar_mark, values in sorted_BarDict_BR.items():
    rStart = values[0]
    rEnd = values[1]
    count = values[2]
    Size = values[3]
    
    angle = str(float(360)/float(count))+"°"

    ####START 
    # Define Points for Bar
    bar_bot_start = XYZ(rStart, 0, -hPit-1)
    if rStart < rPlinth:
        bar_top_start = XYZ(rStart, 0, 700/304.8)
    elif rStart >= rPlinth:
        bar_top_start = XYZ(rStart, 0, -1)
    # Create hidden line 
    startHidLine = doc.Create.NewDetailCurve(view, Line.CreateBound(bar_top_start, bar_bot_start))
    startHidLine.LineStyle = doc.GetElement(ElementId(1019189))
    # Create the text note
    text_note = TextNote.Create(doc, view.Id, XYZ(bar_top_start.X+0.1,0,-hPit -0.1), "START", text_note_type_id)
    text_note.TextNoteType = doc.GetElement(ElementId(1018389))
    # Create the dimension
    refArray = ReferenceArray()
    refArray.Append(CL.GeometryCurve.Reference)
    refArray.Append(startHidLine.GeometryCurve.Reference)
    if rStart <= rPlinth:
        dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(0,0,700/304.8-hcounter),XYZ(rStart,0,700/304.8-hcounter)), refArray, doc.GetElement(ElementId(1018370)))
        dim.TextPosition = dim.TextPosition + XYZ(0, 0, -0.04750)
        hStartTagcounter = -120/304.8
        if rEnd == rEndOld:
            hStartTagcounter = hStartTagcounter-120/304.8
        tag_text_note = TextNote.Create(doc, view.Id, XYZ(bar_top_start.X+0.25,0,-hPit-0.25+hStartTagcounter+45.5/304.8), str(count)+"x"+str(Size)+"-"+str(bar_mark)+"-"+str(angle), text_note_type_id)
        tag_text_note.TextNoteType = doc.GetElement(ElementId(1018389))
        tag_text_note.AddLeader(TextNoteLeaderTypes.TNLT_STRAIGHT_R)
        tag_text_note.GetLeaders()[0].End = XYZ(rStart,0,-hPit-0.25+hStartTagcounter)
        
    ####
    ####END 
    if rEnd != rEndOld:
        ####
        # Define Points for Bar
        bar_bot_end = XYZ(rEnd, 0, -hPit-1+hcounter)
        if rEnd < rPlinth:
            bar_top_end = XYZ(rEnd, 0, 0)
        elif rEnd >= rPlinth:
            bar_top_end = XYZ(rEnd, 0, -1)
        # Create hidden line 
        endHidLine = doc.Create.NewDetailCurve(view, Line.CreateBound(bar_top_end, bar_bot_end))
        endHidLine.LineStyle = doc.GetElement(ElementId(1019189))
        # Create the text note
        text_note = TextNote.Create(doc, view.Id, XYZ(bar_top_end.X-0.53,0,-hPit),"END", text_note_type_id)
        text_note.TextNoteType = doc.GetElement(ElementId(1018389))
        # Create the rebar tag text note
        # Create the dimension
        refArray = ReferenceArray()
        refArray.Append(CL.GeometryCurve.Reference)
        refArray.Append(endHidLine.GeometryCurve.Reference)
        dim = doc.Create.NewDimension(view, Line.CreateBound(XYZ(0,0,-hPit-1+hcounter),XYZ(rEnd,0,-hPit-2-hcounter)), refArray, doc.GetElement(ElementId(1018370)))
        dim.TextPosition = dim.TextPosition + XYZ(0, 0, -0.04750)    
    if rEnd == rEndOld:
        hTagcounter = hTagcounter-120/304.8
    tag_text_note = TextNote.Create(doc, view.Id, XYZ(bar_top_end.X-2.25,0,-hPit+hTagcounter), str(count)+"x"+str(Size)+"-"+str(bar_mark)+"-"+str(angle), text_note_type_id)
    tag_text_note.TextNoteType = doc.GetElement(ElementId(1018389))
    tag_text_note.AddLeader(TextNoteLeaderTypes.TNLT_STRAIGHT_R)
    tag_text_note.GetLeaders()[0].End = XYZ(bar_top_end.X,0,-hPit+hTagcounter-45/304.8)
    hcounter -= 120/304.8
    rEndOld = rEnd




t.Commit()
#close excel
workbook.Close(False)
excel.Quit()

