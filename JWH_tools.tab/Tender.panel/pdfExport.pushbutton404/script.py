from Autodesk.Revit.DB import PDFExportOptions, ElementId
from pyrevit import forms

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document


folder = forms.pick_folder()

project_number = doc.ProjectInformation.Number
project_name = doc.ProjectInformation.Name
tower_supplier = doc.GetElement(ElementId(2106182)).LookupParameter("Tower Supplier").AsString()
turbine_type = doc.GetElement(ElementId(2106182)).LookupParameter("Turbine Type").AsString()
document_name = doc.GetElement(ElementId(426778)).Document.Title
document_rev = document_name[len(document_name)-1]
DWG_name_string = "%s-201-01 - %s - %s - %s Rev%s  " %(project_number ,project_name ,tower_supplier ,turbine_type.replace(' ', '') ,document_rev)
QTY_name_string = "%s - %s - %s - %s Rev%s QTY " %(project_number ,project_name ,tower_supplier ,turbine_type.replace(' ', '') ,document_rev)

pdf_optionsDWG = PDFExportOptions()
pdf_optionsDWG.Combine = False,
pdf_optionsDWG.ZoomType.FitToPage,
pdf_optionsDWG.PaperPlacement.Center,
pdf_optionsDWG.FileName = str(DWG_name_string)


pdf_optionsQTY = PDFExportOptions()
pdf_optionsQTY.Combine = False,
pdf_optionsQTY.PaperFormat.ISO_A4,
pdf_optionsQTY.ZoomType.Zoom,
#pdf_optionsQTY.ZoomPercentage=50,
pdf_optionsQTY.PaperPlacement.Center,
pdf_optionsQTY.FileName = str(QTY_name_string)

pdfOptionsList = [pdf_optionsDWG, pdf_optionsQTY]


listViewsQTY = [ElementId(426778)]
listViewsDWG = [ElementId(2106177)]

doc.Export(folder, listViewsQTY, pdfOptionsList[0])  
doc.Export(folder, listViewsDWG, pdfOptionsList[1])  
