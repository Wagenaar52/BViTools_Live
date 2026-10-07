from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import * 
import Functions as func
from pyrevit import revit, forms, script
import clr
import math
clr.AddReference("System")
clr.AddReference("Microsoft.Office.Interop.Excel")
from System import Int64
from System.Collections.Generic import List
from Microsoft.Office.Interop import Excel

uidoc   = __revit__.ActiveUIDocument
app     = __revit__.Application
doc     = __revit__.ActiveUIDocument.Document

def _cell_value(v):
    """Return numeric cells rounded to 3 decimals, integers as int, text as str."""
    if v is None:
        return None
    try:
        f = float(v)
        return int(f) if f == int(f) else round(f, 3)
    except (TypeError, ValueError):
        return str(v)


schedule_name = "SSI Table"
# path = (r"C:\Users\Wagner.Human\Downloads\WTG_Site_Specific_Informationr1.xlsx")    
headingRows = 4

#prompt user to select excel file
path = forms.pick_file(file_ext="xlsx", title="Select Excel File with SSI Data, must only contain one sheet, otherwise the sheet to be used must be named 'Sheet1'")
schedule_input = forms.ask_for_string(default=schedule_name, prompt="Enter the name of the schedule to update", title="Schedule Name")
heading_input = forms.ask_for_string(default=str(headingRows), prompt="Enter the number of heading rows in the schedule", title="Heading Rows")

if not path:
    script.exit()

if schedule_input:
    schedule_name = schedule_input

try:
    if heading_input:
        headingRows = int(heading_input)
except Exception:
    forms.alert("Heading rows must be an integer. Using default value: {}".format(headingRows))

# Get hold of excel file using ironpython
excel = Excel.ApplicationClass()
excel.Visible = False
excel.DisplayAlerts = False
workbook = None
try:
    workbook = excel.Workbooks.Open(path)
    # if only one sheet, get it by index, otherwise by name
    xl = workbook.Worksheets['Sheet1'] if workbook.Worksheets.Count > 1 else workbook.Worksheets[1]
    # Get hold of cells in excel sheet
    sheetCells = xl.UsedRange.Cells

    sheetValues = [[_cell_value(sheetCells.Item[row, col].Value2) for col in range(1, sheetCells.Columns.Count + 1)] for row in range(1, sheetCells.Rows.Count + 1)]
finally:
    # Always release the Excel process, even if reading fails.
    if workbook is not None:
        workbook.Close(False)
    excel.Quit()

#revome the top  rows from sheetVallues
sheetValues = sheetValues[headingRows:]

stringValues = [[str(cell) for cell in row] for row in sheetValues]
# for row in stringValues:
#     for col in row:
    


for row in stringValues:
    print("|".join(row))


# Get schedule by name
schedule = None
schedules = FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements()
for view in schedules:
    if view.Name == schedule_name:
        schedule = view
        break

if not schedule:
    forms.alert("Schedule '{}' not found.".format(schedule_name))
    script.exit()

#get header of schedule
header = schedule.GetTableData().GetSectionData(SectionType.Header)
header_rows = header.NumberOfRows
header_columns = header.NumberOfColumns

def set_header_cell_text(schedule_view, header_data, row_idx, col_idx, text):
    """Write header text using APIs across Revit versions."""
    try:
        schedule_view.SetCellText(SectionType.Header, row_idx, col_idx, text)
    except Exception:
        header_data.SetCellText(row_idx, col_idx, text)

excel_col_count = len(sheetValues[0]) if sheetValues else 0
excel_row_count = len(sheetValues)

# Data rows are written starting at header row index == headingRows
# (rows 0..headingRows-1 are the existing schedule title/heading rows).
needed_rows = headingRows + excel_row_count
needed_cols = excel_col_count

print("Header rows/cols before: {} x {}".format(header_rows, header_columns))
print("Excel data rows/cols: {} x {}".format(excel_row_count, excel_col_count))

t = Transaction(doc, "Edit SSI Header")
t.Start()
try:
    # Re-fetch the live section data so edits apply inside the transaction.
    table_data = schedule.GetTableData()
    header = table_data.GetSectionData(SectionType.Header)

    # Grow the header so every Excel cell has a real cell to land in.
    # InsertRow/InsertColumn append when given the current count as the index.
    while header.NumberOfColumns < needed_cols:
        header.InsertColumn(header.NumberOfColumns)
    while header.NumberOfRows < needed_rows:
        header.InsertRow(header.NumberOfRows)

    header_rows = header.NumberOfRows
    header_columns = header.NumberOfColumns
    print("Header rows/cols after:  {} x {}".format(header_rows, header_columns))

    rowCount = -1 + headingRows
    for row in sheetValues:
        rowCount += 1
        # Belt-and-braces: never pass an out-of-range index to the table API
        # (that crashes Revit natively, uncatchable in Python).
        if rowCount < 0 or rowCount >= header_rows:
            continue
        colCount = -1
        for col in row:
            colCount += 1
            if colCount < 0 or colCount >= header_columns:
                continue
            try:
                set_header_cell_text(schedule, header, rowCount, colCount, str(col))
            except Exception as e:
                continue
    t.Commit()
except Exception as e:
    if t.HasStarted() and not t.HasEnded():
        t.RollBack()
    forms.alert("Failed to update schedule: {}".format(e), title="SSI Table")
    script.exit()

print("##"* 20)
print("DONE")
