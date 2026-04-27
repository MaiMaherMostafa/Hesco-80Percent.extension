# -*- coding: utf-8 -*-
"""
PyRevit | Title Block Automation + Excel QA Report
IronPython 2.7 Compatible (Excel COM via .NET Interop)
"""

from pyrevit import revit, DB, forms
import os
import clr

# ---------------------------------------------------------
# Excel COM (.NET Interop) – REQUIRED FOR IRONPYTHON
# ---------------------------------------------------------
clr.AddReference("Microsoft.Office.Interop.Excel")
from Microsoft.Office.Interop import Excel

doc = revit.doc

# =========================================================
# CONFIGURATION
# =========================================================

# -------- PART PLAN GROUP 1 (BS & GF) --------
PART_GROUP_1 = {
    "PART 01": "TK_B1XX_BS&GF Part Plan01",
    "PART 02": "TK_B1XX_BS&GF Part Plan02",
    "PART 03": "TK_B1XX_BS&GF Part Plan03",
    "PART 04": "TK_B1XX_BS&GF Part Plan04",
    "PART 05": "TK_B1XX_BS&GF Part Plan05",
    "PART 06": "TK_B1XX_BS&GF Part Plan06",
}
PART_GROUP_1_FLOORS = [
    "FOUNDATION",
    "GROUND FLOOR",
    "FIRST FLOOR",
    "FIRST  FLOOR",
]

# -------- PART PLAN GROUP 2 (TYPICAL) --------
PART_GROUP_2 = {
    "PART 01": "TK_B1XX_Typical Part Plan01",
    "PART 02": "TK_B1XX_Typical Part Plan02",
}
PART_GROUP_2_FLOORS = [
    "SECOND FLOOR",
    "THIRD FLOOR",
    "FOURTH FLOOR",
    "FIFTH FLOOR",
    "SIXTH FLOOR",
    "ROOF FLOOR",
    "UPPER ROOF FLOOR",
]

ALL_GROUP_1_PARAMS = set(PART_GROUP_1.values())
ALL_GROUP_2_PARAMS = set(PART_GROUP_2.values())
ALL_PART_PARAMS = ALL_GROUP_1_PARAMS | ALL_GROUP_2_PARAMS

# -------- KEY PLANS --------
KEY_PLAN_PARAMS = {
    "FOUNDATION": "TK_B1XX_BS Key Plan",
    "GROUND FLOOR": "TK_B1XX_GR_Key Plan",
    "FIRST FLOOR": "TK_B1XX_1ST Key Plan",
    "SECOND FLOOR": "TK_B1XX_TYPICAL Key Plan",
    "THIRD FLOOR": "TK_B1XX_TYPICAL Key Plan",
    "FOURTH FLOOR": "TK_B1XX_4TH Key Plan",
    "FIFTH FLOOR": "TK_B1XX_5TH Key Plan",
    "SIXTH FLOOR": "TK_B1XX_6TH Key Plan",
    "ROOF FLOOR": "TK_B1XX_RF&UR  KEY PLAN",
    "UPPER ROOF FLOOR": "TK_B1XX_RF&UR  KEY PLAN",
}
ALL_KEY_PLAN_PARAMS = set(KEY_PLAN_PARAMS.values())

# -------- ALWAYS ON / OFF --------
ALWAYS_ON_PARAMS = [
    "TK_B01 Site PLAN",
    "TK_PROJECT NORTH",
    "TK_TRUE NORTH",
]

ALWAYS_OFF_PARAMS = [
    "TK_B02 Site PLAN",
]

# -------- REPORT STORAGE --------
RESULTS = []

# =========================================================
# FUNCTIONS
# =========================================================

def get_titleblocks(sheet):
    return list(
        DB.FilteredElementCollector(doc, sheet.Id)
        .OfCategory(DB.BuiltInCategory.OST_TitleBlocks)
        .WhereElementIsNotElementType()
    )

def set_yesno_param(element, name, value):
    p = element.LookupParameter(name)
    if p and not p.IsReadOnly:
        p.Set(1 if value else 0)

def sheet_has_any(name, keywords):
    return any(k in name for k in keywords)

# ---------------------------------------------------------
# Excel COM Export (IronPython SAFE)
# ---------------------------------------------------------
def export_excel(results):
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    filepath = os.path.join(desktop, "TitleBlock_QA_Report.xlsx")

    excel = Excel.ApplicationClass()
    excel.Visible = True

    wb = excel.Workbooks.Add()
    ws = wb.Worksheets[1]
    ws.Name = "Title Block Results"

    headers = [
        "Sheet Number",
        "Sheet Name",
        "Part Group Used",
        "Active Part Parameter",
        "Active Key Plan",
        "Notes",
    ]

    # Write headers
    for col in range(len(headers)):
        cell = ws.Cells[1, col + 1]
        cell.Value2 = headers[col]
        cell.Font.Bold = True

    # Write rows
    row = 2
    for r in results:
        ws.Cells[row, 1].Value2 = r["sheet_number"]
        ws.Cells[row, 2].Value2 = r["sheet_name"]
        ws.Cells[row, 3].Value2 = r["part_group"]
        ws.Cells[row, 4].Value2 = r["active_part"]
        ws.Cells[row, 5].Value2 = r["active_key_plan"]
        ws.Cells[row, 6].Value2 = r["notes"]
        row += 1

    ws.Columns.AutoFit()
    wb.SaveAs(filepath)

    return filepath

# =========================================================
# MAIN
# =========================================================

sheets = forms.select_sheets(
    title="Select Sheets",
    button_name="Apply Title Block Rules"
)

if not sheets:
    forms.alert("No sheets selected.", exitscript=True)

with revit.Transaction("Update Title Block Parameters"):
    for sheet in sheets:
        name = sheet.Name.upper()

        # -------- PART GROUP DETECTION --------
        part_group = None
        active_part_param = None

        if sheet_has_any(name, PART_GROUP_1_FLOORS):
            part_group = 1
            for k, p in PART_GROUP_1.items():
                if k in name:
                    active_part_param = p
                    break

        elif sheet_has_any(name, PART_GROUP_2_FLOORS):
            part_group = 2
            for k, p in PART_GROUP_2.items():
                if k in name:
                    active_part_param = p
                    break

        # -------- KEY PLAN --------
        active_key_plan = None
        for k, p in KEY_PLAN_PARAMS.items():
            if k in name:
                active_key_plan = p
                break

        for tb in get_titleblocks(sheet):

            for p in ALWAYS_ON_PARAMS:
                set_yesno_param(tb, p, True)

            for p in ALWAYS_OFF_PARAMS:
                set_yesno_param(tb, p, False)

            if part_group == 1 and active_part_param:
                for p in ALL_GROUP_1_PARAMS:
                    set_yesno_param(tb, p, p == active_part_param)
                for p in ALL_GROUP_2_PARAMS:
                    set_yesno_param(tb, p, False)

            elif part_group == 2 and active_part_param:
                for p in ALL_GROUP_2_PARAMS:
                    set_yesno_param(tb, p, p == active_part_param)
                for p in ALL_GROUP_1_PARAMS:
                    set_yesno_param(tb, p, False)

            else:
                for p in ALL_PART_PARAMS:
                    set_yesno_param(tb, p, False)

            if active_key_plan:
                for p in ALL_KEY_PLAN_PARAMS:
                    set_yesno_param(tb, p, p == active_key_plan)

        # -------- REPORT --------
        RESULTS.append({
            "sheet_number": sheet.SheetNumber,
            "sheet_name": sheet.Name,
            "part_group":
                "Group 1 (BS & GF)" if part_group == 1 else
                "Group 2 (Typical)" if part_group == 2 else
                "None",
            "active_part": active_part_param or "None",
            "active_key_plan": active_key_plan or "None",
            "notes":
                "No PART detected" if part_group and not active_part_param else
                "No Part Group applied" if not part_group else
                ""
        })

# =========================================================
# EXPORT EXCEL
# =========================================================

file_path = export_excel(RESULTS)

forms.alert(
    "✅ Title Blocks updated successfully\n\n"
    "📊 Excel report created:\n{}".format(file_path),
    title="PyRevit"
)