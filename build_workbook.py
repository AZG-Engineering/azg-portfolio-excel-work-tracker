"""Build "Work Tracker Demo.xlsx" with Excel itself. The result has no macros.

Usage (Windows, with Excel installed):
    python build_workbook.py

What it does:
  1. Starts a private, hidden Excel and builds the four sheets: Dashboard, Tasks,
     How to use, Lists.
  2. Saves the workbook, then rewrites the file's properties so the author and
     "last modified by" read "AZG Engineering" and nothing about this PC is stored.
  3. Copies the clean file to deliverable\\.

All task data is made up. Dates are set relative to the day this is run, so that
some tasks are overdue and some are due this week. Run it again for fresh dates.

Every formula is a classic one (SUMPRODUCT, AGGREGATE, INDEX, MATCH, COUNTIF,
IFERROR), so the workbook uses formulas that exist in Excel 2010 and later.

The look follows the AZG build style guide: its colours, Calibri, no gridlines,
tiles with a coloured left bar, list headers on a pale band, teal progress bars
beside their numbers, and a word on every coloured flag.
"""

from __future__ import annotations

import datetime as dt
import re
import shutil
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
WORKBOOK = ROOT / "Work Tracker Demo.xlsx"
DELIVERABLE = ROOT / "deliverable" / WORKBOOK.name
AUTHOR = "AZG Engineering"
ALL = "(All)"

# How many rows each dashboard list shows.
FOCUS_ROWS = 8
NEXT_ROWS = 5
VIEW_ROWS = 15
OWNER_ROWS = 8

HEADERS = ("ID", "Task", "Owner", "Department", "Status", "Priority", "Start", "Due", "% Complete", "Year", "Notes")
STATUSES = ("Not started", "In progress", "Blocked", "Done")
PRIORITIES = ("High", "Medium", "Low")

# Made-up tasks: task, owner, department, status, priority,
# start and due as days from the build date, % complete, notes.
TASKS = (
    ("Renew office lease", "Taylor R.", "Operations", "Done", "High", -60, -35, 100, "Signed for 3 years"),
    ("Q3 month-end checklist", "Jordan M.", "Finance", "Done", "High", -45, -28, 100, ""),
    ("Update price list", "Casey L.", "Sales", "Done", "Medium", -40, -21, 100, "Sent to all reps"),
    ("Replace front-desk PC", "Riley P.", "IT", "Done", "Low", -30, -14, 100, ""),
    ("New-hire paperwork review", "Morgan T.", "HR", "Done", "Medium", -25, -9, 100, ""),
    ("Vendor contract renewal", "Taylor R.", "Operations", "In progress", "High", -30, -12, 70, "Waiting on vendor redline"),
    ("Customer survey mailing", "Casey L.", "Sales", "In progress", "Medium", -28, -8, 60, ""),
    ("Backup system test", "Riley P.", "IT", "Blocked", "High", -20, -6, 40, "Needs a new drive"),
    ("Expense policy update", "Jordan M.", "Finance", "In progress", "Medium", -21, -4, 80, ""),
    ("Safety training sign-off", "Morgan T.", "HR", "Not started", "High", -10, -2, 0, ""),
    ("Forklift inspection", "Avery K.", "Operations", "In progress", "High", -9, -1, 50, ""),
    ("Trade show booth order", "Casey L.", "Sales", "Blocked", "Medium", -18, -15, 30, "Supplier out of stock"),
    ("Payroll calendar 2027", "Jordan M.", "Finance", "In progress", "Medium", -7, 0, 90, ""),
    ("Warehouse rack labels", "Avery K.", "Operations", "In progress", "Low", -6, 1, 75, ""),
    ("Quote follow-up calls", "Casey L.", "Sales", "Not started", "High", -2, 2, 0, ""),
    ("Laptop refresh order", "Riley P.", "IT", "In progress", "Medium", -12, 3, 55, ""),
    ("Benefits enrollment memo", "Morgan T.", "HR", "In progress", "High", -8, 5, 35, ""),
    ("Fleet insurance renewal", "Taylor R.", "Operations", "Not started", "High", -1, 6, 0, ""),
    ("Monthly sales report", "Casey L.", "Sales", "Done", "Medium", -10, 2, 100, "Finished early"),
    ("Budget draft for 2027", "Jordan M.", "Finance", "In progress", "High", -14, 7, 45, ""),
    ("Website contact form fix", "Riley P.", "IT", "Not started", "Medium", 0, 8, 0, ""),
    ("Holiday schedule posting", "Morgan T.", "HR", "Not started", "Low", 1, 9, 0, ""),
    ("Dock door repair", "Avery K.", "Operations", "Blocked", "High", -5, 10, 20, "Part on order"),
    ("Key account review", "Casey L.", "Sales", "In progress", "High", -3, 12, 25, ""),
    ("Supplier scorecards", "Taylor R.", "Operations", "In progress", "Medium", -10, 13, 40, ""),
    ("Password policy rollout", "Riley P.", "IT", "In progress", "High", -7, 16, 30, ""),
    ("Collections call list", "Jordan M.", "Finance", "Not started", "Medium", 3, 18, 0, ""),
    ("Performance review forms", "Morgan T.", "HR", "In progress", "Medium", -4, 21, 15, ""),
    ("Year-end inventory count", "Avery K.", "Operations", "Not started", "High", 20, 27, 0, ""),
    ("Holiday promo flyer", "Casey L.", "Sales", "In progress", "Medium", -2, 30, 20, ""),
    ("Phone system upgrade", "Riley P.", "IT", "Blocked", "Medium", -15, 35, 10, "Carrier quote pending"),
    ("Year-end document binder", "Jordan M.", "Finance", "Not started", "High", 10, 42, 0, ""),
    ("First-aid kit restock", "Morgan T.", "HR", "Done", "Low", -12, 14, 100, ""),
    ("Delivery route review", "Avery K.", "Operations", "In progress", "Medium", -6, 48, 10, ""),
    ("CRM cleanup", "Casey L.", "Sales", "Not started", "Low", 14, 56, 0, ""),
    ("Office wifi upgrade", "Riley P.", "IT", "Done", "Medium", -35, -18, 100, ""),
    ("Quarterly vendor review", "Jordan M.", "Finance", "Done", "High", -20, -5, 100, ""),
    ("New-year filing setup", "Jordan M.", "Finance", "Not started", "High", 60, 95, 0, ""),
    ("Server room cleanup", "Riley P.", "IT", "Not started", "Low", 70, 100, 0, ""),
    ("Annual safety plan", "Taylor R.", "Operations", "Not started", "Medium", 65, 105, 0, ""),
    ("Spring hiring plan", "Morgan T.", "HR", "Not started", "Medium", 80, 118, 0, ""),
    ("Customer appreciation day", "Casey L.", "Sales", "Blocked", "Low", 30, 125, 5, "Venue not confirmed"),
)

# Each table column also gets a name (T_Owner, T_Due, ...) that grows with the
# table. Named formulas can't hold a table reference like Tasks[Owner] when they
# are created by a script, so they use these instead.
COLUMN_NAMES = {"ID": "T_ID", "Task": "T_Task", "Owner": "T_Owner", "Status": "T_Status",
                "Due": "T_Due", "Year": "T_Year"}

# Names (Formulas > Name Manager). The dashboard's formulas are built from these.
NAMES = (
    ("SelOwner", "=Dashboard!$C$5"),
    ("SelStatus", "=Dashboard!$E$5"),
    ("SelYear", "=Dashboard!$H$5"),
    # 1 for the first table row, 2 for the second, ...
    ("RowNo", "=ROW(T_ID)-MIN(ROW(T_ID))+1"),
    # One TRUE/FALSE per table row: does it pass each filter?
    ("OwnerOK", f'=((SelOwner="{ALL}")+(SelOwner="")+(T_Owner=SelOwner))>0'),
    ("StatusOK", f'=((SelStatus="{ALL}")+(SelStatus="")+(T_Status=SelStatus))>0'),
    ("YearOK", f'=((SelYear="{ALL}")+(SelYear="")+(T_Year=SelYear))>0'),
    # One 1/0 per table row: is it a real task that passes all three filters?
    ("InScope", '=(T_Task<>"")*OwnerOK*StatusOK*YearOK'),
    ("CondOverdue", '=InScope*(T_Status<>"Done")*ISNUMBER(T_Due)*(T_Due<TODAY())'),
    ("CondUpcoming", '=InScope*(T_Status<>"Done")*ISNUMBER(T_Due)*(T_Due>=TODAY())'),
    ("CondFocus", "=CondUpcoming*(T_Due<=TODAY()+6)"),
    # Due date plus a tiny bit that says which row it is, so ties sort cleanly.
    ("SortKey", "=IF(ISNUMBER(T_Due),T_Due,99999)+RowNo/100000"),
    # The same, with 100000 added for Done tasks so that open tasks list first.
    ("SortKeyView", '=IF(T_Status="Done",100000,0)+SortKey'),
    # How many owners come before this one alphabetically.
    ("OwnerRank", '=IF(T_Owner="",-1,COUNTIF(T_Owner,"<"&T_Owner))'),
    ("OwnerList", '=Lists!$A$4:INDEX(Lists!$A$4:$A$34,COUNTIF(Lists!$A$4:$A$34,"?*"))'),
    ("StatusList", "=Lists!$C$4:$C$8"),
    ("YearList", "=Lists!$E$4:INDEX(Lists!$E$4:$E$14,1+COUNT(Lists!$E$5:$E$14))"),
)

# Excel constants.
XL_CENTER, XL_LEFT, XL_RIGHT = -4108, -4131, -4152
XL_EDGE_LEFT, XL_EDGE_TOP, XL_EDGE_BOTTOM, XL_EDGE_RIGHT, XL_INSIDE_HORIZONTAL = 7, 8, 9, 10, 12
XL_THIN, XL_MEDIUM, XL_THICK = 2, -4138, 4
XL_VALIDATE_INPUT_ONLY, XL_VALIDATE_LIST, XL_VALIDATE_DECIMAL = 0, 3, 2
XL_CELL_VALUE, XL_EXPRESSION, XL_EQUAL, XL_GREATER, XL_BETWEEN = 1, 2, 3, 5, 1
XL_OPEN_XML_WORKBOOK = 51

# The palette of the AZG build style guide (section 2). No other colours are used.
NAVY, INK, SLATE = "#1F3A5F", "#1F2937", "#5B6B7F"
LINE, SURFACE, WHITE = "#E2E8F0", "#F5F7FA", "#FFFFFF"
TEAL, RED, GREEN = "#0F766E", "#B42318", "#2E7D32"
AMBER_FILL, AMBER_TEXT = "#FDE7B0", "#7A4A00"

# The Dashboard grid. The main list is on the left (three tile widths), the
# supporting blocks on the right (two tile widths); every tile is 36 wide.
WIDTHS = {"A": 2,
          "B": 9, "C": 27,  # tile 1: Due, Task
          "D": 11, "E": 12, "F": 13,  # tile 2: Flag, Owner, Status
          "G": 10, "H": 7, "I": 9, "J": 10,  # tile 3: Priority, %, bar, Department
          "K": 3,
          "L": 10, "M": 26,  # tile 4
          "N": 12, "O": 10, "P": 6, "Q": 8,  # tile 5
          "R": 2}
DASHBOARD_AREA = "$A$1:$R$38"

TILES = (  # label, formula, colour of the left bar and the number, name, label cells, number cells
    ("TOTAL TASKS", "=SUMPRODUCT(InScope)", NAVY, "KpiTotal", "B7:C7", "B8:C9"),
    ("OVERDUE", "=SUMPRODUCT(CondOverdue)", RED, "KpiOverdue", "D7:F7", "D8:F9"),
    ("DUE THIS WEEK", "=SUMPRODUCT(CondFocus)", AMBER_TEXT, "KpiDueThisWeek", "G7:J7", "G8:J9"),
    ("IN PROGRESS", '=SUMPRODUCT(InScope*(Tasks[Status]="In progress"))', NAVY, "KpiInProgress", "L7:M7", "L8:M9"),
    ("DONE", '=SUMPRODUCT(InScope*(Tasks[Status]="Done"))', NAVY, "KpiDone", "N7:Q7", "N8:Q9"),
)

# Where each Dashboard list sits, for the test. "columns" picks the tested fields
# out of the block, in the order the test expects, skipping the progress-bar column.
LAYOUT = {
    # Due, Task, Flag, Owner, Status, Priority, %, bar, Department
    "view": {"cells": "B13:J27", "columns": (0, 1, 3, 4, 5, 6, 2, 8), "more": "B28"},
    # Due, Task, Days left, Owner, Status, Priority, %, bar, Department
    "next": {"cells": "B32:J36", "columns": (0, 1, 3, 4, 5, 6, 2, 8), "more": "B37"},
    # Due, Task, Owner, Priority, %, bar
    "focus": {"cells": "L13:Q20", "columns": (0, 1, 2, 3, 4), "more": "L21"},
    # Owner, bar, Average %, Tasks, Done, Overdue
    "owners": {"cells": "L25:Q32", "columns": (0, 2, 3, 4, 5), "more": "L33"},
}


def rgb(hex_color: str) -> int:
    """'#RRGGBB' as the number Excel uses for a colour."""
    h = hex_color.lstrip("#")
    return int(h[0:2], 16) + (int(h[2:4], 16) << 8) + (int(h[4:6], 16) << 16)


def serial(day: dt.date) -> int:
    """A date as Excel stores it."""
    return (day - dt.date(1899, 12, 30)).days


def task_rows(today: dt.date) -> list[tuple]:
    rows = []
    for number, (task, owner, department, status, priority, start, due, percent, notes) in enumerate(TASKS, start=1):
        rows.append((f"T-{number:03d}", task, owner, department, status, priority,
                     serial(today + dt.timedelta(days=start)), serial(today + dt.timedelta(days=due)),
                     percent / 100, "", notes))
    return rows


# ---- Sheet builders -----------------------------------------------------------

def add_data_bar(cells) -> None:
    """A solid teal bar on a fixed 0% to 100% scale, with no number over it."""
    bar = cells.FormatConditions.AddDatabar()
    bar.MinPoint.Modify(0, 0)  # fixed scale: 0% is empty,
    bar.MaxPoint.Modify(0, 1)  # 100% is a full bar
    bar.BarColor.Color = rgb(TEAL)
    bar.BarFillType = 0  # solid
    bar.BarBorder.Type = 0  # no border
    bar.ShowValue = False  # the % sits in its own column, so the bar never covers it


def formula_rule(cells, formula: str):
    """A conditional-formatting rule driven by a formula."""
    return cells.FormatConditions.Add(XL_EXPRESSION, None, formula)


def color_when_equal(cells, text: str, ink: str, fill: str | None = None, bold: bool = True) -> None:
    rule = cells.FormatConditions.Add(XL_CELL_VALUE, XL_EQUAL, f'="{text}"')
    rule.Font.Color = rgb(ink)
    rule.Font.Bold = bold
    if fill:
        rule.Interior.Color = rgb(fill)


def add_table_style(workbook) -> str:
    """The guide's data-sheet look: navy header, pale input cells, one thin rule per row."""
    style = workbook.TableStyles.Add("AZG Table")
    whole = style.TableStyleElements(0)  # whole table
    whole.Interior.Color = rgb(SURFACE)
    whole.Font.Color = rgb(INK)
    for edge in (XL_INSIDE_HORIZONTAL, XL_EDGE_BOTTOM):
        whole.Borders(edge).LineStyle = 1
        whole.Borders(edge).Color = rgb(LINE)
    header = style.TableStyleElements(1)  # header row
    header.Interior.Color = rgb(NAVY)
    header.Font.Color = rgb(WHITE)
    header.Font.Bold = True
    return style.Name


def build_tasks(sheet, today: dt.date):
    rows = task_rows(today)
    last = 1 + len(rows)
    sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, 11)).Value = (HEADERS,)
    sheet.Range(sheet.Cells(2, 1), sheet.Cells(last, 11)).Value = tuple(rows)
    table = sheet.ListObjects.Add(1, sheet.Range(sheet.Cells(1, 1), sheet.Cells(last, 11)), None, 1)
    table.Name = "Tasks"
    table.TableStyle = add_table_style(sheet.Parent)
    # A calculated column: new rows get their Year from their Due date.
    table.ListColumns("Year").DataBodyRange.Formula = '=IF([@Due]="","",YEAR([@Due]))'
    for column, name in COLUMN_NAMES.items():
        table.ListColumns(column).DataBodyRange.Name = name

    sheet.Range("G:H").NumberFormat = "mmm d, yyyy"
    sheet.Range("I:I").NumberFormat = "0%"
    for column, width in zip("ABCDEFGHIJK", (8, 30, 13, 13, 13, 10, 14, 25, 13, 8, 30)):
        sheet.Columns(column).ColumnWidth = width
    sheet.Range("A:A").HorizontalAlignment = XL_LEFT
    sheet.Range("J:J").HorizontalAlignment = XL_CENTER
    sheet.Rows(1).RowHeight = 22
    sheet.Rows(1).VerticalAlignment = XL_CENTER

    def dropdown(column: str, choices: tuple[str, ...]) -> None:
        validation = table.ListColumns(column).DataBodyRange.Validation
        validation.Delete()
        validation.Add(XL_VALIDATE_LIST, 1, XL_BETWEEN, ",".join(choices))
        validation.ErrorTitle = column
        validation.ErrorMessage = "Pick one of: " + ", ".join(choices)

    dropdown("Status", STATUSES)
    dropdown("Priority", PRIORITIES)
    percent = table.ListColumns("% Complete").DataBodyRange.Validation
    percent.Delete()
    percent.Add(XL_VALIDATE_DECIMAL, 1, XL_BETWEEN, "0", "1")
    percent.ErrorTitle = "% Complete"
    percent.ErrorMessage = "Type a percentage from 0% to 100%."

    # Year is calculated, not typed: left white, in slate italics, with a note
    # that shows when a Year cell is selected. (A cell comment would store the
    # Office user's name in the file, so the note is an input message instead.)
    year = table.ListColumns("Year").DataBodyRange
    year.Interior.Color = rgb(WHITE)
    year.Font.Color = rgb(SLATE)
    year.Font.Italic = True
    year.Validation.Delete()
    year.Validation.Add(XL_VALIDATE_INPUT_ONLY)
    year.Validation.InputTitle = "Year"
    year.Validation.InputMessage = "Calculated: fills itself in from Due. No need to type here."

    # Overdue flags. The rules cover the whole column, so new rows are covered too.
    # Each flag carries a word as well as a colour, added by the rule's number format.
    sheet.Activate()
    sheet.Range("H1").Select()  # rule formulas are read relative to the active cell
    overdue = '=AND(ISNUMBER($H1),$H1<TODAY(),$E1<>"Done")'
    this_week = '=AND(ISNUMBER($H1),$H1>=TODAY(),$H1<=TODAY()+6,$E1<>"Done")'
    rule = formula_rule(sheet.Range("H:H"), overdue)
    rule.Interior.Color = rgb(RED)
    rule.Font.Color = rgb(WHITE)
    rule.Font.Bold = True
    rule.NumberFormat = 'mmm d, yyyy"   OVERDUE"'
    rule = formula_rule(sheet.Range("H:H"), this_week)
    rule.Interior.Color = rgb(AMBER_FILL)
    rule.Font.Color = rgb(AMBER_TEXT)
    rule.NumberFormat = 'mmm d, yyyy"   this week"'

    sheet.Range("A2").Select()
    sheet.Application.ActiveWindow.FreezePanes = True
    sheet.Range("A1").Select()
    return table


def build_lists(sheet) -> None:
    sheet.Range("A1").Value = "Helper lists for the Dashboard. These are formulas: please don't type on this sheet."
    sheet.Range("A1").Font.Bold = True
    sheet.Range("A1").Font.Color = rgb(NAVY)
    sheet.Range("A2").Value = "Dropdown choices on the left; on the right, which table row each dashboard list line shows."
    sheet.Range("A2").Font.Color = rgb(SLATE)

    def heading(address: str, text: str) -> None:
        cell = sheet.Range(address)
        cell.Value = text
        cell.Font.Bold = True
        cell.Font.Color = rgb(NAVY)
        cell.Interior.Color = rgb(LINE)

    heading("A3", "Owner choices")
    sheet.Range("A4").Value = ALL
    sheet.Range("A5:A34").Formula = (
        "=IFERROR(INDEX(Tasks[Owner],MATCH(AGGREGATE(15,6,"
        'OwnerRank/(COUNTIF(A$4:A4,Tasks[Owner])=0)/(Tasks[Owner]<>""),1),OwnerRank,0)),"")'
    )
    heading("C3", "Status choices")
    sheet.Range("C4:C8").Value = tuple((value,) for value in (ALL, *STATUSES))
    heading("E3", "Year choices")
    sheet.Range("E4").Value = ALL
    sheet.Range("E5:E14").Formula = '=IFERROR(AGGREGATE(15,6,Tasks[Year]/(COUNTIF(E$4:E4,Tasks[Year])=0),1),"")'
    sheet.Range("E4:E14").HorizontalAlignment = XL_LEFT

    def picks(first_column: str, title: str, condition: str, count: int, sort_key: str = "SortKey") -> None:
        number, key, row = (chr(ord(first_column) + i) for i in range(3))
        heading(f"{number}3", "#")
        heading(f"{key}3", title)
        heading(f"{row}3", "Table row")
        last = 3 + count
        sheet.Range(f"{number}4:{number}{last}").Value = tuple((n,) for n in range(1, count + 1))
        # The n-th smallest sort key among the rows that match. Rows that don't
        # match divide by zero, and AGGREGATE (option 6) skips errors.
        sheet.Range(f"{key}4:{key}{last}").Formula = f'=IFERROR(AGGREGATE(15,6,{sort_key}/{condition},${number}4),"")'
        # The sort key is the due date plus a tiny bit that says which row it is.
        sheet.Range(f"{row}4:{row}{last}").Formula = f'=IF({key}4="","",ROUND(MOD({key}4,1)*100000,0))'
        sheet.Range(f"{key}4:{key}{last}").NumberFormat = "0.00000"

    picks("G", "This Week's Focus: key", "CondFocus", FOCUS_ROWS)
    picks("K", "Next due: key", "CondUpcoming", NEXT_ROWS)
    picks("O", "Tasks in view: key", "InScope", VIEW_ROWS, sort_key="SortKeyView")

    for column, width in (("A", 22), ("B", 3), ("C", 16), ("D", 3), ("E", 14), ("F", 3), ("G", 5), ("H", 24),
                          ("I", 11), ("J", 3), ("K", 5), ("L", 18), ("M", 11), ("N", 3), ("O", 5), ("P", 20), ("Q", 11)):
        sheet.Columns(column).ColumnWidth = width
    sheet.Tab.Color = rgb(SLATE)


def build_dashboard(sheet, workbook) -> None:
    excel = sheet.Application
    sheet.Activate()
    excel.ActiveWindow.DisplayGridlines = False
    for column, width in WIDTHS.items():  # column A is the gutter
        sheet.Columns(column).ColumnWidth = width
    sheet.Range("A1:R40").VerticalAlignment = XL_CENTER

    # ---- Top band: title, "As of" line, filters. Row 1 is a spacer.
    sheet.Rows(1).RowHeight = 8
    title = sheet.Range("B2")
    title.Value = "Work Tracker Dashboard"
    title.Font.Size = 20
    title.Font.Bold = True
    title.Font.Color = rgb(NAVY)
    sheet.Rows(2).RowHeight = 30
    as_of = sheet.Range("B3")
    as_of.Formula = '="As of "&TEXT(TODAY(),"ddd, mmm d, yyyy")&"   |   Pick an owner, status or year: the whole page follows."'
    as_of.Font.Color = rgb(SLATE)
    sheet.Rows(4).RowHeight = 6

    def filter_cell(label_at: str, label: str, cell_at: str, merge: str | None, source: str) -> None:
        tag = sheet.Range(label_at)
        tag.Value = label
        tag.Font.Bold = True
        tag.Font.Color = rgb(SLATE)
        tag.HorizontalAlignment = XL_RIGHT
        box = sheet.Range(merge or cell_at)
        if merge:
            box.Merge()
        box.Interior.Color = rgb(LINE)  # the light navy-tinted fill that marks an input
        box.Font.Bold = True
        box.Font.Color = rgb(NAVY)
        box.HorizontalAlignment = XL_LEFT
        box.Borders(XL_EDGE_BOTTOM).Color = rgb(NAVY)
        box.Borders(XL_EDGE_BOTTOM).Weight = XL_THIN
        cell = sheet.Range(cell_at)
        cell.Value = ALL
        validation = cell.Validation
        validation.Delete()
        validation.Add(XL_VALIDATE_LIST, 1, XL_BETWEEN, f"={source}")
        validation.ErrorTitle = label
        validation.ErrorMessage = "Pick a value from the list."

    filter_cell("B5", "Owner", "C5", None, "OwnerList")
    filter_cell("D5", "Status", "E5", "E5:F5", "StatusList")
    filter_cell("G5", "Year", "H5", "H5:I5", "YearList")
    hint = sheet.Range("L5")
    hint.Formula = (f'=IF(AND(OR(SelOwner="{ALL}",SelOwner=""),OR(SelStatus="{ALL}",SelStatus=""),'
                    f'OR(SelYear="{ALL}",SelYear="")),"Showing all tasks","Filtered view: choose {ALL} to clear")')
    hint.Font.Italic = True
    hint.Font.Size = 9
    hint.Font.Color = rgb(SLATE)
    sheet.Rows(5).RowHeight = 22
    sheet.Rows(6).RowHeight = 10

    # ---- KPI row: white tiles of equal width, a coloured left bar, label above, number below.
    for label, formula, color, name, label_cells, number_cells in TILES:
        tag = sheet.Range(label_cells)
        tag.Merge()
        tag.Value = label
        tag.Font.Size = 9
        tag.Font.Color = rgb(SLATE)
        number = sheet.Range(number_cells)
        number.Merge()
        number.Formula = formula
        number.Font.Size = 36
        number.Font.Bold = True
        number.Font.Color = rgb(color)
        for part in (tag, number):
            part.HorizontalAlignment = XL_LEFT
            part.IndentLevel = 1
            part.Borders(XL_EDGE_LEFT).Color = rgb(color)
            part.Borders(XL_EDGE_LEFT).Weight = XL_THICK
        workbook.Names.Add(name, "=Dashboard!" + number.Cells(1, 1).Address)
    sheet.Rows(7).RowHeight = 18
    sheet.Rows(8).RowHeight = 26
    sheet.Rows(9).RowHeight = 26
    sheet.Rows(10).RowHeight = 12

    # ---- Shared styles
    def section(address: str, text: str, note_at: str, note: str, as_formula: bool = False) -> None:
        head = sheet.Range(address)
        head.Value = text
        head.Font.Size = 13
        head.Font.Bold = True
        head.Font.Color = rgb(NAVY)
        extra = sheet.Range(note_at)
        if as_formula:
            extra.Formula = note
        else:
            extra.Value = note
        extra.Font.Size = 9
        extra.Font.Italic = True
        extra.Font.Color = rgb(SLATE)

    def header_row(first: str, last: str, row: int, labels: tuple[str, ...]) -> None:
        cells = sheet.Range(f"{first}{row}:{last}{row}")
        cells.Value = (labels,)
        cells.Font.Bold = True
        cells.Font.Color = rgb(NAVY)
        cells.Interior.Color = rgb(LINE)

    def more_line(address: str, count: str, shown: int, what: str = "more not shown") -> None:
        cell = sheet.Range(address)
        cell.Formula = f'=IF({count}>{shown},"+ "&({count}-{shown})&" {what}","")'
        cell.Font.Size = 9
        cell.Font.Italic = True
        cell.Font.Color = rgb(SLATE)

    def lookup(column: str, pick: str, text: bool = True) -> str:
        return f'INDEX(Tasks[{column}],{pick})' + ('&""' if text else "")

    def bar_of(percent_cell: str) -> str:
        """The progress-bar cell: the same number as its % cell, drawn as a bar."""
        return f'=IF({percent_cell}="","",{percent_cell})'

    # ---- Tasks in view: the main list (left, rows 11-28)
    section("B11", "Tasks in view", "E11", '="Open tasks first, then by due date   |   "&KpiTotal&" match the filters"', as_formula=True)
    header_row("B", "J", 12, ("Due", "Task", "Flag", "Owner", "Status", "Priority", "% done", "", "Department"))
    for n in range(VIEW_ROWS):
        row, pick, key = 13 + n, f"Lists!$Q{4 + n}", f"Lists!$P{4 + n}"
        empty = '"No tasks match these filters"' if n == 0 else '""'
        due = f"MOD(INT({key}),100000)"  # takes the "Done" 100000 back off the sort key
        sheet.Range(f"B{row}").Formula = f'=IF({pick}="","",IF({due}=99999,"",{due}))'
        sheet.Range(f"C{row}").Formula = f'=IF({pick}="",{empty},{lookup("Task", pick)})'
        sheet.Range(f"D{row}").Formula = (f'=IF(OR({pick}="",B{row}="",F{row}="Done"),"",'
                                         f'IF(B{row}<TODAY(),"OVERDUE",IF(B{row}<=TODAY()+6,"This week","")))')
        sheet.Range(f"E{row}").Formula = f'=IF({pick}="","",{lookup("Owner", pick)})'
        sheet.Range(f"F{row}").Formula = f'=IF({pick}="","",{lookup("Status", pick)})'
        sheet.Range(f"G{row}").Formula = f'=IF({pick}="","",{lookup("Priority", pick)})'
        sheet.Range(f"H{row}").Formula = f'=IF({pick}="","",{lookup("% Complete", pick, text=False)})'
        sheet.Range(f"I{row}").Formula = bar_of(f"H{row}")
        sheet.Range(f"J{row}").Formula = f'=IF({pick}="","",{lookup("Department", pick)})'
    more_line("B28", "KpiTotal", VIEW_ROWS)

    # ---- Next due dates (left, rows 30-37), on the same columns as the main list
    section("B30", f"Next {NEXT_ROWS} due dates", "E30", "Not done, soonest first")
    header_row("B", "J", 31, ("Due", "Task", "Days left", "Owner", "Status", "Priority", "% done", "", "Department"))
    for n in range(NEXT_ROWS):
        row, pick, key = 32 + n, f"Lists!$M{4 + n}", f"Lists!$L{4 + n}"
        empty = '"Nothing coming up"' if n == 0 else '""'
        sheet.Range(f"B{row}").Formula = f'=IF({pick}="","",INT({key}))'
        sheet.Range(f"C{row}").Formula = f'=IF({pick}="",{empty},{lookup("Task", pick)})'
        sheet.Range(f"D{row}").Formula = f'=IF({pick}="","",INT({key})-TODAY())'
        sheet.Range(f"E{row}").Formula = f'=IF({pick}="","",{lookup("Owner", pick)})'
        sheet.Range(f"F{row}").Formula = f'=IF({pick}="","",{lookup("Status", pick)})'
        sheet.Range(f"G{row}").Formula = f'=IF({pick}="","",{lookup("Priority", pick)})'
        sheet.Range(f"H{row}").Formula = f'=IF({pick}="","",{lookup("% Complete", pick, text=False)})'
        sheet.Range(f"I{row}").Formula = bar_of(f"H{row}")
        sheet.Range(f"J{row}").Formula = f'=IF({pick}="","",{lookup("Department", pick)})'
    more_line("B37", "SUMPRODUCT(CondUpcoming)", NEXT_ROWS)

    # ---- This Week's Focus (right, rows 11-21)
    section("L11", "This Week's Focus", "N11", "Not done, due in the next 7 days")
    header_row("L", "Q", 12, ("Due", "Task", "Owner", "Priority", "% done", ""))
    for n in range(FOCUS_ROWS):
        row, pick, key = 13 + n, f"Lists!$I{4 + n}", f"Lists!$H{4 + n}"
        empty = '"Nothing due in the next 7 days"' if n == 0 else '""'
        sheet.Range(f"L{row}").Formula = f'=IF({pick}="","",INT({key}))'
        sheet.Range(f"M{row}").Formula = f'=IF({pick}="",{empty},{lookup("Task", pick)})'
        sheet.Range(f"N{row}").Formula = f'=IF({pick}="","",{lookup("Owner", pick)})'
        sheet.Range(f"O{row}").Formula = f'=IF({pick}="","",{lookup("Priority", pick)})'
        sheet.Range(f"P{row}").Formula = f'=IF({pick}="","",{lookup("% Complete", pick, text=False)})'
        sheet.Range(f"Q{row}").Formula = bar_of(f"P{row}")
    more_line("L21", "KpiDueThisWeek", FOCUS_ROWS)

    # ---- Progress by owner (right, rows 23-33)
    section("L23", "Progress by owner", "N23", "Follows the filters")
    header_row("L", "Q", 24, ("Owner", "Average % done", "", "Tasks", "Done", "Overdue"))
    for n in range(OWNER_ROWS):
        row, owner = 25 + n, f"Lists!$A{5 + n}"
        sheet.Range(f"L{row}").Formula = f'=IF({owner}="","",{owner})'
        sheet.Range(f"M{row}").Formula = bar_of(f"N{row}")
        sheet.Range(f"N{row}").Formula = (f'=IF(OR(L{row}="",O{row}=0),"",'
                                         f'SUMPRODUCT(InScope*(Tasks[Owner]=L{row})*Tasks[% Complete])/O{row})')
        sheet.Range(f"O{row}").Formula = f'=IF(L{row}="","",SUMPRODUCT(InScope*(Tasks[Owner]=L{row})))'
        sheet.Range(f"P{row}").Formula = f'=IF(L{row}="","",SUMPRODUCT(InScope*(Tasks[Owner]=L{row})*(Tasks[Status]="Done")))'
        sheet.Range(f"Q{row}").Formula = f'=IF(L{row}="","",SUMPRODUCT(CondOverdue*(Tasks[Owner]=L{row})))'
    more_line("L33", 'COUNTIF(Lists!$A$5:$A$34,"?*")', OWNER_ROWS, "more owners not shown")

    # ---- Number formats, alignment, colour
    # The left and right lists share rows, so every row from 11 down is one height.
    sheet.Range("11:38").RowHeight = 18
    for address in ("B13:B27", "B32:B36", "L13:L20"):
        sheet.Range(address).NumberFormat = "mmm d"
        sheet.Range(address).HorizontalAlignment = XL_LEFT
    for address in ("H13:H27", "H32:H36", "P13:P20", "N25:N32"):
        sheet.Range(address).NumberFormat = "0%"
    for address in ("I13:I27", "I32:I36", "Q13:Q20", "M25:M32"):
        add_data_bar(sheet.Range(address))
    sheet.Range("H12,H31,P12").HorizontalAlignment = XL_RIGHT
    sheet.Range("N24:Q24").HorizontalAlignment = XL_CENTER
    sheet.Range("N25:Q32").HorizontalAlignment = XL_CENTER
    sheet.Range("D32:D36").NumberFormat = '[=0]"today";[=1]"1 day";0" days"'
    sheet.Range("D32:D36").HorizontalAlignment = XL_LEFT

    # Flags: a word and a colour, never the colour alone.
    color_when_equal(sheet.Range("D13:D27"), "OVERDUE", WHITE, RED)
    color_when_equal(sheet.Range("D13:D27"), "This week", AMBER_TEXT, AMBER_FILL, bold=False)
    sheet.Range("D13:D27").HorizontalAlignment = XL_CENTER
    for address in ("F13:F27", "F32:F36"):
        color_when_equal(sheet.Range(address), "Done", GREEN, bold=False)
        color_when_equal(sheet.Range(address), "Blocked", AMBER_TEXT)
    for address in ("G13:G27", "G32:G36", "O13:O20"):
        color_when_equal(sheet.Range(address), "High", INK)  # bold, no colour: priority is not a status
    rule = sheet.Range("Q25:Q32").FormatConditions.Add(XL_CELL_VALUE, XL_GREATER, "=0")
    rule.Font.Color = rgb(RED)
    rule.Font.Bold = True

    # ---- Freeze the top band, set the print area, one page wide
    sheet.Range("A7").Select()
    excel.ActiveWindow.FreezePanes = True
    setup = sheet.PageSetup
    setup.PrintArea = DASHBOARD_AREA
    setup.Orientation = 2  # landscape
    setup.Zoom = False
    setup.FitToPagesWide = 1
    setup.FitToPagesTall = False
    sheet.Tab.Color = rgb(NAVY)
    sheet.Range("C5").Select()


def build_how_to(sheet, today: dt.date) -> None:
    excel = sheet.Application
    sheet.Activate()
    excel.ActiveWindow.DisplayGridlines = False
    sheet.Columns("A").ColumnWidth = 2
    sheet.Columns("B").ColumnWidth = 4
    sheet.Columns("C").ColumnWidth = 112
    sheet.Rows(1).RowHeight = 8
    lines = (
        ("How to use this tracker", "", "title"),
        ("", "Demo built with made-up sample data. Every task and name here is invented.", "note"),
        ("", "", ""),
        ("Add or change tasks", "", "head"),
        ("1", "Go to the Tasks sheet and type in the first empty row under the table. The table grows by itself.", ""),
        ("2", "Status and Priority are dropdowns. % Complete takes 0% to 100%. Year fills itself in from the Due date.", ""),
        ("3", "To remove a task, right-click its row in the table and choose Delete > Table Rows.", ""),
        ("", "", ""),
        ("Read the dashboard", "", "head"),
        ("4", "Pick an Owner, a Status or a Year in the shaded cells at the top. Everything on the page follows. Choose (All) to clear.", ""),
        ("5", "Overdue = not Done and the Due date is before today. Due this week = not Done and due today or in the next 6 days.", ""),
        ("6", f"Tasks in view shows up to {VIEW_ROWS} tasks, Next due dates {NEXT_ROWS}, This Week's Focus {FOCUS_ROWS}. "
              "A line under each list says how many more are not shown.", ""),
        ("7", "In the Tasks sheet, an overdue Due date turns red and says OVERDUE; one due this week turns amber and says so. "
              "They update by themselves, every day.", ""),
        ("", "", ""),
        ("Good to know", "", "head"),
        ("8", "Nothing here needs macros. It uses formulas that exist in Excel 2010 and later; "
              "built and tested in Excel for Microsoft 365 on Windows.", ""),
        ("9", "The Lists sheet holds helper formulas for the dashboard. Leave it as it is.", ""),
        ("10", f"The sample dates were set on {today:%b} {today.day}, {today.year}, so some tasks were overdue and some due that week. "
               "The dashboard always compares with today's date, so the counts move as days pass.", ""),
    )
    for offset, (left, right, style) in enumerate(lines):
        row = 2 + offset
        number, text = sheet.Range(f"B{row}"), sheet.Range(f"C{row}")
        if style in ("title", "head"):
            number.Value = left
            number.Font.Bold = True
            number.Font.Color = rgb(NAVY)
            number.Font.Size = 20 if style == "title" else 13
        else:
            if left:
                number.Value = int(left)
                number.Font.Color = rgb(SLATE)
            text.Value = right
            if style == "note":
                text.Font.Size = 9
                text.Font.Italic = True
                text.Font.Color = rgb(SLATE)
        sheet.Rows(row).RowHeight = 30 if style == "title" else 20
    sheet.Range("B2:B30").HorizontalAlignment = XL_LEFT
    sheet.Range("A1").Select()


# ---- Build ----------------------------------------------------------------

def build(today: dt.date) -> None:
    import win32com.client as win32

    WORKBOOK.unlink(missing_ok=True)
    excel = win32.DispatchEx("Excel.Application")  # a private instance, never the user's open Excel
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Add()
        normal = workbook.Styles("Normal").Font  # the guide's table text: Calibri 11, ink
        normal.Name = "Calibri"
        normal.Size = 11
        normal.Color = rgb(INK)
        while workbook.Worksheets.Count < 4:
            workbook.Worksheets.Add(None, workbook.Worksheets(workbook.Worksheets.Count))
        while workbook.Worksheets.Count > 4:
            workbook.Worksheets(workbook.Worksheets.Count).Delete()
        dashboard, tasks, how_to, lists = (workbook.Worksheets(i) for i in range(1, 5))
        dashboard.Name, tasks.Name, how_to.Name, lists.Name = "Dashboard", "Tasks", "How to use", "Lists"

        build_tasks(tasks, today)
        for name, refers_to in NAMES:
            try:
                workbook.Names.Add(name, refers_to)
            except Exception as error:
                raise RuntimeError(f"Excel rejected the name {name}: {refers_to}") from error
        build_lists(lists)
        build_dashboard(dashboard, workbook)
        build_how_to(how_to, today)

        properties = workbook.BuiltinDocumentProperties
        properties("Title").Value = "Work Tracker Demo"
        properties("Subject").Value = "Demo built with made-up sample data"
        properties("Author").Value = AUTHOR
        properties("Company").Value = AUTHOR

        dashboard.Activate()
        dashboard.Range("C5").Select()
        excel.CalculateFull()
        workbook.SaveAs(str(WORKBOOK), XL_OPEN_XML_WORKBOOK)
        workbook.Close(False)
    finally:
        excel.Quit()


# ---- Clean the saved file -------------------------------------------------

def clean_package(path: Path) -> list[str]:
    """Rewrite the file's properties and drop anything that describes this PC."""
    done = []
    with zipfile.ZipFile(path) as source:
        parts = [(info, source.read(info.filename)) for info in source.infolist()]

    def text_of(name: str) -> str | None:
        for info, data in parts:
            if info.filename == name:
                return data.decode("utf-8")
        return None

    edits: dict[str, str] = {}
    dropped = [info.filename for info, _ in parts
               if info.filename.startswith("xl/printerSettings/") or info.filename == "docProps/custom.xml"]

    core = text_of("docProps/core.xml")
    core = re.sub(r"<dc:creator>.*?</dc:creator>", f"<dc:creator>{AUTHOR}</dc:creator>", core, flags=re.S)
    if "<cp:lastModifiedBy>" in core:
        core = re.sub(r"<cp:lastModifiedBy>.*?</cp:lastModifiedBy>", f"<cp:lastModifiedBy>{AUTHOR}</cp:lastModifiedBy>", core, flags=re.S)
    else:
        core = core.replace("</dc:creator>", f"</dc:creator><cp:lastModifiedBy>{AUTHOR}</cp:lastModifiedBy>", 1)
    edits["docProps/core.xml"] = core
    done.append("author and last-modified-by set")

    app = text_of("docProps/app.xml")
    app = re.sub(r"<Manager>.*?</Manager>", "", app, flags=re.S)
    if "<Company>" in app:
        app = re.sub(r"<Company>.*?</Company>", f"<Company>{AUTHOR}</Company>", app, flags=re.S)
    edits["docProps/app.xml"] = app

    book = text_of("xl/workbook.xml")
    cleaned = re.sub(r"<mc:AlternateContent[^>]*>\s*<mc:Choice[^>]*>\s*<x15ac:absPath[^>]*/>\s*</mc:Choice>\s*</mc:AlternateContent>", "", book)
    if cleaned != book:
        done.append("stored folder path removed")
    edits["xl/workbook.xml"] = cleaned

    if dropped:
        done.append("removed " + ", ".join(dropped))
        for info, data in parts:
            name = info.filename
            if name.startswith("xl/worksheets/_rels/"):
                edits[name] = re.sub(r'<Relationship [^>]*printerSettings[^>]*/>', "", data.decode("utf-8"))
            elif re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name):
                edits[name] = re.sub(r'(<pageSetup[^>]*?) r:id="[^"]*"', r"\1", data.decode("utf-8"))
            elif name == "_rels/.rels":
                edits[name] = re.sub(r'<Relationship [^>]*custom-properties[^>]*/>', "", data.decode("utf-8"))
            elif name == "[Content_Types].xml":
                types = re.sub(r'<Override [^>]*docProps/custom\.xml[^>]*/>', "", data.decode("utf-8"))
                if not any(i.filename.endswith(".bin") and i.filename not in dropped for i, _ in parts):
                    types = re.sub(r'<Default Extension="bin"[^>]*/>', "", types)
                edits[name] = types

    temporary = path.with_suffix(".tmp")
    with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as target:
        for info, data in parts:
            if info.filename in dropped:
                continue
            if info.filename in edits:
                data = edits[info.filename].encode("utf-8")
            target.writestr(info, data, zipfile.ZIP_DEFLATED)
    temporary.replace(path)
    return done


def main() -> int:
    today = dt.date.today()
    print(f"Building {WORKBOOK.name} with dates relative to {today.isoformat()} ...")
    build(today)
    for line in clean_package(WORKBOOK):
        print("  cleaned:", line)
    DELIVERABLE.parent.mkdir(exist_ok=True)
    shutil.copyfile(WORKBOOK, DELIVERABLE)
    print(f"Saved {WORKBOOK.name} ({WORKBOOK.stat().st_size:,} bytes) and deliverable\\{DELIVERABLE.name}")
    print(f"{len(TASKS)} made-up tasks on the Tasks sheet")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
