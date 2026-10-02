"""Check the workbook in Excel itself, against the same rules written in Python.

Usage (Windows, with Excel installed):
    python tests\\test_workbook.py                      checks "Work Tracker Demo.xlsx"
    python tests\\test_workbook.py deliverable          checks the copy in deliverable\\

What it does, in a private, hidden Excel with the workbook opened read-only:
  1. opens the workbook and recalculates;
  2. reads the Tasks table and works out, in Python, what every number and every
     list line on the Dashboard should be; compares them all;
  3. repeats that with the Owner, Status and Year filters set;
  4. appends 3 rows to the table and checks the totals moved by exactly the
     expected amounts, and that the lists and dropdowns picked the rows up;
  5. types a row under the table, adds a task with no due date, and adds 8 more
     rows for one owner so that owner has more tasks than the lists show, and
     checks the "+ N more not shown" lines;
  6. closes without saving and checks the file on disk did not change.

Exit code 0 when every check passes.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import build_workbook as bw  # noqa: E402  (list lengths and labels only)

ALL = bw.ALL
TODAY = bw.serial(dt.date.today())
FIELDS = ("id", "task", "owner", "department", "status", "priority", "start", "due", "percent", "year", "notes")
FAILURES: list[str] = []


def check(label: str, actual, expected, show: str | None = None) -> None:
    if actual == expected:
        print(f"PASS  {label}" + (f": {show}" if show else ""))
    else:
        FAILURES.append(label)
        print(f"FAIL  {label}\n        Excel:  {actual}\n        Python: {expected}")


def norm(value):
    """Make a cell value comparable: blanks become '', numbers are rounded."""
    if value is None:
        return ""
    if isinstance(value, (int, float)):
        return round(float(value), 6)
    return value


def grid(cells) -> list[list]:
    values = cells.Value2
    if not isinstance(values, tuple):
        values = ((values,),)
    return [[norm(value) for value in row] for row in values]


def column(cells) -> list:
    return [row[0] for row in grid(cells)]


def is_number(value) -> bool:
    return isinstance(value, (int, float))


def more(count: int, shown: int, what: str = "more not shown") -> str:
    return f"+ {count - shown} {what}" if count > shown else ""


# ---- The rules, in Python ---------------------------------------------------

def read_tasks(workbook) -> list[dict]:
    table = workbook.Worksheets("Tasks").ListObjects("Tasks")
    return [dict(zip(FIELDS, row)) for row in grid(table.DataBodyRange)]


def expected_dashboard(tasks: list[dict], owner, status, year) -> dict:
    scope = [(i, t) for i, t in enumerate(tasks, start=1)
             if t["task"] != ""
             and (owner == ALL or t["owner"] == owner)
             and (status == ALL or t["status"] == status)
             and (year == ALL or t["year"] == year)]
    open_dated = [(i, t) for i, t in scope if t["status"] != "Done" and is_number(t["due"])]
    overdue = [(i, t) for i, t in open_dated if t["due"] < TODAY]
    upcoming = sorted(((i, t) for i, t in open_dated if t["due"] >= TODAY), key=lambda it: (it[1]["due"], it[0]))
    focus = [(i, t) for i, t in upcoming if t["due"] <= TODAY + 6]
    view = sorted(scope, key=lambda it: (it[1]["status"] == "Done",
                                         it[1]["due"] if is_number(it[1]["due"]) else 99999, it[0]))

    def flag(t: dict) -> str:
        if t["status"] == "Done" or not is_number(t["due"]):
            return ""
        return "OVERDUE" if t["due"] < TODAY else ("This week" if t["due"] <= TODAY + 6 else "")

    def padded(rows: list[list], length: int, width: int, empty_text: str) -> list[list]:
        rows = rows[:length] + [[""] * width for _ in range(length - len(rows[:length]))]
        if rows[0][1] == "":
            rows[0][1] = empty_text
        return rows

    owners = sorted({t["owner"] for t in tasks if t["owner"] != ""}, key=str.lower)
    owner_rows = []
    for name in owners[:bw.OWNER_ROWS]:
        mine = [t for _, t in scope if t["owner"] == name]
        average = round(sum(t["percent"] or 0 for t in mine) / len(mine), 6) if mine else ""
        owner_rows.append([name, average, float(len(mine)), float(sum(t["status"] == "Done" for t in mine)),
                           float(sum(t["owner"] == name for _, t in overdue))])
    owner_rows += [[""] * 5 for _ in range(bw.OWNER_ROWS - len(owner_rows))]

    return {
        "totals": {
            "Total": float(len(scope)),
            "Done": float(sum(t["status"] == "Done" for _, t in scope)),
            "In progress": float(sum(t["status"] == "In progress" for _, t in scope)),
            "Overdue": float(len(overdue)),
            "Due this week": float(len(focus)),
        },
        "focus": padded([[t["due"], t["task"], t["owner"], t["priority"], t["percent"]] for _, t in focus],
                        bw.FOCUS_ROWS, 5, "Nothing due in the next 7 days"),
        "focus_more": more(len(focus), bw.FOCUS_ROWS),
        "next": padded([[t["due"], t["task"], t["owner"], t["status"], t["priority"], t["percent"],
                         t["due"] - TODAY, t["department"]] for _, t in upcoming], bw.NEXT_ROWS, 8, "Nothing coming up"),
        "next_more": more(len(upcoming), bw.NEXT_ROWS),
        "view": padded([[t["due"] if is_number(t["due"]) else "", t["task"], t["owner"], t["status"], t["priority"],
                         t["percent"], flag(t), t["department"]] for _, t in view],
                       bw.VIEW_ROWS, 8, "No tasks match these filters"),
        "view_more": more(len(scope), bw.VIEW_ROWS),
        "owners": owner_rows,
        "owners_more": more(len(owners), bw.OWNER_ROWS, "more owners not shown"),
        "owner_choices": [ALL, *owners],
        "year_choices": [ALL, *sorted({t["year"] for t in tasks if is_number(t["year"])})],
    }


# ---- What Excel shows -------------------------------------------------------

def read_dashboard(workbook) -> dict:
    sheet = workbook.Worksheets("Dashboard")

    def named(name: str):
        return norm(workbook.Names(name).RefersToRange.Cells(1, 1).Value2)

    def block(name: str) -> list[list]:
        """A list's lines, read from where the build put it, as the tested fields in order."""
        spec = bw.LAYOUT[name]
        return [[row[i] for i in spec["columns"]] for row in grid(sheet.Range(spec["cells"]))]

    def more_line(name: str):
        return norm(sheet.Range(bw.LAYOUT[name]["more"]).Value2)

    return {
        "totals": {"Total": named("KpiTotal"), "Done": named("KpiDone"), "In progress": named("KpiInProgress"),
                   "Overdue": named("KpiOverdue"), "Due this week": named("KpiDueThisWeek")},
        "focus": block("focus"),
        "focus_more": more_line("focus"),
        "next": block("next"),
        "next_more": more_line("next"),
        "view": block("view"),
        "view_more": more_line("view"),
        "owners": block("owners"),
        "owners_more": more_line("owners"),
        "owner_choices": column(workbook.Names("OwnerList").RefersToRange),
        "year_choices": column(workbook.Names("YearList").RefersToRange),
    }


def error_cells(workbook) -> list[str]:
    """Addresses of any cell showing an Excel error (#N/A, #VALUE!, ...)."""
    found = []
    for sheet_name in ("Dashboard", "Lists", "Tasks"):
        used = workbook.Worksheets(sheet_name).UsedRange
        for r, row in enumerate(used.Value2, start=used.Row):
            for c, value in enumerate(row, start=used.Column):
                if isinstance(value, int) and -2146826300 < value < -2146826200:
                    found.append(f"{sheet_name}!R{r}C{c}")
    return found


def set_filters(excel, workbook, owner=ALL, status=ALL, year=ALL) -> None:
    for name, value in (("SelOwner", owner), ("SelStatus", status), ("SelYear", year)):
        workbook.Names(name).RefersToRange.Cells(1, 1).Value = value
    excel.CalculateFull()


def verify(excel, workbook, title: str, owner=ALL, status=ALL, year=ALL) -> dict:
    """Set the filters, then compare everything on the Dashboard with Python."""
    set_filters(excel, workbook, owner, status, year)
    want = expected_dashboard(read_tasks(workbook), owner, status, year)
    got = read_dashboard(workbook)
    print(f"\n-- {title}  [Owner={owner}, Status={status}, Year={year}]")
    numbers = ", ".join(f"{k}={int(v)}" for k, v in got["totals"].items())
    check("5 totals match the Python counts", got["totals"], want["totals"], numbers)
    shown = sum(1 for row in want["focus"] if row[0] != "")
    check("This Week's Focus lines and its 'more' line", (got["focus"], got["focus_more"]),
          (want["focus"], want["focus_more"]), f"{shown} lines, more line = {want['focus_more']!r}")
    shown = sum(1 for row in want["next"] if row[0] != "")
    check("Next due dates lines and its 'more' line", (got["next"], got["next_more"]),
          (want["next"], want["next_more"]), f"{shown} lines, more line = {want['next_more']!r}")
    shown = sum(1 for row in want["view"] if row[2] != "")
    check("Tasks in view lines and its 'more' line", (got["view"], got["view_more"]),
          (want["view"], want["view_more"]), f"{shown} lines, more line = {want['view_more']!r}")
    check("Progress by owner", (got["owners"], got["owners_more"]), (want["owners"], want["owners_more"]),
          f"{sum(1 for row in want['owners'] if row[0])} owners")
    check("Owner and Year dropdown choices", (got["owner_choices"], got["year_choices"]),
          (want["owner_choices"], want["year_choices"]),
          f"{len(want['owner_choices']) - 1} owners, years {[int(y) for y in want['year_choices'][1:]]}")
    check("no cell shows an Excel error", error_cells(workbook), [])
    return got


def append_row(table, values: dict) -> None:
    """Add a row to the end of the table, leaving the Year formula to fill itself."""
    cells = table.ListRows.Add().Range
    for index, field in enumerate(FIELDS, start=1):
        if field != "year":
            cells.Cells(1, index).Value = values.get(field, "")


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main(argv: list[str]) -> int:
    import win32com.client as win32

    path = bw.DELIVERABLE if argv[:1] == ["deliverable"] else bw.WORKBOOK
    before = digest(path)
    print(f"Workbook: {path.relative_to(ROOT)}   today = {dt.date.today().isoformat()}")

    excel = win32.DispatchEx("Excel.Application")  # a private instance, never the user's open Excel
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        print(f"Excel {excel.Version} build {excel.Build}, calculation engine {excel.CalculationVersion}")
        workbook = excel.Workbooks.Open(str(path), 0, True)  # no link updates, read-only
        excel.CalculateFull()
        table = workbook.Worksheets("Tasks").ListObjects("Tasks")
        check("no macros in the workbook", workbook.HasVBProject, False)
        start_rows = table.ListRows.Count
        print(f"Tasks table: {start_rows} rows, {table.ListColumns.Count} columns")
        check("table has the 11 fields", [c.Name for c in table.ListColumns], list(bw.HEADERS))

        # 2. Everything on the Dashboard, no filters.
        base = verify(excel, workbook, "All tasks")["totals"]

        # 3. With filters.
        tasks = read_tasks(workbook)
        owners = sorted({t["owner"] for t in tasks})
        busiest = max(owners, key=lambda name: sum(t["owner"] == name for t in tasks))
        got = verify(excel, workbook, "One owner", owner=busiest)
        listed = [row[2] for row in got["view"] if row[2] != ""]
        check("filtered list shows only that owner's tasks", set(listed), {busiest}, f"{len(listed)} lines, all {busiest}")
        verify(excel, workbook, "One status", status="In progress")
        verify(excel, workbook, "One year", year=max(t["year"] for t in tasks if is_number(t["year"])))
        verify(excel, workbook, "Owner and status together", owner=busiest, status="Blocked")
        verify(excel, workbook, "A filter that matches nothing", owner=owners[0], status="Blocked", year=1999.0)

        # 4. Append 3 rows: one done, one overdue, one due this week.
        print("\n-- Appending 3 rows to the table")
        append_row(table, dict(id="T-901", task="Test: finished task", owner="Quinn B.", department="IT", status="Done",
                               priority="Low", start=TODAY - 5, due=TODAY + 2, percent=1))
        append_row(table, dict(id="T-902", task="Test: overdue task", owner=busiest, department="Sales",
                               status="In progress", priority="High", start=TODAY - 10, due=TODAY - 3, percent=0.5))
        append_row(table, dict(id="T-903", task="Test: due this week", owner=busiest, department="Sales",
                               status="Not started", priority="Medium", start=TODAY, due=TODAY + 1, percent=0))
        after = verify(excel, workbook, "All tasks, after adding 3 rows")
        moved = {key: int(after["totals"][key] - base[key]) for key in base}
        check("totals moved by exactly the expected amounts", moved,
              {"Total": 3, "Done": 1, "In progress": 1, "Overdue": 1, "Due this week": 1}, str(moved))
        check("table grew by 3 rows", table.ListRows.Count, start_rows + 3)
        check("the column names grew with the table", workbook.Names("T_Owner").RefersToRange.Rows.Count, start_rows + 3)
        years = column(table.ListColumns("Year").DataBodyRange)[-3:]
        due_years = [float((dt.date.today() + dt.timedelta(days=offset)).year) for offset in (2, -3, 1)]
        check("Year filled itself in on the new rows", years, due_years, str(years))
        check("new owner appears in the Owner dropdown", "Quinn B." in after["owner_choices"], True)
        check("new task appears in This Week's Focus", "Test: due this week" in [row[1] for row in after["focus"]], True)
        check("new overdue task is flagged in Tasks in view",
              ["Test: overdue task", "OVERDUE"] in [[row[1], row[6]] for row in after["view"]], True)

        # 4b. The other way a table grows: type in the row under it. Excel then
        # stretches the table over the new row, which is what Resize does here.
        print("\n-- Typing a row under the table (the table stretches to take it in)")
        sheet = workbook.Worksheets("Tasks")
        below = table.Range.Row + table.Range.Rows.Count
        typed = ("T-904", "Test: typed under the table", "Quinn B.", "IT", "Blocked", "Low", TODAY, TODAY + 40, 0.25)
        sheet.Range(sheet.Cells(below, 1), sheet.Cells(below, 9)).Value = (typed,)
        table.Resize(sheet.Range(table.Range.Cells(1, 1), sheet.Cells(below, 11)))
        typed_in = verify(excel, workbook, "All tasks, after typing a row under the table")
        check("total went up by one more", int(typed_in["totals"]["Total"] - after["totals"]["Total"]), 1)
        check("Year filled itself in on the typed row", column(table.ListColumns("Year").DataBodyRange)[-1],
              float((dt.date.today() + dt.timedelta(days=40)).year))

        # 4c. A task with no due date: counted, listed after the dated open tasks, never flagged.
        print("\n-- Adding a task with no due date")
        append_row(table, dict(id="T-905", task="Test: no due date", owner="Quinn B.", department="IT",
                               status="Not started", priority="Low", start=TODAY, due="", percent=0))
        got = verify(excel, workbook, "The new owner's three tasks", owner="Quinn B.")
        check("open dated task, then the undated one, then the done one", [row[:2] + [row[6]] for row in got["view"][:3]],
              [[float(TODAY + 40), "Test: typed under the table", ""], ["", "Test: no due date", ""],
               [float(TODAY + 2), "Test: finished task", ""]])

        # 5. More tasks for one owner than the lists can show, all due this week.
        print(f"\n-- Adding 8 more rows for {busiest}, due this week")
        for n in range(8):
            append_row(table, dict(id=f"T-91{n}", task=f"Test: extra task {n + 1}", owner=busiest, department="Sales",
                                   status="Not started", priority="Low", start=TODAY, due=TODAY + n % 7, percent=0))
        got = verify(excel, workbook, "One owner with more tasks than the lists show", owner=busiest)
        count = int(got["totals"]["Total"])
        listed = [row[2] for row in got["view"]]
        check(f"all {bw.VIEW_ROWS} list lines belong to that owner", listed, [busiest] * bw.VIEW_ROWS)
        check("'more not shown' count is right", got["view_more"], f"+ {count - bw.VIEW_ROWS} more not shown",
              f"{count} tasks, {bw.VIEW_ROWS} shown, line reads {got['view_more']!r}")
        due_soon = int(got["totals"]["Due this week"])
        check("This Week's Focus 'more' count is right", got["focus_more"], f"+ {due_soon - bw.FOCUS_ROWS} more not shown",
              f"{due_soon} due this week, {bw.FOCUS_ROWS} shown, line reads {got['focus_more']!r}")

        # 6. Close without saving.
        workbook.Close(False)
    finally:
        excel.Quit()

    print()
    check("file on disk unchanged after closing without saving", digest(path), before, before[:16] + "...")
    print(f"\n{'ALL CHECKS PASSED' if not FAILURES else str(len(FAILURES)) + ' CHECK(S) FAILED: ' + '; '.join(FAILURES)}")
    return 1 if FAILURES else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
