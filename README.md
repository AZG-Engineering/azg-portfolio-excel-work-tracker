# Excel work tracker with a live dashboard

One workbook, no macros. People type tasks into a table; the Dashboard sheet keeps
itself up to date:

- **5 totals:** Total, Overdue, Due this week, In progress, Done.
- **Tasks in view:** the task list, open tasks first, with OVERDUE and "This week"
  flags.
- **Next 5 due dates**, soonest first.
- **This Week's Focus:** what is not done and due in the next 7 days.
- **Progress by owner**, with progress bars.
- **Three filters** (Owner, Status, Year) as dropdown cells. Pick one and the whole
  page follows.

In the Tasks sheet, an overdue due date turns red and says OVERDUE, and one due this
week turns amber and says so, by themselves, every day.

**Demo built with made-up sample data.** Every task, name and department is invented.

Built by AZG Engineering.

## What's in the folder

| Path | What it is |
| --- | --- |
| `Work Tracker Demo.xlsx` | The workbook. |
| `deliverable\Work Tracker Demo.xlsx` | An identical clean copy, to hand out. |
| `build_workbook.py` | Builds the workbook from scratch with Excel, then cleans the file's properties. |
| `tests\test_workbook.py` | Opens the workbook in Excel and checks every number and list line against the same rules in Python. |
| `tools\make_screenshots.py` | Rebuilds the three images in `screenshots\`. |
| `tools\check_metadata.py` | Shows the author fields and checks nothing about this PC is stored in the files. |
| `tools\showcase.py`, `tools\fonts\` | The AZG showcase kit (the screenshot template) and the Inter font it uses. |
| `screenshots\` | Three 1600x1200 PNG images. |

## How to use the workbook

Open `Work Tracker Demo.xlsx`. The **How to use** sheet says it in ten lines. In short:

- **Tasks sheet:** type a task in the first empty row under the table; the table
  grows by itself. Status and Priority are dropdowns. Year fills itself in from Due.
- **Dashboard sheet:** pick an Owner, a Status or a Year in the shaded cells at the
  top. Choose `(All)` to clear.
- **Lists sheet:** helper formulas. Leave it alone.

**The dashboard depends on today's date.** Overdue means not Done and due before
today; Due this week means not Done and due today or in the next 6 days. The sample
dates were set relative to the day the workbook was built, so that some tasks were
overdue and some due that week. As days pass, more of them become overdue. Rebuild
for fresh dates (below).

## Excel versions

**Works in Excel 2010 and later, including Microsoft 365.** Lists show a fixed number
of rows; on 365 the same views could be built with FILTER/SORT.

- Functions used: `SUMPRODUCT`, `AGGREGATE`, `INDEX`, `MATCH`, `COUNTIF`, `COUNT`,
  `IFERROR`, `IF`, `AND`, `OR`, `ISNUMBER`, `INT`, `MOD`, `ROUND`, `ROW`, `MIN`,
  `TODAY`, `TEXT`, `YEAR`.
- `AGGREGATE` is the newest of these. It arrived in Excel 2010, so the workbook does
  not work in Excel 2007.
- No dynamic-array functions (no FILTER, SORT, UNIQUE, XLOOKUP or LET), no macros,
  no add-ins, no array formulas that need Ctrl+Shift+Enter.
- The progress bars use the fixed 0% to 100% scale that Excel 2010 added.
- **Tested on:** Excel 2016 (Windows). It has not been opened here in 2010, 2013,
  2019 or Microsoft 365; nothing in it is newer than 2010.

Because the lists are a fixed length, each one has a line underneath saying how many
more match but are not shown: 8 lines for This Week's Focus, 5 for Next due dates,
15 for Tasks in view, 8 owners.

## How it works

The Dashboard's formulas are short because the logic sits in named formulas
(**Formulas > Name Manager**):

| Name | What it is |
| --- | --- |
| `SelOwner`, `SelStatus`, `SelYear` | The three filter cells. |
| `T_ID`, `T_Task`, `T_Owner`, `T_Status`, `T_Due`, `T_Year` | Columns of the Tasks table. They grow with the table. |
| `InScope` | 1 for each task that passes all three filters, else 0. |
| `CondOverdue` | In scope, not Done, due before today. |
| `CondUpcoming` | In scope, not Done, due today or later. |
| `CondFocus` | `CondUpcoming` and due within the next 6 days. |
| `SortKey`, `SortKeyView` | The due date plus a tiny bit that identifies the row, so the lists can be sorted by date with ties in table order. `SortKeyView` puts Done tasks last. |
| `OwnerList`, `StatusList`, `YearList` | What the three dropdowns offer. |
| `KpiTotal`, `KpiDone`, `KpiInProgress`, `KpiOverdue`, `KpiDueThisWeek` | The five total cells. |

A total is one line, for example Overdue: `=SUMPRODUCT(CondOverdue)`.

A list works in two steps. On the **Lists** sheet, `AGGREGATE(15,6,SortKey/CondFocus,n)`
finds the n-th smallest sort key among the matching rows, and the next cell turns
that key into a table row number. On the Dashboard, each cell then just reads that
row: `=IF(Lists!$I4="","",INDEX(Tasks[Task],Lists!$I4)&"")`.

## How it looks

The look follows the AZG build style guide (`Products\AZG build style guide.md`):

- the guide's colours and no others, Calibri throughout, no gridlines;
- a top band with the title, the "As of" line and the three filter cells, frozen in
  place when the page scrolls;
- five white tiles of equal width, each with a coloured bar on the left: navy, or
  red for Overdue and amber-brown for Due this week;
- the main list on the left, the supporting blocks on the right;
- list headers in navy on a pale band, no vertical lines, no boxes;
- teal progress bars with the % in its own column, so a bar never covers a number;
- every coloured flag also says a word: "OVERDUE", "This week";
- the Dashboard prints landscape, one page wide.

Two places where the guide's wording was not followed, on purpose:

- **The note on the Year column is an input message, not a cell comment.** It
  appears when a Year cell is selected. A comment would store the Office user's
  name inside the file.
- **The Tasks table has no progress bars.** With exactly 11 fields there is no
  spare column for a bar beside the number, and a bar under the number would cover
  it. The bars are on the Dashboard.

## How to change it

Small changes can be made straight in the workbook. For anything structural, change
`build_workbook.py` and rebuild, so the test and screenshots stay in step.

| To change | Where |
| --- | --- |
| The sample tasks | `TASKS` in `build_workbook.py`. |
| How many lines a list shows | `FOCUS_ROWS`, `NEXT_ROWS`, `VIEW_ROWS`, `OWNER_ROWS` in `build_workbook.py`. Larger numbers also need the Dashboard rows below moved down in `build_dashboard`, and the addresses in `LAYOUT` (which tell the test where each list is) updated to match. |
| Where a block sits, or a column's width | `WIDTHS`, `TILES` and `build_dashboard` in `build_workbook.py`, then `LAYOUT`. The formulas' logic doesn't depend on position. |
| What counts as "this week" | The `+6` in the name `CondFocus`, and in the amber rule on the Tasks sheet (`build_tasks`). |
| The statuses | `STATUSES` in `build_workbook.py`. "Done" is the one that closes a task; it appears in the names `CondOverdue`, `CondUpcoming` and `SortKeyView`. |
| The priorities | `PRIORITIES` in `build_workbook.py`. |
| A new table column | Add it to `HEADERS` and `TASKS`, and to `FIELDS` in the test. Give it a name in `COLUMN_NAMES` only if a named formula needs it. |
| Colours | The palette constants near the top of `build_workbook.py` (`NAVY`, `TEAL`, `RED` and so on). They mirror section 2 of the style guide; change the guide first. |

Limits to keep in mind: up to 30 owners and 10 years in the dropdowns, and Due must
hold real dates. A task with no due date is counted and listed last, but can't be
overdue.

## Rebuild, test, screenshots

Windows with Excel installed. Excel runs hidden in the background; your own open
Excel windows are not touched. Everything installs into this folder only.

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

```powershell
.\.venv\Scripts\python.exe build_workbook.py
```

```powershell
.\.venv\Scripts\python.exe tests\test_workbook.py
```

```powershell
.\.venv\Scripts\python.exe tools\make_screenshots.py
```

```powershell
.\.venv\Scripts\python.exe tools\check_metadata.py
```

- **Build:** writes both workbook copies with dates relative to today. After Excel
  saves, the script rewrites the file's properties (author and "last modified by"
  become "AZG Engineering") and removes the local folder path and any printer
  settings Excel stored. Don't re-save the deliverable from Excel afterwards: Excel
  would stamp your Office user name on it again. Rebuild instead.
- **Test:** opens the workbook read-only, compares the Dashboard with Python's own
  counts under several filters, adds test rows in memory, and closes without saving.
  Add `deliverable` to the command to test the other copy.
- **Screenshots:** composites, not screen grabs. Excel prints each view to PDF,
  which is then drawn as an image and framed by the AZG showcase kit
  (`tools\showcase.py`). Don't edit that copy: change the kit in
  `Products\showcase-kit\` and sync it. A 400 px wide copy of each image goes to
  `screenshots\_work\` to check the headline numbers read at thumbnail size.

## Dependencies and licences

The workbook itself depends on nothing but Excel. Nothing is copied into this
repository from a third party: no templates, icons or code.

The tooling uses these pip packages, installed into `.venv`:

| Package | Version | Licence | Used for |
| --- | --- | --- | --- |
| pywin32 | 312 | PSF | driving Excel to build, test and print the workbook |
| Pillow | 12.3.0 | MIT-CMU | composing the screenshots, checking them |
| PyMuPDF | 1.28.2 | **AGPL-3.0, or a paid Artifex licence** | screenshots only: PDF to image |

PyMuPDF is the one to know about. Only `tools\make_screenshots.py` uses it. The
workbook, the build script and the test do not. If this is ever sold or handed over
with its tooling, leave PyMuPDF out or swap that step for another PDF renderer.

Kept in this repository, for the screenshots only:

| Item | Version | Licence | Where |
| --- | --- | --- | --- |
| Inter font (Regular, SemiBold) | 4.1 | SIL Open Font License 1.1 | `tools\fonts\`, with the licence text in `tools\fonts\OFL.txt` |
| AZG showcase kit | 1.0.0 | AZG Engineering's own | `tools\showcase.py`, an exact copy of `Products\showcase-kit\showcase.py` |

The workbook uses Calibri, which ships with Office. It is referenced by name, not
embedded. No system font is drawn into the framed images.
