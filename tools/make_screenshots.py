"""Rebuild the showcase images in screenshots/ (PNG, exactly 1600x1200).

Usage (Windows, with Excel installed):
    python tools\\make_screenshots.py

The images are composites on the AZG showcase template (tools\\showcase.py), not
screen grabs, so no window, account name or file path can appear in them. Excel
itself prints each view to PDF (in a private, hidden instance, with the workbook
opened read-only and never saved), PyMuPDF draws the PDF, and the kit frames it.

  0  the cover: the three leading totals and four lines of the main list, two
     flagged OVERDUE and two "This week". Cover-safe: no title strip, and
     everything sits in a band that survives a crop to 16:9 or 2:1
  1  the dashboard: the three leading totals and the main list
  2  the same view with the Owner filter set
  3  before / after: the plain task list, and the dashboard it feeds

The tool prints the size the text and the tile numbers come out at. The style
guide wants text at 20 px or more and headline numbers at 72 px or more.

A 400 px wide copy of each image is written to screenshots\\_work\\ so the
headline numbers can be checked at thumbnail size. For the cover, both crops
and their 400 px copies are written there too.
"""

from __future__ import annotations

from pathlib import Path

import pymupdf
from PIL import Image, ImageChops

import showcase as kit  # tools\showcase.py: a copy of the AZG showcase kit

ROOT = Path(__file__).resolve().parent.parent
WORKBOOK = ROOT / "Work Tracker Demo.xlsx"
SHOTS = ROOT / "screenshots"
WORK = SHOTS / "_work"
FILTER_OWNER = "Casey L."
COVER_HEADLINE = "See overdue work at a glance"
LIST_ROWS = range(13, 28)  # the Dashboard rows of the "Tasks in view" list
FLAG_COLUMN = "D"

# The views Excel prints: name -> (sheet, range, {filter cell name: value},
# columns to keep or None for all, cells to blank for this view only).
# Nothing here is saved: the workbook is open read-only and closed without saving.
# Order matters: a blanked cell stays blank for the views after it.
VIEWS = {
    "main": ("Dashboard", "A1:J24", {}, None, ()),  # top band, three tiles, the main list
    "owner": ("Dashboard", "A1:J24", {"SelOwner": FILTER_OWNER}, None, ()),
    # The cover: the tiles, the list's heading, and four list lines chosen in
    # export_views (the other list lines are hidden for this view only).
    "cover": ("Dashboard", "A7:J27", {}, None, ()),
    # Two tiles and the list's first columns. The list's side note runs past
    # column F and would be cut mid-sentence, so it is blanked for this view.
    "after": ("Dashboard", "A1:F22", {}, None, ("E11",)),
    # Task and Status, 30 tasks. More columns or rows would push the text under 20 px.
    "before": ("Tasks", "A1:K31", {}, ("B", "E"), ()),
}


def autocrop(image: Image.Image, pad: int = 12) -> Image.Image:
    image = image.convert("RGB")
    box = ImageChops.difference(image, Image.new("RGB", image.size, "white")).getbbox()
    if box is None:
        return image
    cropped = image.crop(box)
    framed = Image.new("RGB", (cropped.width + 2 * pad, cropped.height + 2 * pad), "white")
    framed.paste(cropped, (pad, pad))
    return framed


def pdf_image(pdf_path: Path, dpi: int = 300) -> Image.Image:
    with pymupdf.open(pdf_path) as document:
        pixmap = document[0].get_pixmap(dpi=dpi, alpha=False)
        return Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)


def cover_rows(sheet) -> list[int]:
    """Four neighbouring list rows for the cover: two flagged OVERDUE, then two "This week"."""
    first, last = LIST_ROWS[0], LIST_ROWS[-1]
    flags = [row[0] for row in sheet.Range(f"{FLAG_COLUMN}{first}:{FLAG_COLUMN}{last}").Value]
    for index in range(2, len(flags) - 1):
        if flags[index - 2:index + 2] == ["OVERDUE", "OVERDUE", "This week", "This week"]:
            return [first + index + offset for offset in (-2, -1, 0, 1)]
    raise SystemExit("The list doesn't show two OVERDUE lines followed by two 'This week' lines today. "
                     "Rebuild the workbook for fresh dates, then run this again.")


def export_views() -> tuple[dict[str, Image.Image], dict]:
    """Have Excel print each view to its own PDF. Nothing is saved to the workbook.

    Returns the images and a few facts read from the workbook for the cover.
    """
    import win32com.client as win32

    WORK.mkdir(parents=True, exist_ok=True)
    pdfs, facts = {}, {}
    excel = win32.DispatchEx("Excel.Application")  # a private instance, never the user's open Excel
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(str(WORKBOOK), 0, True)  # no link updates, read-only
        for name, (sheet_name, address, filters, keep_columns, blank_cells) in VIEWS.items():
            for filter_name in ("SelOwner", "SelStatus", "SelYear"):
                workbook.Names(filter_name).RefersToRange.Cells(1, 1).Value = filters.get(filter_name, "(All)")
            sheet = workbook.Worksheets(sheet_name)
            for cell in blank_cells:
                sheet.Range(cell).ClearContents()
            excel.CalculateFull()
            if keep_columns:
                for index in range(1, sheet.Range(address).Columns.Count + 1):
                    letter = chr(ord("A") + index - 1)
                    sheet.Columns(letter).Hidden = letter not in keep_columns
            if name == "cover":
                shown = cover_rows(sheet)
                for row in LIST_ROWS:
                    sheet.Rows(row).Hidden = row not in shown
                facts["totals"] = {label: int(workbook.Names(cell).RefersToRange.Cells(1, 1).Value)
                                   for label, cell in (("total", "KpiTotal"), ("overdue", "KpiOverdue"),
                                                       ("due this week", "KpiDueThisWeek"))}
                facts["lines"] = [" | ".join(str(sheet.Range(f"{column}{row}").Text) for column in "BCD") for row in shown]
            setup = sheet.PageSetup
            setup.Orientation = 2 if sheet_name == "Dashboard" else 1
            setup.Zoom = False
            setup.FitToPagesWide = 1
            setup.FitToPagesTall = 1
            setup.PrintHeadings = False
            setup.PrintGridlines = False
            pdfs[name] = WORK / f"{name}.pdf"
            sheet.Range(address).ExportAsFixedFormat(0, str(pdfs[name]), 0, False, True)
            if name == "cover":
                for row in LIST_ROWS:  # put the list back for the views that follow
                    sheet.Rows(row).Hidden = False
        workbook.Close(False)
    finally:
        excel.Quit()
    return {name: autocrop(pdf_image(path)) for name, path in pdfs.items()}, facts


def finish(card: Image.Image, caption: str, file_name: str, note: str = "") -> None:
    path = kit.compose(card, caption, SHOTS / file_name)
    small = kit.thumbnail(path, WORK)
    print(f"  {path.name}  {kit.WIDTH}x{kit.HEIGHT}  {note}  (thumbnail: _work\\{small.name})")


def filled_card(view: Image.Image) -> tuple[Image.Image, float]:
    card = kit.new_card()
    _, scale = kit.place(card, view, kit.inner(card))
    return card, scale


def main() -> int:
    SHOTS.mkdir(exist_ok=True)
    print(f"showcase kit {kit.__version__}. Excel is printing the views to PDF ...")
    views, facts = export_views()
    print("Writing screenshots:")

    card = kit.new_cover_card()
    _, scale = kit.place(card, views["cover"], kit.inner(card))
    made = kit.cover(card, COVER_HEADLINE, SHOTS / "0-cover.png")
    previews = kit.cover_previews(made["path"], WORK)
    totals = ", ".join(f"{value} {label}" for label, value in facts["totals"].items())
    print(f"  {made['path'].name}  {kit.WIDTH}x{kit.HEIGHT}  cover-safe, {made['style']} style, headline at {made['headline_size']} px, "
          f"text {45.8 * scale:.0f} px, tile numbers {150 * scale:.0f} px")
    print(f"    numbers shown: {totals}")
    for line in facts["lines"]:
        print(f"    list line: {line}")
    print(f"    everything drawn sits in {made['drawn']}, inside the safe band {kit.COVER_BAND}")
    print("    previews in _work\\: " + ", ".join(path.name for path in previews.values()))

    card, scale = filled_card(views["main"])
    finish(card, "See what's overdue at a glance", "1-dashboard-overview.png",
           note=f"text {45.8 * scale:.0f} px, tile numbers {150 * scale:.0f} px")

    card, scale = filled_card(views["owner"])
    finish(card, "Pick an owner, the page follows", "2-filter-by-owner.png",
           note=f"text {45.8 * scale:.0f} px, tile numbers {150 * scale:.0f} px")

    card, scales = kit.before_after(views["before"], views["after"], share=0.34,
                                    before_note="a plain task list", after_note="the dashboard it feeds")
    # The exports are drawn at 300 dpi, so 11 pt text is 45.8 px before scaling.
    finish(card, "From a task list to a live dashboard", "3-before-after-list-to-dashboard.png",
           note=f"text {45.8 * scales[0]:.0f} px before, {45.8 * scales[1]:.0f} px after, "
                f"tile numbers {150 * scales[1]:.0f} px")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
