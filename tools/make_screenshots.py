"""Rebuild the showcase images in screenshots/ (PNG, exactly 1600x1200).

Usage (Windows, with Excel installed):
    python tools\\make_screenshots.py

The images are composites on the AZG showcase template (tools\\showcase.py), not
screen grabs, so no window, account name or file path can appear in them. Excel
itself prints each view to PDF (in a private, hidden instance, with the workbook
opened read-only and never saved), PyMuPDF draws the PDF, and the kit frames it.

  1  the dashboard: the three leading totals and the main list
  2  the same view with the Owner filter set
  3  before / after: the plain task list, and the dashboard it feeds

The tool prints the size the text and the tile numbers come out at. The style
guide wants text at 20 px or more and headline numbers at 72 px or more.

A 400 px wide copy of each image is written to screenshots\\_work\\ so the
headline numbers can be checked at thumbnail size.
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

# The views Excel prints: name -> (sheet, range, {filter cell name: value},
# columns to keep or None for all, cells to blank for this view only).
# Nothing here is saved: the workbook is open read-only and closed without saving.
VIEWS = {
    "main": ("Dashboard", "A1:J24", {}, None, ()),  # top band, three tiles, the main list
    "owner": ("Dashboard", "A1:J24", {"SelOwner": FILTER_OWNER}, None, ()),
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


def export_views() -> dict[str, Image.Image]:
    """Have Excel print each view to its own PDF. Nothing is saved to the workbook."""
    import win32com.client as win32

    WORK.mkdir(parents=True, exist_ok=True)
    pdfs = {}
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
            setup = sheet.PageSetup
            setup.Orientation = 2 if sheet_name == "Dashboard" else 1
            setup.Zoom = False
            setup.FitToPagesWide = 1
            setup.FitToPagesTall = 1
            setup.PrintHeadings = False
            setup.PrintGridlines = False
            pdfs[name] = WORK / f"{name}.pdf"
            sheet.Range(address).ExportAsFixedFormat(0, str(pdfs[name]), 0, False, True)
        workbook.Close(False)
    finally:
        excel.Quit()
    return {name: autocrop(pdf_image(path)) for name, path in pdfs.items()}


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
    views = export_views()
    print("Writing screenshots:")

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
