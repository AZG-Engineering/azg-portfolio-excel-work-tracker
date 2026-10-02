"""Rebuild the portfolio screenshots in screenshots/ (PNG, exactly 1600x1200).

Usage (Windows, with Excel installed):
    python tools\\make_screenshots.py

The images are composites, not screen grabs, so no window, account name or file
path can appear in them. Excel itself prints each view to PDF (in a private,
hidden instance, with the workbook opened read-only and never saved), PyMuPDF
draws the PDF, and Pillow adds the title strip.
"""

from __future__ import annotations

from pathlib import Path

import pymupdf
from PIL import Image, ImageChops, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
WORKBOOK = ROOT / "Work Tracker Demo.xlsx"
SHOTS = ROOT / "screenshots"
WORK = SHOTS / "_work"
FONTS = Path(r"C:\Windows\Fonts")

WIDTH, HEIGHT = 1600, 1200
STRIP_HEIGHT = 104
MARGIN = 40
CONTENT_BOX = (MARGIN, STRIP_HEIGHT + 28, WIDTH - MARGIN, HEIGHT - 56)
BACKGROUND = "#F4F6F8"
STRIP = "#1F3A5F"
FOOTNOTE = "Demo \u00b7 sample data"
FILTER_OWNER = "Casey L."

# file name, caption (4 to 8 words), sheet, range, {filter cell name: value}
VIEWS = (
    ("1-dashboard-overview.png", "Dashboard updates as tasks are added", "Dashboard", "A1:P37", {}),
    ("2-filter-by-owner.png", "Pick an owner, the whole page follows", "Dashboard", "A1:P37", {"SelOwner": FILTER_OWNER}),
    ("3-task-table-overdue-flags.png", "Overdue tasks flag themselves in red", "Tasks", "A1:I24", {}),
)


def font(file_name: str, size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(str(FONTS / file_name), size)


def fit(image: Image.Image, max_width: int, max_height: int) -> Image.Image:
    scale = min(max_width / image.width, max_height / image.height)
    size = (max(1, round(image.width * scale)), max(1, round(image.height * scale)))
    return image.resize(size, Image.LANCZOS)


def autocrop(image: Image.Image, pad: int = 10) -> Image.Image:
    image = image.convert("RGB")
    box = ImageChops.difference(image, Image.new("RGB", image.size, "white")).getbbox()
    if box is None:
        return image
    return image.crop((max(box[0] - pad, 0), max(box[1] - pad, 0),
                       min(box[2] + pad, image.width), min(box[3] + pad, image.height)))


def compose(content: Image.Image, caption: str, out_path: Path) -> None:
    """Title strip on top, the content scaled to fill the rest, footnote bottom right."""
    canvas = Image.new("RGB", (WIDTH, HEIGHT), BACKGROUND)
    draw = ImageDraw.Draw(canvas)
    draw.rectangle((0, 0, WIDTH, STRIP_HEIGHT), fill=STRIP)
    draw.text((MARGIN + 8, STRIP_HEIGHT // 2), caption, font=font("seguisb.ttf", 50), fill="white", anchor="lm")

    left, top, right, bottom = CONTENT_BOX
    scaled = fit(content, right - left - 4, bottom - top - 4)
    card = Image.new("RGB", (scaled.width + 4, scaled.height + 4), "#C5CCD3")
    card.paste(scaled, (2, 2))
    canvas.paste(card, (left + (right - left - card.width) // 2, top + (bottom - top - card.height) // 2))

    draw.text((WIDTH - MARGIN, HEIGHT - 28), FOOTNOTE, font=font("segoeui.ttf", 22), fill="#8A94A0", anchor="rm")
    assert canvas.size == (WIDTH, HEIGHT)
    canvas.save(out_path, "PNG")
    print(f"  {out_path.name}  {canvas.width}x{canvas.height}")


def export_views() -> list[Path]:
    """Have Excel print each view to its own PDF. Nothing is saved to the workbook."""
    import win32com.client as win32

    WORK.mkdir(parents=True, exist_ok=True)
    pdfs = []
    excel = win32.DispatchEx("Excel.Application")  # a private instance, never the user's open Excel
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(str(WORKBOOK), 0, True)  # no link updates, read-only
        for file_name, _, sheet_name, address, filters in VIEWS:
            for name in ("SelOwner", "SelStatus", "SelYear"):
                workbook.Names(name).RefersToRange.Value = filters.get(name, "(All)")
            excel.CalculateFull()
            sheet = workbook.Worksheets(sheet_name)
            setup = sheet.PageSetup
            setup.Orientation = 2  # landscape
            setup.Zoom = False
            setup.FitToPagesWide = 1
            setup.FitToPagesTall = 1
            setup.PrintHeadings = sheet_name == "Tasks"
            setup.PrintGridlines = False
            pdf_path = WORK / (Path(file_name).stem + ".pdf")
            sheet.Range(address).ExportAsFixedFormat(0, str(pdf_path), 0, False, False)
            pdfs.append(pdf_path)
        workbook.Close(False)
    finally:
        excel.Quit()
    return pdfs


def pdf_image(pdf_path: Path, dpi: int = 300) -> Image.Image:
    with pymupdf.open(pdf_path) as document:
        pixmap = document[0].get_pixmap(dpi=dpi, alpha=False)
        return Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)


def main() -> int:
    SHOTS.mkdir(exist_ok=True)
    print("Excel is printing the views to PDF ...")
    pdfs = export_views()
    print("Writing screenshots:")
    for (file_name, caption, *_), pdf_path in zip(VIEWS, pdfs):
        compose(autocrop(pdf_image(pdf_path)), caption, SHOTS / file_name)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
