"""Show what the workbook files and screenshots say about who made them.

Usage:
    python tools\\check_metadata.py

For each .xlsx here it opens the file as the zip package it is and prints:
  - author and "last modified by" (docProps/core.xml);
  - Company and Manager (docProps/app.xml);
  - whether there are custom properties, printer settings, macros, or a stored
    local folder path;
  - whether the Windows user name of whoever runs this, or a "Users" folder path,
    appears in any part of the package. The user name itself is never printed.
It also checks the two workbook copies are identical and every screenshot is
1600x1200 with nothing embedded.
"""

from __future__ import annotations

import getpass
import hashlib
import re
import sys
import zipfile
from pathlib import Path

from PIL import Image

ROOT = Path(__file__).resolve().parent.parent
EXPECTED = "AZG Engineering"
SKIP = {".venv", ".git", "_work", "__pycache__"}
MAIN_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"


def files(pattern: str) -> list[Path]:
    return sorted(p for p in ROOT.rglob(pattern)
                  if not SKIP.intersection(p.relative_to(ROOT).parts) and not p.name.startswith("~$"))


def tag(xml: str, name: str) -> str | None:
    found = re.search(rf"<{name}(?: [^>]*)?>(.*?)</{name}>", xml, flags=re.S)
    return found.group(1) if found else None


def check_workbook(path: Path, user: bytes) -> int:
    problems = 0
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        parts = {name: package.read(name) for name in names}
    core = parts["docProps/core.xml"].decode("utf-8")
    app = parts["docProps/app.xml"].decode("utf-8")
    book = parts["xl/workbook.xml"].decode("utf-8")
    types = parts["[Content_Types].xml"].decode("utf-8")

    facts = [
        ("file type is .xlsx (no macros possible)", path.suffix.lower() == ".xlsx" and MAIN_TYPE in types, path.suffix),
        ("author", tag(core, "dc:creator") == EXPECTED, repr(tag(core, "dc:creator"))),
        ("last modified by", tag(core, "cp:lastModifiedBy") == EXPECTED, repr(tag(core, "cp:lastModifiedBy"))),
        ("Company", tag(app, "Company") in (None, "", EXPECTED), repr(tag(app, "Company"))),
        ("Manager", tag(app, "Manager") in (None, "", EXPECTED), repr(tag(app, "Manager"))),
        ("custom properties", "docProps/custom.xml" not in names, "none" if "docProps/custom.xml" not in names else "PRESENT"),
        ("printer settings parts", not any("printerSettings" in n for n in names),
         "none" if not any("printerSettings" in n for n in names) else "PRESENT"),
        ("macro project", not any("vba" in n.lower() for n in names),
         "none" if not any("vba" in n.lower() for n in names) else "PRESENT"),
        ("stored local folder path (absPath)", "absPath" not in book, "none" if "absPath" not in book else "PRESENT"),
    ]
    with_user = [n for n, data in parts.items() if user in data.lower()]
    with_path = [n for n, data in parts.items() if b"users\\" in data.lower() or b"users/" in data.lower()]
    facts.append((f"Windows user name in any of the {len(parts)} package parts", not with_user, with_user or "no part"))
    facts.append(('a "Users" folder path in any package part', not with_path, with_path or "no part"))

    print(path.relative_to(ROOT))
    for label, good, shown in facts:
        problems += not good
        print(f"  {'OK ' if good else 'BAD'} {label}: {shown}")
    return problems


def main() -> int:
    problems = 0
    user = getpass.getuser().lower().encode()

    workbooks = files("*.xlsx")
    for path in workbooks:
        problems += check_workbook(path, user)
    digests = {hashlib.sha256(p.read_bytes()).hexdigest() for p in workbooks}
    same = len(workbooks) == 2 and len(digests) == 1
    problems += not same
    print(f"{'OK ' if same else 'BAD'} the {len(workbooks)} workbook copies are byte-for-byte identical: {same}")

    for path in files("*.png"):
        with Image.open(path) as image:
            size, extra = image.size, sorted(image.info)
        leaked = user in path.read_bytes().lower()
        good = size == (1600, 1200) and not extra and not leaked
        problems += not good
        print(f"{'OK ' if good else 'BAD'} {path.relative_to(ROOT)}: {size[0]}x{size[1]}, "
              f"embedded text fields={extra or 'none'}, user name in file={'YES' if leaked else 'no'}")

    print("RESULT:", "all good" if not problems else f"{problems} problem(s)")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
