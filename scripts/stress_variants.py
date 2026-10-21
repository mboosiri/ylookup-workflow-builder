"""Write Phase 3 stress variants of the input workbook to data/stress/, keeping Excel's cached formula results.

Usage: python scripts/stress_variants.py data/import-capital-call-notice-diu-input.xlsx
"""

from pathlib import Path
import sys
import tempfile
import zipfile

from openpyxl import load_workbook

from ylookup.workbook import restore_caches


def edit(source, destination, change):
    with tempfile.TemporaryDirectory() as folder:
        raw = Path(folder) / "raw.xlsx"
        book = load_workbook(source)
        change(book)
        book.save(raw)
        restore_caches(source, raw, destination)


def rename_tabs(source, destination, names):
    # Rename in workbook.xml only, so formulas and their cached results stay untouched.
    with zipfile.ZipFile(source) as old, zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as new:
        for item in old.infolist():
            data = old.read(item.filename)
            if item.filename == "xl/workbook.xml":
                for before, after in names.items():
                    data = data.replace(f'name="{before}"'.encode(), f'name="{after}"'.encode())
            new.writestr(item, data)


def set_cells(*cells):
    def change(book):
        for sheet, cell, value in cells:
            book[sheet][cell] = value
    return change


def swap_constants(book):
    sheet = book["Investors"]
    for row in (1, 2):
        sheet[f"B{row}"].value, sheet[f"C{row}"].value = sheet[f"C{row}"].value, sheet[f"B{row}"].value


VARIANTS = {
    "columns-renamed": set_cells(("Capital Call Tracker", "D7", "Limited Partner"),
                                 ("Capital Call Tracker", "P7", "Paid Amount"), ("Investors", "C1", "Deal Name")),
    "columns-reordered": swap_constants,
    "missing-data": set_cells(("Capital Call Tracker", "D12", None), ("Investors", "C2", None)),
    "ambiguous-amounts": set_cells(("Capital Call Tracker", "P7", "Amount 1"), ("Capital Call Tracker", "Q7", "Amount 2")),
}


def main(source):
    folder = Path("data/stress")
    folder.mkdir(parents=True, exist_ok=True)
    tabs = {"Capital Call Tracker": "Call Workings", "Investors": "Fund Static"}
    rename_tabs(source, folder / "tabs-renamed.xlsx", tabs)
    for name, change in VARIANTS.items():
        edit(source, folder / f"{name}.xlsx", change)
    edit(source, folder / "combined-raw.xlsx", lambda book: (VARIANTS["columns-renamed"](book), swap_constants(book)))
    rename_tabs(folder / "combined-raw.xlsx", folder / "combined.xlsx", tabs)
    (folder / "combined-raw.xlsx").unlink()
    for path in sorted(folder.glob("*.xlsx")):
        print(path)


if __name__ == "__main__":
    main(Path(sys.argv[1]))
