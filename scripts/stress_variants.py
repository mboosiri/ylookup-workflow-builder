"""Write Phase 3 stress variants of the input workbook to data/stress/, keeping Excel's cached formula results.

Usage: python scripts/stress_variants.py data/capital-call-input.xlsx
"""

from pathlib import Path
import sys
import tempfile
import zipfile

from openpyxl import load_workbook

from ylookup.workbook import restore_caches


TRACKER, FUND, CHART = "Capital Call Tracker", "Investors", "PCS Chart of Accounts (Global)"


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


def swap_columns(sheet, left, right, rows):
    def change(book):
        for row in rows:
            a, b = book[sheet][f"{left}{row}"], book[sheet][f"{right}{row}"]
            a.value, b.value = b.value, a.value
    return change


def both(*changes):
    def change(book):
        for each in changes:
            each(book)
    return change


def extra_investor(book):
    sheet = book[TRACKER]
    values = {"B": "011", "C": "Series 1", "D": "Olivia Stone", "E": 1000000, "J": 298755.19}
    for column, value in values.items():
        sheet[f"{column}19"] = value
    for column in "FGHIKLMN":
        sheet[f"{column}19"] = None
    sheet["J20"] = "=SUM(J9:J19)"


VARIANTS = {
    # The original set.
    "columns-renamed": set_cells((TRACKER, "D7", "Limited Partner"), (TRACKER, "P7", "Paid Amount"), (FUND, "C1", "Deal Name")),
    "columns-reordered": swap_columns(FUND, "B", "C", (1, 2)),
    "missing-data": set_cells((TRACKER, "D12", None), (FUND, "C2", None)),
    "ambiguous-amounts": set_cells((TRACKER, "P7", "Amount 1"), (TRACKER, "Q7", "Amount 2")),
    # Harder ones: the review's adapt gaps, a renamed lookup table, a reshuffled tracker, new data.
    "constants-renamed": set_cells((FUND, "A1", "Fund Entity"), (FUND, "F1", "Cash Account"), (FUND, "E1", "Journal Batch Type")),
    "chart-renamed": set_cells((CHART, "B1", "Ledger Code"), (CHART, "D1", "Movement Type")),
    "tracker-reshuffled": both(swap_columns(TRACKER, "C", "D", range(7, 20)),
                               set_cells((TRACKER, "D7", "Series"), (TRACKER, "O7", "Value Date"))),
    "distractor-column": set_cells((TRACKER, "J7", "Call #8 - Total Call"), (TRACKER, "R7", "Call #8 - Estimate"),
                                   *[(TRACKER, f"R{row}", 1000.0 * row) for row in range(9, 19)]),
    "duplicate-name": set_cells((TRACKER, "D18", "Aaron Patel")),
    "unknown-payer": set_cells((TRACKER, "D15", "Dana Brookfield")),
    "extra-investor": extra_investor,
}
TABS = {"chart-tab-renamed": ("chart-renamed", {CHART: "Ledger Map"}),
        "tabs-renamed": (None, {TRACKER: "Call Workings", FUND: "Fund Static"})}


def main(source):
    folder = Path("data/stress")
    folder.mkdir(parents=True, exist_ok=True)
    for name, change in VARIANTS.items():
        edit(source, folder / f"{name}.xlsx", change)
    for name, (base, tabs) in TABS.items():
        rename_tabs(folder / f"{base}.xlsx" if base else source, folder / f"{name}.xlsx", tabs)
    edit(source, folder / "combined-raw.xlsx", both(VARIANTS["columns-renamed"], VARIANTS["columns-reordered"]))
    rename_tabs(folder / "combined-raw.xlsx", folder / "combined.xlsx", TABS["tabs-renamed"][1])
    (folder / "combined-raw.xlsx").unlink()
    for path in sorted(folder.glob("*.xlsx")):
        print(path)


if __name__ == "__main__":
    main(Path(sys.argv[1]))
