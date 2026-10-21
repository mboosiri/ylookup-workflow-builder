import hashlib
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook

from ylookup.cli import inspect_workbook


class InspectionTest(unittest.TestCase):
    def test_preview_ignores_formatted_empty_rows_and_preserves_source(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "source.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Renamed roster"
            sheet.append(["ID", "Investor", "Called"])
            sheet.append(["003", "Aaron Patel", "=100+5"])
            for index in range(8):
                sheet.append([f"record-{index}"])
            sheet.cell(100, 1).number_format = "0.00"
            workbook.save(path)
            workbook.close()
            before = hashlib.sha256(path.read_bytes()).digest()

            result = inspect_workbook(path)["sheets"][0]

            self.assertEqual(result["name"], "Renamed roster")
            self.assertEqual(result["reported_rows"], 100)
            self.assertEqual(result["populated_rows"], 10)
            self.assertEqual(len(result["preview"]), 8)
            self.assertEqual(result["preview"][1][0]["value"], "003")
            self.assertEqual(result["preview"][1][2]["value"], "=100+5")
            self.assertEqual(hashlib.sha256(path.read_bytes()).digest(), before)


if __name__ == "__main__":
    unittest.main()
