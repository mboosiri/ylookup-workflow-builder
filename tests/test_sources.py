from datetime import date
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
import time
import unittest

from openpyxl import Workbook

from ylookup.sources import (LEAVE, bind_tables, equivalent, find_tables, link_rows, match_names, pdf_tables, rekey,
                             sheet_rows)
from fixtures import statement


def table_book(path, far_cell=None):
    book = Workbook()
    sheet = book.active
    sheet.title = "Roster"
    sheet.append(["Notes <== read me"])
    sheet.append(["ID", "Investor", "Called"])
    sheet.append([None, None, "USD"])
    sheet.append(["001", "Taylor Reed", 100])
    sheet.append([None, None, 7])
    sheet.append(["002", "Morgan Vale", 50])
    sheet.append([None, None, "=SUM(C4:C6)"])
    if far_cell:
        sheet[far_cell] = "stray"
    other = book.create_sheet("Lookup")
    other.append(["Code", "Type"])
    other.append(["C1", "Cash"])
    book.save(path)
    book.close()


class TablesTest(unittest.TestCase):
    def test_tables_are_header_rows_with_records_units_totals_and_orphans(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "book.xlsx"
            table_book(path)
            roster, lookup = find_tables(sheet_rows(path))
        self.assertEqual((roster["sheet"], roster["row"], roster["key"]), ("Roster", 2, "ID"))
        self.assertEqual([r["row"] for r in roster["records"]], [4, 6])
        self.assertEqual(roster["units"], {"Called": "USD"})
        self.assertEqual(roster["orphans"], [5])
        self.assertEqual(roster["post"][0]["formulas"][2], "=SUM(C4:C6)")
        self.assertEqual(len(lookup["records"]), 1)

    def test_reading_stops_after_a_long_blank_run_so_inflated_sheets_stay_fast(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "book.xlsx"
            table_book(path, far_cell="A40000")
            started = time.perf_counter()
            rows = sheet_rows(path)
            self.assertLess(time.perf_counter() - started, 2)
        self.assertLess(max(number for number, *_ in rows["Roster"]), 100)

    def test_tables_bind_by_header_fingerprint_not_tab_name_or_position(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "book.xlsx"
            table_book(path)
            found = find_tables(sheet_rows(path))
        spec = {"outputs": {}, "tables": {
            "Partners": {"sheet": "Old tab", "headers": ["ID", "Investor", "Called", "Gone"], "columns": {
                "Investor": {"header": "Investor", "types": ["text"]}, "Called": {"header": "Called", "types": ["number"]}}},
            "Codes": {"sheet": "Lookup", "headers": ["Account Code", "Movement"], "columns": {
                "Account Code": {"header": "Account Code", "types": ["text"]}}}}}
        bound, problems = bind_tables(spec, found)
        self.assertEqual(bound["Partners"]["sheet"], "Roster")
        self.assertEqual(bound["Partners"]["records"][0]["values"], {"Investor": "Taylor Reed", "Called": 100})
        # Every header renamed, but the sheet name still identifies the table; the column is then asked for.
        self.assertEqual(bound["Codes"]["sheet"], "Lookup")
        self.assertEqual([(p["code"], p["options"]) for p in problems], [("column_mapping", ["Code", "Type"])])
        answer = {problems[0]["id"]: "Code"}
        bound, problems = bind_tables(spec, found, answer)
        self.assertEqual((problems, bound["Codes"]["records"][0]["values"]), ([], {"Account Code": "C1"}))
        wrong_type = {"column_mapping:Partners.Called": "Investor"}
        spec["tables"]["Partners"]["columns"]["Called"]["header"] = "Amount Called"
        _, problems = bind_tables(spec, found, {**answer, **wrong_type})
        self.assertIn("was rejected", problems[0]["question"])

    def test_text_comparison_is_exact_for_codes_and_tolerant_for_long_references(self):
        self.assertFalse(equivalent("10000.2", "100002"))
        self.assertFalse(equivalent("Yes", "yes"))
        self.assertFalse(equivalent(1, "1"))
        self.assertFalse(equivalent(None, 0))
        self.assertTrue(equivalent(Decimal("95"), 95.0))
        self.assertTrue(equivalent("", None))
        long = "TT|03MX260209783088  RDO9 AARON PAT PDF EL|CHASE BANK SG74645658765005HS DE14704456781000045388|SUH68Y90JG95"
        self.assertTrue(equivalent(long, long.replace("|", " ").replace("RDO9", "RD O9")))


def pdf_table(*texts):
    tables, questions = pdf_tables(list(texts), "bank")
    return (tables[0] if tables else None), questions


class PdfTableTest(unittest.TestCase):
    def test_columns_come_from_the_header_and_wrapped_text_stays_with_its_row(self):
        table, questions = pdf_table(statement())
        self.assertEqual(questions, [])
        self.assertEqual(list(table["headers"].values()), ["Date", "Description", "Withdrawal", "Deposit", "Balance"])
        self.assertEqual(table["records"][0]["cells"], [date(2026, 2, 24), "REF TAYLOR REE MORE WRAPPED REFERENCE", None,
                                                         Decimal("95.00"), Decimal("900.00")])
        self.assertEqual(table["units"], {"Deposit": "USD", "Balance": "USD"})

    def test_any_column_names_and_date_format_work(self):
        text = (f"{'Booked':<12}{'Payer reference':<40}{'Amount in':>12}{'Running total':>16}\n"
                f"{'2026-03-01':<12}{'MORGAN VALE INV 7':<40}{'50.00':>12}{'950.00':>16}\n")
        table, _ = pdf_table(text)
        self.assertEqual(list(table["headers"].values()), ["Booked", "Payer reference", "Amount in", "Running total"])
        self.assertEqual(table["records"][0]["cells"], [date(2026, 3, 1), "MORGAN VALE INV 7", Decimal("50.00"), Decimal("950.00")])

    def test_unreadable_layouts_and_undated_amounts_ask(self):
        self.assertEqual(pdf_table("scanned image")[1][0]["code"], "source_tables")
        undated = statement().splitlines()[2].replace("24 Feb 2026", "           ")
        table, questions = pdf_table(statement().replace("   MORE WRAPPED REFERENCE", undated))
        self.assertEqual(len(table["records"]), 1)
        self.assertEqual([q["code"] for q in questions], ["source_row"])

    def test_amounts_in_another_form_ask_instead_of_joining_the_text(self):
        for amount in ("€95.00", "EUR 95.00", "1.095,00", "1,095", "95.00 CR", "(95.00)"):
            with self.subTest(amount=amount):
                table, questions = pdf_table(statement(amount=amount))
                self.assertEqual([q["code"] for q in questions], ["source_amount"])
                self.assertNotIn(amount, table["records"][0]["cells"][1])
        table, questions = pdf_table(statement(description="INV 2026  REF 12345"))
        self.assertEqual(questions, [])
        self.assertEqual(table["records"][0]["cells"][1], "INV 2026 REF 12345 MORE WRAPPED REFERENCE")


class LinkTest(unittest.TestCase):
    def table(self, column, *values):
        found = {"sheet": "T", "row": 1, "headers": {0: column}, "key": None, "units": {}, "orphans": [], "post": [],
                 "rows": len(values), "records": [{"row": i + 2, "cells": [v], "formulas": [None]} for i, v in enumerate(values)],
                 "all": []}
        return rekey(found, {column: 0})

    def test_names_match_in_full_or_by_first_name_and_surname_prefix(self):
        names = {0: "Taylor Reed", 1: "Taylor Reeves"}
        self.assertEqual(match_names("REF TAYLOR REE", names), ([0, 1], "first_name_and_surname_prefix"))
        self.assertEqual(match_names("REF Taylor Reed paid", names), ([0], "full_name"))
        self.assertEqual(match_names("REF TAYLOR REE", names, "full_name_only"), ([], "full_name"))

    def test_doubt_becomes_a_question_and_an_answer_settles_it(self):
        rows, people = self.table("Text", "REF TAYLOR REE", "MORGAN VALE"), self.table("Name", "Taylor Reed", "Taylor Reeves", "Morgan Vale")
        result = link_rows(rows, "Text", people, "Name", "full_or_unique_prefix", {})
        self.assertEqual(result["rows"], {0: None, 1: 2})
        item = result["questions"][0]
        self.assertEqual(item["options"], ["Taylor Reed", "Taylor Reeves", LEAVE])
        settled = link_rows(rows, "Text", people, "Name", "full_or_unique_prefix", {item["id"]: "Taylor Reeves"})
        self.assertEqual((settled["rows"], settled["records"], settled["questions"]), ({0: 1, 1: 2}, {1: [0], 2: [1]}, []))
        left = link_rows(rows, "Text", people, "Name", "full_or_unique_prefix", {item["id"]: LEAVE})
        self.assertEqual((left["rows"][0], left["questions"]), (None, []))


if __name__ == "__main__":
    unittest.main()
