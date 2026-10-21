from datetime import date
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
import time
import unittest

from openpyxl import Workbook

from ylookup.sources import (LEAVE, bind_tables, equivalent, find_tables, match_names, match_receipts, parse_statement,
                             rekey, sheet_rows)
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


class StatementTest(unittest.TestCase):
    def test_deposit_is_not_running_balance_and_wrapped_description_is_kept(self):
        result = parse_statement([(1, statement())])
        self.assertEqual(result["questions"], [])
        receipt = result["transactions"][0]
        self.assertEqual((receipt["amount"], receipt["balance"], receipt["date"]), ("95.00", "900.00", "2026-02-24"))
        self.assertEqual(receipt["description"], "REF TAYLOR REE MORE WRAPPED REFERENCE")

    def test_unsupported_layout_missing_amount_and_undated_rows_ask(self):
        self.assertTrue(any(q["code"] == "bank_layout" for q in parse_statement([(1, "scanned image")])["questions"]))
        missing = parse_statement([(1, statement().replace("95.00", "     "))])
        self.assertEqual(missing["transactions"], [])
        undated = statement().splitlines()[2].replace("24 Feb 2026", "           ")
        result = parse_statement([(1, statement().replace("   MORE WRAPPED REFERENCE", undated))])
        self.assertEqual(len(result["transactions"]), 1)
        self.assertTrue(any(q["code"] == "receipt_date" for q in result["questions"]))
        conflict = parse_statement([(1, statement().replace("   MORE WRAPPED REFERENCE", "   EUR|95.00|1.00"))])
        self.assertTrue(any(q["code"] == "receipt_currency" for q in conflict["questions"]))


class MatchTest(unittest.TestCase):
    def table(self, *names):
        found = {"sheet": "R", "row": 1, "headers": {0: "Name"}, "key": "Name", "units": {}, "orphans": [], "post": [], "rows": len(names),
                 "records": [{"row": i + 2, "cells": [n], "formulas": [None]} for i, n in enumerate(names)],
                 "all": []}
        return rekey(found, {"Name": 0})

    def test_names_match_in_full_or_by_first_name_and_surname_prefix(self):
        names = {0: "Taylor Reed", 1: "Taylor Reeves"}
        self.assertEqual(match_names("REF TAYLOR REE", names), ([0, 1], "first_name_and_surname_prefix"))
        self.assertEqual(match_names("REF Taylor Reed paid", names), ([0], "full_name"))
        self.assertEqual(match_names("REF TAYLOR REE", names, "full_name_only"), ([], "full_name"))

    def test_doubt_becomes_a_question_and_an_answer_settles_it(self):
        table = self.table("Taylor Reed", "Taylor Reeves")
        result = match_receipts(parse_statement([(1, statement())]), table, "Name", "full_or_unique_prefix", {})
        item = result["questions"][0]
        self.assertEqual((item["code"], item["options"]), ("receipt_match", ["Taylor Reed", "Taylor Reeves", LEAVE]))
        settled = match_receipts(parse_statement([(1, statement())]), table, "Name", "full_or_unique_prefix", {item["id"]: "Taylor Reeves"})
        self.assertEqual(settled["questions"], [])
        self.assertEqual(settled["receipts"], {1: {"amount": Decimal("95.00"), "description": "REF TAYLOR REE MORE WRAPPED REFERENCE",
                                                   "date": date(2026, 2, 24)}})
        left = match_receipts(parse_statement([(1, statement())]), table, "Name", "full_or_unique_prefix", {item["id"]: LEAVE})
        self.assertEqual((left["receipts"], left["deposits"][0]["record"]), ({}, None))

    def test_withdrawals_duplicates_and_repeat_payers_are_questions(self):
        table = self.table("Taylor Reed", "Morgan Vale")
        cases = {"withdrawal": [(1, statement(kind="withdrawal"))],
                 "duplicate_transaction": [(1, statement()), (2, statement())]}
        for code, pages in cases.items():
            self.assertEqual([q["code"] for q in match_receipts(parse_statement(pages), table, "Name", "full_or_unique_prefix", {})["questions"]], [code])
        later = statement().replace("24 Feb 2026", "25 Feb 2026").replace("REF TAYLOR REE", "SECOND TAYLOR REE")
        twice = match_receipts(parse_statement([(1, statement()), (2, later)]), table, "Name", "full_or_unique_prefix", {})
        self.assertEqual(twice["questions"][0]["options"], ["2026-02-24", "2026-02-25"])
        self.assertEqual(twice["receipts"][0]["amount"], Decimal("190.00"))
        stop = {"withdrawal:2026-02-24 withdrawal 95.00": "Stop for review"}
        with self.assertRaisesRegex(ValueError, "Stopped by a recorded answer"):
            match_receipts(parse_statement([(1, statement(kind="withdrawal"))]), table, "Name", "full_or_unique_prefix", stop)


if __name__ == "__main__":
    unittest.main()
