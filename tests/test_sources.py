import hashlib
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from openpyxl import Workbook

from ylookup.events import Events
from ylookup.sources import match_investor, parse_statement, prepare_sources, read_workbook


RULE = {"version": 1, "intent": "Remaining amount", "operation": "subtract",
        "left": "Amount Called", "right": "Amount Received", "output": "Amount outstanding"}


def statement(description="REF TAYLOR REE", currency="USD", kind="deposit"):
    header = f"{'Date':<15}{'Description':<55}{'Withdrawal':>14}{'Deposit':>14}{'Balance':>14}"
    debit, credit = ("", "95.00") if kind == "deposit" else ("95.00", "")
    record = f"{'24 Feb 2026':<15}{description:<55}{debit:>14}{credit:>14}{'900.00':>14}"
    return f"Currency : {currency}\n{header}\n{record}\n   MORE WRAPPED REFERENCE\n   Thank you for banking\n"


def workbook(path, renamed=False, missing=False, duplicate=False):
    book = Workbook()
    sheet = book.active
    sheet.title = "Partners" if renamed else "Roster"
    if renamed:
        sheet.append(["Capital Call Amount", "LP Name", "Partner ID"])
        sheet.append(["USD", None, None])
        sheet.append([None if missing else 100, "Taylor Reed", "001"])
        sheet.append([50, "Morgan Vale", "001" if duplicate else "002"])
    else:
        sheet.append(["ID", "Investor", "Amount Called"])
        sheet.append([None, None, "USD"])
        sheet.append(["001", "Taylor Reed", None if missing else 100])
        sheet.append(["001" if duplicate else "002", "Morgan Vale", 50])
    accounts = book.create_sheet("Lookup renamed" if renamed else "Accounts")
    accounts.append(["Account Code", "Transaction Type"])
    accounts.append(["00100.2", "Cash received"])
    constants = book.create_sheet("Fund constants")
    constants.append(["Legal Entity", "Bank Account", "Batch Type"])
    constants.append(["Example fund", "Example account", "General"])
    blank_output = book.create_sheet("Output template")
    blank_output.append(["Legal Entity", "Bank Account", "Batch Type"])
    book.save(path)
    book.close()


class SourcesTest(unittest.TestCase):
    def test_deposit_is_not_running_balance_and_wrapped_description_is_preserved(self):
        result = parse_statement([(1, statement())])
        self.assertEqual(result["questions"], [])
        receipt = result["transactions"][0]
        self.assertEqual(receipt["amount"], "95.00")
        self.assertEqual(receipt["balance"], "900.00")
        self.assertEqual(receipt["date"], "2026-02-24")
        self.assertIn("WRAPPED REFERENCE", receipt["description"])
        self.assertNotIn("Thank you", receipt["description"])

    def test_unsupported_layout_and_missing_amount_request_review(self):
        unsupported = parse_statement([(1, "scanned image without readable table")])
        self.assertTrue(any(q["code"] == "bank_layout" for q in unsupported["questions"]))
        missing = parse_statement([(1, statement().replace("95.00", "     "))])
        self.assertEqual(missing["transactions"], [])
        self.assertTrue(any(q["code"] == "receipt_amount" for q in missing["questions"]))

    def test_missing_receipt_date_does_not_merge_money_into_previous_description(self):
        undated = statement().splitlines()[2].replace("24 Feb 2026", "           ")
        text = statement().replace("   MORE WRAPPED REFERENCE", undated)
        result = parse_statement([(1, text)])
        self.assertEqual(len(result["transactions"]), 1)
        self.assertTrue(any(q["code"] == "receipt_date" for q in result["questions"]))

    def test_explicit_total_ends_roster_before_footer_notes(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "source.xlsx"
            workbook(path)
            from openpyxl import load_workbook
            book = load_workbook(path)
            sheet = book["Roster"]
            sheet.append([None, None, "=SUM(C3:C4)"])
            sheet.append(["Footer text", "Instructions", "Other content"])
            book.save(path)
            book.close()
            result = read_workbook(path)
            self.assertEqual(len(result["roster"]), 2)
            self.assertEqual(result["questions"], [])

    def test_identity_uses_evidence_not_first_match(self):
        roster = [{"id": "001", "name": "Taylor Reed"}, {"id": "002", "name": "Taylor Reeves"}]
        matches, method = match_investor("REF TAYLOR REE", roster)
        self.assertEqual(len(matches), 2)
        self.assertEqual(method, "first_name_and_surname_prefix")
        full, method = match_investor("REF Taylor Reed paid", roster)
        self.assertEqual([row["id"] for row in full], ["001"])
        self.assertEqual(method, "full_name")
        self.assertEqual(match_investor("someone unknown", roster)[0], [])

    def test_bindings_follow_renamed_tabs_headers_and_column_order(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "source.xlsx"
            workbook(path, renamed=True)
            before = hashlib.sha256(path.read_bytes()).digest()
            result = read_workbook(path)
            self.assertEqual(result["questions"], [])
            self.assertEqual(result["bindings"]["roster"]["sheet"], "Partners")
            self.assertEqual(result["bindings"]["roster"]["columns"]["id"], 3)
            self.assertEqual(result["roster"][0]["id"], "001")
            self.assertEqual(result["accounts"][0]["account"], "00100.2")
            self.assertEqual(hashlib.sha256(path.read_bytes()).digest(), before)

    def test_missing_amount_and_duplicate_id_are_not_silently_accepted(self):
        with TemporaryDirectory() as directory:
            for flag in ["missing", "duplicate"]:
                path = Path(directory) / (flag + ".xlsx")
                workbook(path, **{flag: True})
                self.assertTrue(read_workbook(path)["questions"])

    def test_ambiguous_roster_binding_requests_review(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "competing.xlsx"
            book = Workbook()
            for sheet in [book.active, book.create_sheet("Other roster")]:
                sheet.append(["ID", "Investor", "Amount Called"])
                sheet.append(["001", "Taylor Reed", 100])
            book.save(path)
            book.close()
            result = read_workbook(path)
            self.assertNotIn("roster", result["bindings"])
            self.assertTrue(any(q["code"] == "binding" for q in result["questions"]))

    def test_preparation_applies_saved_rule_and_retains_currency_question(self):
        with TemporaryDirectory() as directory:
            folder = Path(directory)
            path, bank = folder / "source.xlsx", folder / "statement.pdf"
            workbook(path)
            bank.write_bytes(b"mocked PDF test fixture")
            for currency, status in [("USD", "success"), ("EUR", "needs_input")]:
                events = Events(folder / currency)
                with patch("ylookup.sources.read_statement", return_value=parse_statement([(1, statement(currency=currency))])):
                    result = prepare_sources(path, bank, RULE, events)
                self.assertEqual(result["status"], status)
                evidence = json.loads((events.directory / "sources.json").read_text())
                self.assertEqual(evidence["tracker"][0]["outstanding"], "5.00")
                self.assertEqual(evidence["tracker"][1]["received"], "0.00")
                if currency == "EUR":
                    self.assertEqual(evidence["questions"][0]["code"], "currency_conflict")

    def test_unmatched_duplicate_repeated_receipts_and_withdrawals_block_tracker(self):
        with TemporaryDirectory() as directory:
            folder = Path(directory)
            path, bank = folder / "source.xlsx", folder / "statement.pdf"
            workbook(path)
            bank.write_bytes(b"mocked PDF test fixture")
            unmatched = parse_statement([(1, statement(description="UNRELATED NAME"))])
            duplicate = parse_statement([(1, statement()), (2, statement())])
            repeated = parse_statement([(1, statement()), (2, statement(description="SECOND TAYLOR REE"))])
            withdrawal = parse_statement([(1, statement(kind="withdrawal"))])
            for index, evidence in enumerate([unmatched, duplicate, repeated, withdrawal]):
                events = Events(folder / str(index))
                with patch("ylookup.sources.read_statement", return_value=evidence):
                    result = prepare_sources(path, bank, RULE, events)
                self.assertEqual(result["status"], "needs_input")
                self.assertEqual(result["tracker_records"], 0)

    def test_conflicting_currency_inside_reference_is_not_ignored(self):
        result = parse_statement([(1, statement().replace("   MORE WRAPPED REFERENCE", "   EUR|95.00|1.00"))])
        self.assertTrue(any(q["code"] == "receipt_currency" for q in result["questions"]))


if __name__ == "__main__":
    unittest.main()
