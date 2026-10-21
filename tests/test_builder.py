import copy
from datetime import datetime
import hashlib
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
import xml.etree.ElementTree as ET
from unittest.mock import patch
import zipfile

from openpyxl import Workbook, load_workbook

from ylookup.builder import build_workbook
from ylookup.events import Events
from ylookup.planner import CHECKS, accepted_questions, execute_plan, required_fields, run_workflow, validate_plan
from ylookup.sources import parse_statement
from ylookup.workbook import HEADERS, equivalent, write_workbook
from test_sources import statement


def reference_plan():
    field = lambda path: {"field": path}
    line = lambda left, right: {"line": [left, right]}
    types = line("Cash - Received - USD", "Due From Investors")
    comment = {"concat": ["Capital Call Received from ", field("tracker.investor")]}
    fields = {
        "Legal Entity": field("fund.entity"), "Legal Entity Domain": field("fund.entity_domain"),
        "Deal": field("fund.deal"), "Position": field("fund.position"), "Batch Index": 1,
        "Batch Type": field("fund.batch_type"), "JE Index": field("journal.index"),
        "Trans Index": field("line.index"), "GL Date": field("tracker.receipt_date"),
        "Effective Date": field("tracker.receipt_date"), "Transaction Currency": field("currency"),
        "Allocation Rule": line("Non-Dominant", "User Provided"), "Transaction Type": types,
        "Account": {"account": "transaction.type"}, "Is Debit": line("Yes", "No"),
        "Amount (Local)": field("tracker.received"), "Amount (LE)": field("tracker.received"),
        "Transaction Comments": comment, "Batch Comments": comment,
        "Bank Account": line(field("fund.bank_account"), None),
        "Investor": line(None, field("tracker.investor")), "Investor Domain": None,
        "Vehicle": line(None, field("fund.entity")),
        "Investor Amount (LE)": line(None, field("tracker.received")),
        "Investor Amount (Local)": line(None, field("tracker.received")),
        "Transaction GL Reference": field("transaction.ref")}
    return {"version": 2, "intent": "Reconcile deposits and prepare the import", "steps": [
        {"op": op, "intent": intent} for op, intent in [
            ("prepare_sources", "Read current source evidence"),
            ("emit_journal", "Create balanced entries from receipts"),
            ("write_workbook", "Write and check the new workbook")]],
        "numeric_rule": {"version": 1, "intent": "Calculate the unpaid amount", "operation": "subtract",
                         "left": "called", "right": "received", "output": "outstanding"},
        "matching_policy": "full_or_unique_prefix",
        "journal": {"include_unpaid": True, "currency_source": "tracker", "fields": fields},
        "checks": sorted(CHECKS)}


def fixtures(folder):
    source, target, bank, context = [folder / name for name in ["input.xlsx", "target.xlsx", "bank.pdf", "context.txt"]]
    book = Workbook()
    sheet = book.active
    sheet.title = "Roster"
    sheet.append(["ID", "Investor", "Amount Called", "Receipt Date", "Amount Received", "Amount outstanding"])
    sheet.append([None, None, "USD"])
    sheet.append(["001", "Taylor Reed", 100])
    sheet.append(["002", "Morgan Vale", 50])
    sheet.append([None, None, "=SUM(C3:C4)"])
    account = book.create_sheet("Accounts")
    account.append(["Account Code", "Transaction Type"])
    account.append(["CASH-CODE", "Cash - Received - USD"])
    account.append(["RECEIVABLE-CODE", "Due From Investors"])
    constants = book.create_sheet("Fund")
    constants.append(["Legal Entity", "Legal Entity Domain", "Deal", "Position", "Bank Account", "Batch Type"])
    constants.append(["Example private fund", "Private domain", "Private deal", "Private position", "Private bank account", "General"])
    stage = book.create_sheet("Staging")
    stage.append(["Date", "Line Item Description", "Withdrawal", "Deposit", "Matched Investor"])
    output = book.create_sheet("Import")
    output.append(HEADERS)
    book.save(source)
    # Independently specified expected workbook, rather than using the executor as oracle.
    sheet.cell(3, 4, datetime(2026, 2, 24))
    sheet.cell(3, 5, 95)
    sheet.cell(3, 6, 5)
    sheet.cell(4, 5, 0)
    sheet.cell(4, 6, 50)
    stage.append([datetime(2026, 2, 24), "REF TAYLOR REE MORE WRAPPED REFERENCE", None, 95, "Taylor Reed"])
    for index, name, amount, paid_date in [(1, "Taylor Reed", 95, datetime(2026, 2, 24)), (2, "Morgan Vale", 0, None)]:
        comment = "Capital Call Received from " + name
        common = ["Example private fund", "Private domain", "Private deal", "Private position", 1, "General", index]
        output.append(common + [1, paid_date, paid_date, "USD", "Non-Dominant", "Cash - Received - USD", "CASH-CODE", "Yes", amount, amount, comment, comment, "Private bank account", None, None, None, None, None, index * 2 - 1])
        output.append(common + [2, paid_date, paid_date, "USD", "User Provided", "Due From Investors", "RECEIVABLE-CODE", "No", amount, amount, comment, comment, None, name, None, "Example private fund", amount, amount, index * 2])
    book.save(target)
    book.close()
    bank.write_bytes(b"PDF content replaced by a parsed fixture in this test")
    context.write_text("Reconcile current deposits to investors and prepare balanced journal entries.")
    return source, target, bank, context


class FakePlanner:
    model = "test-model"
    prompt_version = "test-full-v1"
    max_tokens = 3500

    def __init__(self, plans):
        self.plans = iter(plans)
        self.messages = []

    def complete(self, messages, events, attempt):
        self.messages.append(copy.deepcopy(messages))
        return json.dumps({"status": "ready", "workflow": next(self.plans)})


class FullBuildTest(unittest.TestCase):
    def test_build_repairs_target_difference_and_writes_passing_workbook(self):
        with TemporaryDirectory() as directory:
            folder = Path(directory)
            source, target, bank, context = fixtures(folder)
            originals = [hashlib.sha256(p.read_bytes()).digest() for p in [source, target, bank]]
            wrong = reference_plan()
            wrong["journal"]["fields"]["Batch Index"] = 2
            client = FakePlanner([wrong, reference_plan()])
            events = Events(folder / "run")
            with patch("ylookup.sources.read_statement", return_value=parse_statement([(1, statement())])):
                result = build_workbook(source, target, bank, context, {}, client, events)
            self.assertEqual(result["status"], "success")
            self.assertEqual(result["attempts"], 2)
            self.assertEqual(result["journal_rows"], 4)
            self.assertIn("Batch Index", client.messages[1][-1]["content"])
            self.assertTrue(json.loads((events.directory / "candidate-1/differences.json").read_text()))
            self.assertEqual(json.loads((events.directory / "candidate-2/differences.json").read_text()), [])
            saved = json.loads(Path(result["workflow"]).read_text())
            self.assertEqual(saved["plan"], reference_plan())
            self.assertIn("Legal Entity", saved["checks"]["required_fields"]["1"])
            self.assertEqual(saved["bindings"]["roster"]["sheet"], "Roster")
            self.assertEqual([hashlib.sha256(p.read_bytes()).digest() for p in [source, target, bank]], originals)
            book = load_workbook(result["output"], data_only=True)
            self.assertEqual(book["Roster"]["F3"].value, 5)
            self.assertEqual(book["Roster"]["F4"].value, 50)
            self.assertEqual(book["Roster"]["E5"].value, 95)
            book.close()
            sent = json.dumps(client.messages)
            for private in ["Taylor Reed", "Morgan Vale", "Private bank account", "Example private fund", "CASH-CODE", "2026-02-24", "REF TAYLOR"]:
                self.assertNotIn(private, sent)

    def test_fields_matching_the_target_are_locked_against_regression(self):
        with TemporaryDirectory() as directory:
            folder = Path(directory)
            source, target, bank, context = fixtures(folder)
            wrong = reference_plan()
            wrong["journal"]["fields"]["Batch Index"] = 2
            regressed = reference_plan()
            regressed["journal"]["fields"]["Allocation Rule"] = "Non-Dominant"
            client = FakePlanner([wrong, regressed])
            events = Events(folder / "run")
            with patch("ylookup.sources.read_statement", return_value=parse_statement([(1, statement())])):
                result = build_workbook(source, target, bank, context, {}, client, events)
            self.assertEqual(result["status"], "success")
            self.assertEqual(result["attempts"], 2)
            feedback = json.loads(client.messages[1][-1]["content"])
            self.assertIn("Allocation Rule", feedback["locked_fields"])
            self.assertNotIn("Batch Index", feedback["locked_fields"])
            trace = [json.loads(line) for line in events.path.read_text().splitlines()]
            self.assertIn({"event": "locked_fields", "attempt": 2, "restored": ["Allocation Rule"]},
                          [{k: e[k] for k in ["event", "attempt", "restored"]} for e in trace if e["event"] == "locked_fields"])

    def test_invalid_journal_receives_field_diffs_before_any_workbook_is_written(self):
        with TemporaryDirectory() as directory:
            folder = Path(directory)
            source, target, bank, context = fixtures(folder)
            wrong = reference_plan()
            wrong["journal"]["fields"]["Investor"] = {"field": "tracker.investor"}
            wrong["journal"]["fields"]["Amount (Local)"] = "unknown"
            wrong["journal"]["fields"]["Transaction Comments"] = {"concat": ["prefix", None]}
            client = FakePlanner([wrong, reference_plan()])
            with patch("ylookup.sources.read_statement", return_value=parse_statement([(1, statement())])):
                result = build_workbook(source, target, bank, context, {}, client, Events(folder / "run"))
            self.assertEqual(result["status"], "success")
            self.assertFalse((folder / "run/candidate-1/output.xlsx").exists())
            feedback = json.loads(client.messages[1][-1]["content"])
            self.assertIn("validation_error", feedback)
            self.assertIn("concat received null", feedback["validation_error"])
            self.assertIn("tracker.investor", feedback["validation_error"])
            self.assertTrue(any(d["field"] == "Investor" for d in feedback["differences"]))
            self.assertNotIn("Taylor Reed", json.dumps(feedback))

    def test_currency_question_blocks_before_model_and_pair_scoped_answer_is_persisted(self):
        with TemporaryDirectory() as directory:
            folder = Path(directory)
            source, target, bank, context = fixtures(folder)
            client = FakePlanner([reference_plan()])
            answer = {"currency_conflict": {"decision": "reproduce_tracker_amounts_without_fx", "statement_currency": "EUR", "tracker_currency": "USD", "reason": "Explicit exercise assumption; no FX conversion."}}
            with patch("ylookup.sources.read_statement", return_value=parse_statement([(1, statement(currency="EUR"))])):
                blocked = build_workbook(source, target, bank, context, {}, client, Events(folder / "blocked"))
                self.assertEqual(blocked["status"], "needs_input")
                self.assertEqual(client.messages, [])
                self.assertFalse((folder / "blocked/workflow.json").exists())
                result = build_workbook(source, target, bank, context, answer, client, Events(folder / "accepted"))
            self.assertEqual(result["status"], "success")
            self.assertEqual(json.loads(Path(result["workflow"]).read_text())["accepted_answers"], answer)
            evidence = {"questions": [{"code": "currency_conflict"}], "statement_currency": "GBP", "tracker_currency": "USD"}
            self.assertEqual(len(accepted_questions(evidence, answer)), 1)

    def test_unusable_target_or_numeric_fund_identifier_stops_before_model(self):
        with TemporaryDirectory() as directory:
            folder = Path(directory)
            source, target, bank, context = fixtures(folder)
            client = FakePlanner([])
            with patch("ylookup.sources.read_statement", return_value=parse_statement([(1, statement())])):
                with self.assertRaisesRegex(ValueError, "no journal example"):
                    build_workbook(source, source, bank, context, {}, client, Events(folder / "empty"))
                book = load_workbook(source)
                book["Fund"]["E2"] = 12345
                book.save(source)
                book.close()
                with self.assertRaisesRegex(ValueError, "requires a text value"):
                    build_workbook(source, target, bank, context, {}, client, Events(folder / "numeric"))
            self.assertEqual(client.messages, [])

    def test_structured_model_inquiry_stops_without_saving_a_workflow(self):
        with TemporaryDirectory() as directory:
            folder = Path(directory)
            source, target, bank, context = fixtures(folder)
            client = FakePlanner([])
            with patch("ylookup.sources.read_statement", return_value=parse_statement([(1, statement())])):
                with patch.object(client, "complete", return_value=json.dumps({"status": "needs_input", "workflow": None, "question": "Which grouping policy applies?"})):
                    result = build_workbook(source, target, bank, context, {}, client, Events(folder / "run"))
            self.assertEqual(result["status"], "needs_input")
            self.assertIn("grouping", result["question"])
            self.assertFalse((folder / "run/workflow.json").exists())

    def test_missing_plan_fields_weakened_checks_and_literal_lookup_are_rejected(self):
        for mutation in [lambda p: p["checks"].remove("receipt_totals"),
                         lambda p: p["journal"]["fields"].pop("Account"),
                         lambda p: p["journal"]["fields"].update({"Account": "CASH-CODE"}),
                         lambda p: p["journal"]["fields"].update({"Legal Entity": "CONST_entity"}),
                         lambda p: p["journal"]["fields"].update({"Legal Entity": "{field:fund.entity}"}),
                         lambda p: p.update({"accepted_answers": {}})]:
            plan = reference_plan()
            mutation(plan)
            with self.assertRaises(ValueError):
                validate_plan(plan)

    def test_validation_reports_multiple_errors_and_blank_is_not_zero(self):
        plan = reference_plan()
        plan["steps"].append({"op": "check_balance", "intent": "Invalid extra tool"})
        plan["journal"]["fields"]["Account"] = "ACCOUNT_1"
        plan["journal"]["fields"]["Amount (Local)"] = {"field": "tracker.outstanding"}
        with self.assertRaises(ValueError) as raised:
            validate_plan(plan)
        for detail in ["steps:", "Account:", "Amount (Local):"]:
            self.assertIn(detail, str(raised.exception))
        self.assertFalse(equivalent("", 0, "received"))
        self.assertFalse(equivalent(None, 0, "received"))
        self.assertTrue(equivalent("", None, "received"))

    def test_balanced_but_wrong_amounts_and_invalid_operation_order_fail(self):
        with TemporaryDirectory() as directory:
            folder = Path(directory)
            source, _, bank, _ = fixtures(folder)
            for index, plan in enumerate([reference_plan(), reference_plan()]):
                if index == 0:
                    plan["journal"]["fields"].update({"Amount (Local)": 0, "Amount (LE)": 0})
                else:
                    plan["steps"][0], plan["steps"][1] = plan["steps"][1], plan["steps"][0]
                run = folder / str(index)
                events = Events(run)
                with patch("ylookup.sources.read_statement", return_value=parse_statement([(1, statement())])):
                    with self.assertRaises(ValueError):
                        execute_plan(plan, source, bank, run, {}, events)
                self.assertFalse((run / "output.xlsx").exists())

    def test_budget_exhaustion_does_not_save_a_runnable_workflow(self):
        with TemporaryDirectory() as directory:
            folder = Path(directory)
            source, target, bank, context = fixtures(folder)
            wrong = reference_plan()
            wrong["journal"]["include_unpaid"] = False
            with patch("ylookup.sources.read_statement", return_value=parse_statement([(1, statement())])):
                result = build_workbook(source, target, bank, context, {}, FakePlanner([wrong]), Events(folder / "run"), max_attempts=1)
            self.assertEqual(result["status"], "budget_exhausted")
            self.assertFalse((folder / "run/workflow.json").exists())

    def test_run_stops_on_missing_source_value_and_reports_moved_tables(self):
        with TemporaryDirectory() as directory:
            folder = Path(directory)
            source, target, bank, _ = fixtures(folder)
            book = load_workbook(target, data_only=True)
            journal = [dict(zip(HEADERS, row)) for row in book["Import"].iter_rows(min_row=2, values_only=True)]
            book.close()
            required = required_fields(journal)
            self.assertIn("Deal", required["1"])
            self.assertNotIn("GL Date", required["1"])  # non-payers have no receipt date
            self.assertIn("Investor", required["2"])
            bindings = {"constants": {"sheet": "Fund", "header_row": 1, "columns": {"entity": 1, "bank_account": 5, "batch_type": 6}}}
            workflow = folder / "workflow.json"
            workflow.write_text(json.dumps({"schema_version": 2, "plan": reference_plan(), "accepted_answers": {},
                                            "checks": {"required_fields": required}, "bindings": bindings}))
            book = load_workbook(source)
            book["Fund"].title = "Fund Static"
            book.save(source)
            book.close()
            with patch("ylookup.sources.read_statement", return_value=parse_statement([(1, statement())])):
                moved = run_workflow(workflow, source, bank, folder / "moved")
                self.assertEqual(moved["status"], "success")
                self.assertEqual(moved["journal_rows"], 4)
                self.assertIn("constants: tab 'Fund' -> 'Fund Static'", moved["binding_changes"])
                book = load_workbook(source)
                book["Fund Static"]["C2"] = None
                book.save(source)
                book.close()
                blank = run_workflow(workflow, source, bank, folder / "blank")
            self.assertEqual(blank["status"], "needs_input")
            self.assertEqual({(q["field"], q["line"]) for q in blank["questions"]}, {("Deal", 1), ("Deal", 2)})
            self.assertFalse((folder / "blank/output.xlsx").exists())
            self.assertTrue((folder / "blank/questions.json").exists())

    def test_output_preserves_unchanged_formula_cache_and_refuses_overwrite(self):
        with TemporaryDirectory() as directory:
            folder = Path(directory)
            source, _, bank, _ = fixtures(folder)
            book = load_workbook(source)
            book.create_sheet("Reference")["A1"] = "=1+2"
            book.save(source)
            book.close()
            with zipfile.ZipFile(source) as archive:
                parts = {name: archive.read(name) for name in archive.namelist()}
            formula_parts = [name for name, data in parts.items() if b"<f>1+2</f>" in data]
            self.assertEqual(len(formula_parts), 1)
            xml = ET.fromstring(parts[formula_parts[0]])
            xml.find(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v").text = "3"
            parts[formula_parts[0]] = ET.tostring(xml)
            with zipfile.ZipFile(source, "w", zipfile.ZIP_DEFLATED) as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            plan = reference_plan()
            plan["journal"]["fields"]["Transaction Comments"] = "=1+2"
            run = folder / "run"
            events = Events(run)
            with patch("ylookup.sources.read_statement", return_value=parse_statement([(1, statement())])):
                result = execute_plan(plan, source, bank, run, {}, events)
            book = load_workbook(result["output"], data_only=True)
            self.assertEqual(book["Reference"]["A1"].value, 3)
            book.close()
            book = load_workbook(result["output"], data_only=False)
            self.assertEqual(book["Reference"]["A1"].value, "=1+2")
            self.assertEqual(book["Import"]["R2"].value, "=1+2")
            self.assertEqual(book["Import"]["R2"].data_type, "s")
            book.close()
            for destination in [source, Path(result["output"])]:
                with self.assertRaises(ValueError):
                    write_workbook(source, destination, result["evidence"], [])


if __name__ == "__main__":
    unittest.main()
