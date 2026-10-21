import copy
import hashlib
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from openpyxl import load_workbook

from ylookup.builder import build_workflow, learn, reply_schema
from ylookup.events import Events
from ylookup.planner import run_workflow, validate_plan
from fixtures import PRIVATE, FakeModel, fixture_plan, fixtures, parsed, statement


EUR = "currency_conflict:EUR vs USD"


class LearnTest(unittest.TestCase):
    def test_outputs_candidates_and_checks_come_from_the_target(self):
        with TemporaryDirectory() as directory:
            source, target, bank, _ = fixtures(Path(directory))
            with patch("ylookup.planner.read_statement", return_value=parsed()):
                learned = learn(source, target, bank, {})
        outputs = learned["spec"]["outputs"]
        self.assertEqual({k: v["op"] for k, v in outputs.items()},
                         {"Staging": "append_rows", "Roster": "fill_columns", "Import": "append_rows"})
        self.assertEqual(outputs["Roster"]["columns"], ["Receipt Date", "Amount Received", "Amount outstanding"])
        self.assertEqual(learned["shapes"]["Import"], ("Roster", False, 2))
        view = {o["table"]: o for o in learned["profile"]["outputs"]}
        columns = view["Import"]["columns"]
        self.assertIn({"lookup": {"table": "Accounts", "where": "Transaction Type", "equals": {"out": "Type"}, "return": "Account Code"}},
                      columns["Account"]["all lines"])
        self.assertEqual(columns["Comment"]["all lines"][0], {"concat": ["Paid by ", {"col": "Roster.Investor"}]})
        self.assertEqual(columns["Partner"], {"line 1": [None], "line 2": [{"col": "Roster.Investor"}]})
        self.assertEqual(view["Roster"]["columns"]["Amount outstanding"]["all rows"],
                         [{"sub": [{"col": "Roster.Amount Called"}, {"var": "receipt.amount"}]}])
        kinds = {(c["column"], c["kind"]) for c in learned["checks"] if c["table"] == "Import"}
        for expected in [("Entry", "entry_counter"), ("Line", "line_counter"), ("Ref", "sequence"), ("Debit", "constant"),
                         ("Amount", "same_in_entry"), ("Amount", "total"), ("Account", "expression"), ("Partner", "blank")]:
            self.assertIn(expected, kinds)
        self.assertNotIn(("Entity", "constant"), kinds)  # an input value, so only uniformity is learned
        sent = json.dumps(learned["mask"](learned["profile"]))
        for private in PRIVATE:
            self.assertNotIn(private, sent)
        json.dumps(reply_schema(learned["schema_view"]))


class BuildTest(unittest.TestCase):
    def build(self, replies, statement_pages=None, answers=None, ask=None):
        directory = TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        folder = Path(directory.name)
        source, target, bank, context = fixtures(folder)
        client = FakeModel(replies)
        events = Events(folder / "run")
        with patch("ylookup.planner.read_statement", return_value=parsed(*(statement_pages or []))):
            result = build_workflow(source, target, bank, context, answers or {}, client, events, ask=ask)
        return result, client, events, (source, bank)

    def test_differences_go_back_matching_columns_are_locked_and_the_workflow_runs(self):
        wrong = fixture_plan()
        wrong["steps"][3]["columns"]["Debit"] = {"line": ["N", "Y"]}
        regressed = fixture_plan()
        regressed["steps"][3]["columns"]["Comment"] = "Paid"
        result, client, events, (source, bank) = self.build([wrong, regressed])
        self.assertEqual(result["status"], "success")
        self.assertEqual(result["attempts"], 2)
        feedback = json.loads(client.messages[1][-1]["content"])
        self.assertEqual([d["column"] for d in feedback["differences"]], ["Debit", "Debit"])
        self.assertIn("Import.Comment", feedback["locked_columns"])
        saved = json.loads(Path(result["workflow"]).read_text())
        self.assertEqual(saved["plan"]["steps"][3]["columns"]["Comment"], fixture_plan()["steps"][3]["columns"]["Comment"])
        self.assertEqual(set(saved["tables"]["Accounts"]["columns"]), {"Account Code", "Transaction Type"})
        for private in PRIVATE:
            self.assertNotIn(private, json.dumps(client.messages))
        with patch("ylookup.planner.read_statement", return_value=parsed()):
            run = run_workflow(result["workflow"], source, bank, events.directory / "replay")
        self.assertEqual(run["status"], "success")
        book = load_workbook(run["output"], data_only=True)
        self.assertEqual([book["Roster"]["F3"].value, book["Roster"]["F4"].value, book["Roster"]["E5"].value], [5, 50, 95])
        self.assertEqual(book["Import"]["F3"].value, "RECEIVABLE-CODE")
        book.close()

    def test_a_model_question_is_asked_and_the_answer_is_used_and_recorded(self):
        question = {"status": "needs_input", "workflow": None, "question": "Should non-payers get zero entries?",
                    "options": ["Yes, keep them", "No, leave them out"]}
        asked = []

        def ask(items):
            asked.extend(items)
            return {items[0]["id"]: {"answer": "Yes, keep them"}}

        result, client, _, _ = self.build([question, fixture_plan()], ask=ask)
        self.assertEqual(result["status"], "success")
        self.assertEqual(asked[0]["options"], ["Yes, keep them", "No, leave them out"])
        self.assertIn("Yes, keep them", client.messages[1][-1]["content"])
        recorded = json.loads(Path(result["workflow"]).read_text())["answers"]
        self.assertEqual(recorded[0]["question"], question["question"])
        self.assertEqual(recorded[0]["answer"], "Yes, keep them")

    def test_without_a_person_questions_stop_with_ids_an_answers_file_can_settle(self):
        blocked, client, events, _ = self.build([], [statement(currency="EUR")])
        self.assertEqual(blocked["status"], "needs_input")
        self.assertEqual(client.messages, [])
        self.assertEqual(json.loads((events.directory / "questions.json").read_text())[0]["id"], EUR)
        answer = {EUR: "Post the amounts as they are, without FX conversion"}
        result, client, events, (source, bank) = self.build([fixture_plan()], [statement(currency="EUR")], answer)
        self.assertEqual(result["status"], "success")
        self.assertIn("without FX conversion", client.messages[0][1]["content"])
        saved = json.loads(Path(result["workflow"]).read_text())
        self.assertEqual([a["id"] for a in saved["answers"]], [EUR])
        with patch("ylookup.planner.read_statement", return_value=parsed(statement(currency="EUR"))):
            self.assertEqual(run_workflow(result["workflow"], source, bank, events.directory / "eur")["status"], "success")
        with patch("ylookup.planner.read_statement", return_value=parsed(statement(currency="GBP"))):
            other = run_workflow(result["workflow"], source, bank, events.directory / "gbp")
        self.assertEqual([q["code"] for q in other["questions"]], ["currency_conflict"])
        stop = {EUR: "Stop: the amounts need converting first"}
        with patch("ylookup.planner.read_statement", return_value=parsed(statement(currency="EUR"))):
            halted = run_workflow(result["workflow"], source, bank, events.directory / "halt", answers=stop)
        self.assertEqual(halted["status"], "stopped")

    def test_invalid_plans_get_every_problem_back(self):
        broken = copy.deepcopy(fixture_plan())
        broken["steps"][3]["columns"]["Amount"] = {"var": "receipt.total"}
        broken["steps"][3]["columns"]["Comment"] = "Paid by V_1"
        broken["steps"][3]["lines"] = 3
        del broken["steps"][1]
        result, client, _, _ = self.build([broken, fixture_plan()])
        self.assertEqual(result["status"], "success")
        error = json.loads(client.messages[1][-1]["content"])["validation_error"]
        for detail in ["receipt.total", "anonymised token", "line needs exactly 3", "no step produces Staging"]:
            self.assertIn(detail, error)
        with self.assertRaises(ValueError):
            validate_plan({"intent": "x", "steps": [], "checks": []}, {"tables": {}, "outputs": {}})

    def test_common_small_model_slips_are_repaired_not_rejected(self):
        sloppy = fixture_plan()
        columns = sloppy["steps"][3]["columns"]
        columns["Entity"] = "{col: Fund.Legal Entity}"
        columns["Entry"] = {"line": ["{var: entry}"]}
        sloppy["steps"][1]["columns"]["Deposit"] = {"line": [{"var": "receipt.amount"}, {"var": "receipt.amount"}]}
        result, _, _, _ = self.build([sloppy])
        self.assertEqual((result["status"], result["attempts"]), ("success", 1))
        saved = json.loads(Path(result["workflow"]).read_text())["plan"]["steps"]
        self.assertEqual(saved[3]["columns"]["Entity"], {"col": "Fund.Legal Entity"})
        self.assertEqual(saved[1]["columns"]["Deposit"], {"var": "receipt.amount"})

    def test_a_repeated_failure_takes_the_only_verified_candidate_and_says_so(self):
        stubborn = fixture_plan()
        stubborn["steps"][3]["columns"]["Line"] = {"var": "entry"}
        result, _, events, _ = self.build([stubborn, copy.deepcopy(stubborn)])
        self.assertEqual((result["status"], result["attempts"]), ("success", 2))
        trace = [json.loads(line) for line in events.path.read_text().splitlines()]
        self.assertIn(["Import.Line"], [e["columns"] for e in trace if e["event"] == "adopted_candidates"])
        payers_only = fixture_plan()
        payers_only["steps"][3]["only_received"] = True
        result, client, events, _ = self.build([payers_only, copy.deepcopy(payers_only)])
        self.assertEqual((result["status"], result["attempts"]), ("success", 2))
        feedback = json.loads(client.messages[1][-1]["content"])["differences"][0]
        self.assertEqual(feedback["iterations_that_fit"], [{"for_each": "Roster", "only_received": False, "lines": 2}])
        self.assertIn("Import iteration", events.path.read_text())

    def test_a_target_identical_to_the_input_teaches_nothing(self):
        with TemporaryDirectory() as directory:
            source, _, bank, _ = fixtures(Path(directory))
            with patch("ylookup.planner.read_statement", return_value=parsed()):
                with self.assertRaisesRegex(ValueError, "nothing to learn"):
                    learn(source, source, bank, {})
            self.assertTrue(hashlib.sha256(Path(source).read_bytes()).digest())


if __name__ == "__main__":
    unittest.main()
