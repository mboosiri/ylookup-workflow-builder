import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from openpyxl import load_workbook

from ylookup.adapt import adapt_workflow
from ylookup.planner import run_workflow
from fixtures import FakeModel, PRIVATE, built, edit, parsed


def reply(mappings, question=None):
    return {"mappings": mappings, "question": question}


class AdaptTest(unittest.TestCase):
    def setUp(self):
        directory = TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.folder = Path(directory.name)
        self.workflow, self.source, self.bank, self.context = built(self.folder)
        self.round = 0

    def adapt(self, replies, answers=None, ask=None):
        self.round += 1
        client = FakeModel(replies)
        with patch("ylookup.planner.read_statement", return_value=parsed()):
            result = adapt_workflow(self.workflow, self.source, self.bank, self.context, client,
                                    self.folder / f"adapt-{self.round}", answers, ask)
        return result, client

    def test_renamed_fund_constants_map_one_to_one_and_v2_runs_directly(self):
        edit(self.source, ("Fund", "A1", "Fund Entity"), ("Fund", "B1", "Cash Account"), ("Fund", "C1", "Journal Batch Type"))
        result, client = self.adapt([reply({"Fund.Legal Entity": "Fund Entity", "Fund.Bank Account": "Cash Account",
                                            "Fund.Batch Type": "Journal Batch Type"})])
        self.assertEqual(result["status"], "success")
        saved = json.loads(Path(result["workflow"]).read_text())
        self.assertEqual(saved["workflow_version"], 2)
        self.assertEqual(saved["tables"]["Fund"]["columns"]["Bank Account"]["header"], "Cash Account")
        self.assertEqual(saved["adaptations"][0]["mapped"]["Fund.Legal Entity"],
                         {"from": "Legal Entity", "to": "Fund Entity", "decided_by": "model, verified by code"})
        request = client.messages[0][1]["content"]
        self.assertIn("feeds Import.Entity", request)
        for private in PRIVATE:
            self.assertNotIn(private, request)
        with patch("ylookup.planner.read_statement", return_value=parsed()):
            again = run_workflow(result["workflow"], self.source, self.bank, self.folder / "v2")
        self.assertEqual(again["status"], "success")

    def test_renamed_chart_headers_keep_the_right_table_and_a_renamed_tab_too_asks_which_table(self):
        edit(self.source, ("Accounts", "A1", "Ledger Code"), ("Accounts", "B1", "Movement"))
        result, _ = self.adapt([reply({"Accounts.Account Code": "Ledger Code", "Accounts.Transaction Type": "Movement"})])
        self.assertEqual(result["status"], "success")
        self.assertEqual(set(result["mapping"]), {"Accounts.Account Code", "Accounts.Transaction Type"})
        book = load_workbook(self.source)
        book["Accounts"].title = "Codes"
        book.save(self.source)
        asked, client = self.adapt([])
        self.assertEqual(asked["status"], "needs_input")
        self.assertEqual(client.messages, [])
        choice = asked["questions"][0]
        self.assertEqual((choice["code"], choice["options"]), ("table_choice", ["Codes!row 1: Ledger Code, Movement"]))
        result, _ = self.adapt([reply({"Accounts.Account Code": "Ledger Code", "Accounts.Transaction Type": "Movement"})],
                               {choice["id"]: choice["options"][0]})
        self.assertEqual(result["status"], "success")
        self.assertEqual(result["mapping"]["Accounts"]["decided_by"], "person")

    def test_a_swapped_model_pick_fails_the_learned_checks_and_nothing_is_saved(self):
        edit(self.source, ("Accounts", "A1", "Ledger Code"), ("Accounts", "B1", "Movement"))
        result, _ = self.adapt([reply({"Accounts.Account Code": "Movement", "Accounts.Transaction Type": "Ledger Code"})])
        self.assertNotEqual(result["status"], "success")
        self.assertFalse(list(self.folder.glob("adapt-*/workflow.json")))

    def test_two_empty_columns_the_workflow_writes_need_a_person(self):
        edit(self.source, ("Roster", "E1", "Amount 1"), ("Roster", "F1", "Amount 2"))
        guess = reply({"Roster.Amount Received": "Amount 1", "Roster.Amount outstanding": "Amount 2"})
        asked, _ = self.adapt([guess])
        self.assertEqual(asked["status"], "needs_input")
        self.assertTrue(all("a person must choose" in q["question"] for q in asked["questions"]))
        answers = {"column_mapping:Roster.Amount Received": "Amount 1", "column_mapping:Roster.Amount outstanding": "Amount 2"}
        result, client = self.adapt([], answers)
        self.assertEqual(result["status"], "success")
        self.assertEqual(client.messages, [])
        self.assertEqual(result["mapping"]["Roster.Amount Received"]["decided_by"], "person")

    def test_a_pick_with_the_wrong_value_type_is_rejected_and_a_person_finishes_interactively(self):
        edit(self.source, ("Roster", "C1", "Called"), ("Roster", "G1", "Notes"), ("Roster", "G3", "call back"))
        prompts = []

        def ask(items):
            prompts.extend(items)
            return {items[0]["id"]: {"answer": "Called"}}

        result, _ = self.adapt([reply({"Roster.Amount Called": "Notes"})], ask=ask)
        self.assertIn("rejected by code", prompts[0]["question"])
        self.assertEqual(prompts[0]["options"], ["Called", "Notes"])
        self.assertEqual(result["status"], "success")
        self.assertEqual(result["mapping"]["Roster.Amount Called"]["decided_by"], "person")

    def test_an_input_that_fits_needs_no_model_and_a_removed_column_stops(self):
        book = load_workbook(self.source)
        book["Roster"].title = "Partners"
        book.save(self.source)
        fits, client = self.adapt([])
        self.assertEqual((fits["status"], fits["adapted"], client.messages), ("success", False, []))
        self.assertIn("Roster: tab 'Roster' -> 'Partners'", fits["binding_changes"])
        edit(self.source, ("Fund", "C1", None), ("Fund", "C2", None))
        gone, client = self.adapt([])
        self.assertEqual(gone["status"], "stopped")
        self.assertIn("Fund.Batch Type", gone["reason"])


if __name__ == "__main__":
    unittest.main()
