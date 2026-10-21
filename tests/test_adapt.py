import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from openpyxl import load_workbook

from ylookup.adapt import adapt_workflow
from ylookup.events import Events
from ylookup.planner import execute_plan, required_fields
from ylookup.sources import parse_statement
from ylookup.workbook import observed_labels
from test_builder import fixtures, reference_plan
from test_sources import statement


class FakeMapper:
    model = "test-model"

    def __init__(self, reply):
        self.reply, self.messages, self.schema = reply, [], None

    def complete(self, messages, events, attempt):
        self.messages.append(messages)
        return json.dumps(self.reply)


def saved_workflow(folder, source, bank):
    with patch("ylookup.sources.read_statement", return_value=parse_statement([(1, statement())])):
        result = execute_plan(reference_plan(), source, bank, folder / "build", {}, Events(folder / "build"))
    path = folder / "workflow.json"
    path.write_text(json.dumps({"schema_version": 2, "workflow_version": 1, "plan": reference_plan(),
                                "accepted_answers": {}, "bindings": result["evidence"]["bindings"],
                                "checks": {"required_fields": required_fields(json.loads((folder / "build/journal.json").read_text()))},
                                "headers": observed_labels(source, result["evidence"])}))
    return path


def rename(source, cells):
    book = load_workbook(source)
    for sheet, cell, value in cells:
        book[sheet][cell] = value
    book.save(source)
    book.close()


class AdaptTest(unittest.TestCase):
    def run_adapt(self, reply, renames, answers=None):
        directory = TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        folder = Path(directory.name)
        source, _, bank, context = fixtures(folder)
        workflow = saved_workflow(folder, source, bank)
        rename(source, renames)
        client = FakeMapper(reply)
        with patch("ylookup.sources.read_statement", return_value=parse_statement([(1, statement())])):
            result = adapt_workflow(workflow, source, bank, context, client, folder / "adapt", answers)
        return result, client

    def test_renamed_column_is_mapped_verified_and_saved_as_new_version(self):
        result, client = self.run_adapt({"mappings": {"tracker.received": "Paid Amount"}, "question": None},
                                        [("Roster", "E1", "Paid Amount")])
        self.assertEqual(result["status"], "success")
        self.assertTrue(result["adapted"])
        saved = json.loads(Path(result["workflow"]).read_text())
        self.assertEqual(saved["workflow_version"], 2)
        self.assertEqual(saved["headers"]["tracker.received"], "Paid Amount")
        self.assertEqual(saved["adaptations"][0]["mapped"]["tracker.received"]["from"], "Amount Received")
        request = json.loads(client.messages[0][1]["content"])
        self.assertEqual(request["missing_fields"][0]["field"], "tracker.received")
        self.assertNotIn("Taylor Reed", client.messages[0][1]["content"])
        self.assertTrue(Path(result["output"]).exists())

    def test_unsure_model_or_wrong_value_type_asks_instead_of_saving(self):
        renames = [("Roster", "E1", "Paid Amount"), ("Roster", "G1", "Notes"), ("Roster", "G3", "call back")]
        unsure, _ = self.run_adapt({"mappings": {"tracker.received": None}, "question": "Paid Amount or Notes?"}, renames)
        self.assertEqual(unsure["status"], "needs_input")
        self.assertIn("Notes", unsure["questions"][0]["candidates"])
        self.assertEqual(unsure["questions"][-1]["code"], "model_question")
        wrong, _ = self.run_adapt({"mappings": {"tracker.received": "Notes"}, "question": None}, renames)
        self.assertEqual(wrong["status"], "needs_input")
        self.assertEqual(wrong["questions"][0]["code"], "mapping_rejected")
        self.assertFalse(Path(wrong["run_dir"], "workflow.json").exists())

    def test_two_empty_columns_need_a_person_and_the_answer_is_saved(self):
        renames = [("Roster", "E1", "Amount 1"), ("Roster", "F1", "Amount 2")]
        guess = {"mappings": {"tracker.received": "Amount 1", "tracker.outstanding": "Amount 2"}, "question": None}
        asked, _ = self.run_adapt(guess, renames)
        self.assertEqual(asked["status"], "needs_input")
        self.assertEqual({q["field"] for q in asked["questions"] if q["code"] == "column_mapping"},
                         {"tracker.received", "tracker.outstanding"})
        answers = {"tracker.received": "Amount 1", "tracker.outstanding": "Amount 2"}
        answered, client = self.run_adapt({}, renames, answers)
        self.assertEqual(answered["status"], "success")
        self.assertEqual(client.messages, [])
        saved = json.loads(Path(answered["workflow"]).read_text())
        self.assertEqual(saved["adaptations"][0]["mapped"]["tracker.received"]["decided_by"], "person")

    def test_fitting_input_needs_no_model_and_missing_column_asks_without_model(self):
        fits, client = self.run_adapt({}, [])
        self.assertFalse(fits["adapted"])
        self.assertEqual(fits["status"], "success")
        self.assertEqual(client.messages, [])
        gone, client = self.run_adapt({}, [("Fund", "C1", None), ("Fund", "C2", None)])
        self.assertEqual(gone["status"], "needs_input")
        self.assertEqual(gone["questions"][0]["code"], "column_missing")
        self.assertEqual(client.messages, [])


if __name__ == "__main__":
    unittest.main()
