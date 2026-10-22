"""Where the example cannot decide, the description must: a script that takes the first candidate cannot."""
import copy
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from openpyxl import load_workbook

from ylookup.baseline import first_candidate_plan
from ylookup.builder import build_workflow, learn
from ylookup.events import Events
from ylookup.planner import run_workflow
from fixtures import FULL_ONLY, NEXT_MONTH, PAID_IN_FULL, PARTIAL_OK, FakeModel, fixture_plan, pages, paid_in_full


def full_only_plan():
    plan = fixture_plan()
    plan["steps"][0]["policy"] = "full_name_only"
    return plan


class ModelMattersTest(unittest.TestCase):
    def setUp(self):
        directory = TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.folder = Path(directory.name)
        self.source, self.target, self.bank, _ = paid_in_full(self.folder)

    def build(self, plan, name):
        notes = self.folder / f"{name}.txt"
        notes.write_text(PARTIAL_OK)
        with patch("ylookup.sources.pdf_pages", return_value=pages(PAID_IN_FULL)):
            result = build_workflow(self.source, self.target, [self.bank], notes, {}, FakeModel([plan]), Events(self.folder / name))
        self.assertEqual((result["status"], result["attempts"]), ("success", 1))
        return result["workflow"]

    def next_month(self, workflow, name, answers=None):
        with patch("ylookup.sources.pdf_pages", return_value=pages(NEXT_MONTH)):
            return run_workflow(workflow, self.source, [self.bank], self.folder / name, answers=answers)

    def taylor_amount(self, result):
        book = load_workbook(result["output"], data_only=True)
        return book["Import"]["H2"].value

    def test_code_alone_cannot_build_where_the_example_needs_reading(self):
        # Type and Account are each verified only as a lookup of the other; breaking that loop needs a judgement
        # (the line labels are template constants), which the first-candidate script cannot make in 5 attempts.
        with patch("ylookup.sources.pdf_pages", return_value=pages(PAID_IN_FULL)):
            profile = learn(self.source, self.target, [self.bank], {}, PARTIAL_OK)["profile"]
        notes = self.folder / "first.txt"
        notes.write_text(PARTIAL_OK)
        plan = first_candidate_plan(profile)
        with patch("ylookup.sources.pdf_pages", return_value=pages(PAID_IN_FULL)):
            result = build_workflow(self.source, self.target, [self.bank], notes, {},
                                    FakeModel([copy.deepcopy(plan) for _ in range(5)]), Events(self.folder / "first"))
        self.assertEqual(result["status"], "budget_exhausted")

    def test_the_example_cannot_decide_so_only_the_description_can(self):
        partial, full = self.build(fixture_plan(), "partial"), self.build(full_only_plan(), "full")  # both match it
        month = self.next_month(partial, "partial-run")
        self.assertEqual((month["status"], self.taylor_amount(month)), ("success", 100))
        month = self.next_month(full, "full-run")
        self.assertEqual([q["code"] for q in month["questions"]], ["link"])  # the cut-short reference is not trusted
        answered = self.next_month(full, "full-answered", {month["questions"][0]["id"]: "Taylor Reed"})
        self.assertEqual(self.taylor_amount(answered), 100)


if __name__ == "__main__":
    unittest.main()
