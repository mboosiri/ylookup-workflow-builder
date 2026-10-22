"""Where the model matters: the same jobs built by code alone (first candidate everywhere, resent 5 times so code's own
fallbacks still apply) and by the model, and each built workflow run on next month's input.

Jobs come from tests/fixtures.py, plus the real capital-call files when they are in data/. Needs OPENROUTER_API_KEY.
Usage: python scripts/model_matters.py [--repeats 3]
"""

import argparse
import copy
import json
from pathlib import Path
import shutil
import sys
from tempfile import mkdtemp
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tests"))
from fixtures import (FULL_ONLY, NEXT_MONTH, PAID_IN_FULL, PARTIAL_OK, FakeModel, fixtures, pages, paid_in_full,  # noqa: E402
                      payroll, statement)

from ylookup.baseline import first_candidate_plan  # noqa: E402
from ylookup.builder import PROMPT_VERSION, build_workflow, learn  # noqa: E402
from ylookup.events import Events, run_name  # noqa: E402
from ylookup.model import OpenRouter, settings  # noqa: E402
from ylookup.planner import run_workflow  # noqa: E402

DATA = Path("data")
pick = lambda *names: next((DATA / n for n in names if (DATA / n).exists()), None)


def job(name, make, statement=None, context=None, answers=None, month=None, expect=None):
    """make(folder) -> (input, target, sources, notes). month: next month's statement; expect(result) -> bool."""
    return {"name": name, "make": make, "statement": statement, "context": context, "answers": answers or {},
            "month": month, "expect": expect}


def with_context(make, context):
    def made(folder):
        source, target, sources, notes = make(folder)
        notes.write_text(context)
        return source, target, sources, notes
    return made


def build(spec, folder, client):
    source, target, sources, notes = spec["make"](folder)
    asked = []
    ask = lambda items: asked.extend(q["code"] for q in items) or {}
    with patch("ylookup.sources.pdf_pages", return_value=pages(spec["statement"])) if spec["statement"] else _nothing():
        result = build_workflow(source, target, sources, notes, dict(spec["answers"]), client, Events(folder / run_name("build")), ask=ask)
    return result, source, sources, asked


def first_candidate_client(spec, folder):
    source, target, sources, notes = spec["make"](folder)
    with patch("ylookup.sources.pdf_pages", return_value=pages(spec["statement"])) if spec["statement"] else _nothing():
        plan = first_candidate_plan(learn(source, target, sources, spec["answers"], notes.read_text())["profile"])
    return FakeModel([copy.deepcopy(plan) for _ in range(5)])


class _nothing:
    def __enter__(self): return self
    def __exit__(self, *exc): return False


def next_month(spec, result, source, sources, folder):
    if result["status"] != "success" or not spec["month"]:
        return None
    with patch("ylookup.sources.pdf_pages", return_value=pages(spec["month"])):
        return spec["expect"](run_workflow(result["workflow"], source, sources, folder / run_name("month")))


def jobs():
    found = [
        job("test template (Type and Account read each other)", _fixture, statement=statement()),
        job("paid in full, 'partial names are fine'", with_context(lambda f: _paid(f), PARTIAL_OK), statement=PAID_IN_FULL,
            month=NEXT_MONTH, expect=lambda r: r["status"] == "success"),
        job("paid in full, 'full names only'", with_context(lambda f: _paid(f), FULL_ONLY), statement=PAID_IN_FULL,
            month=NEXT_MONTH, expect=lambda r: [q["code"] for q in r.get("questions", [])] == ["link"]),
        job("payroll, rules only in the description", lambda f: _payroll(f)),
    ]
    real = [pick("capital-call-input.xlsx", "import-capital-call-notice-diu-input.xlsx"),
            pick("capital-call-output.xlsx", "import-capital-call-notice-diu-output.xlsx"),
            pick("bank-statement.pdf", "BS-Test-4.pdf")]
    if all(real):
        found.insert(0, job("capital call (the brief's files)", lambda f: (real[0], real[1], [real[2]], Path("examples/reconciliation-context.txt")),
                            answers=json.loads(Path("examples/answers.json").read_text())))
    return found


def _fixture(folder):
    source, target, bank, context = fixtures(folder)
    return source, target, [bank], context


def _paid(folder):
    source, target, bank, context = paid_in_full(folder)
    return source, target, [bank], context


def _payroll(folder):
    source, target, notes = payroll(folder)
    return source, target, [], notes


def main(repeats):
    key, model = settings()
    root = Path(mkdtemp(prefix="model-matters-"))
    rows = []
    for spec in jobs():
        folder = root / run_name("job")
        folder.mkdir()
        alone, source, sources, _ = build(spec, folder, first_candidate_client(spec, folder))
        row = {"job": spec["name"], "code alone": alone["status"], "code alone next month": next_month(spec, alone, source, sources, folder),
               "model": []}
        for _ in range(repeats):
            folder = root / run_name("job")
            folder.mkdir()
            result, source, sources, asked = build(spec, folder, OpenRouter(key, model, PROMPT_VERSION, 8000))
            row["model"].append({"status": result["status"], "calls": result.get("attempts"), "asked": asked,
                                 "next month as described": next_month(spec, result, source, sources, folder)})
        rows.append(row)
        print(json.dumps(row), flush=True)
    print("\n| Job | Code alone | Model (each build) | Next month as the description says: code alone / model |\n|---|---|---|---|")
    for row in rows:
        builds = ", ".join(f"{m['status']}" + (f" in {m['calls']}" if m["calls"] else "") + (" (asked)" if m["asked"] else "")
                           for m in row["model"])
        month = [m["next month as described"] for m in row["model"] if m["next month as described"] is not None]
        print(f"| {row['job']} | {row['code alone']} | {builds} | "
              + (f"{row['code alone next month']} / {sum(month)} of {len(month)}" if month else "-") + " |")
    shutil.rmtree(root, ignore_errors=True)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--repeats", type=int, default=3)
    main(parser.parse_args().repeats)
