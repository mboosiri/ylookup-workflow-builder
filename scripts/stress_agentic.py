"""Live scoreboard for two behaviours on the given files: the build iterates, and doubt becomes a question.

Builds the workflow N times (calls, iterated or not), then adapts the committed workflow on each
stress variant N times and records whether it asked. Expectations are about the kind of variation,
never about values. Needs OPENROUTER_API_KEY; costs about $0.002 per build.

Usage: python scripts/stress_variants.py data/import-capital-call-notice-diu-input.xlsx
       python scripts/stress_agentic.py [--repeats 3]
"""

import argparse
import json
from pathlib import Path
import sys
import time

from ylookup.adapt import PROMPT_VERSION as ADAPT_PROMPT_VERSION, adapt_workflow
from ylookup.builder import PROMPT_VERSION, build_workflow
from ylookup.events import Events
from ylookup.model import OpenRouter, settings
from ylookup.sources import LEAVE

DATA, STRESS = Path("data"), Path("data/stress")
INPUT, TARGET = DATA / "import-capital-call-notice-diu-input.xlsx", DATA / "import-capital-call-notice-diu-output.xlsx"
SOURCES = [DATA / "BS-Test-4.pdf"]
CONTEXT, ANSWERS = Path("examples/reconciliation-context.txt"), Path("examples/answers.json")
WORKFLOW = Path("workflows/reconciliation-v3.json")
# What a careful builder does with each kind of variation: "ask", "either" (ask or map correctly), "quiet" (no question).
EXPECT = {"duplicate-name": "ask", "unknown-payer": "ask", "missing-data": "ask", "ambiguous-amounts": "ask",
          "distractor-column": "either", "tracker-reshuffled": "either",
          "tabs-renamed": "quiet", "columns-reordered": "quiet", "extra-investor": "quiet"}


def recorder(asked):
    """An ask() that records every question and answers like a cautious person: leave unclear rows unlinked,
    else take the first option; free-text questions are left unanswered (counted as asked)."""
    def ask(items):
        asked.extend(q["code"] for q in items)
        return {q["id"]: (LEAVE if LEAVE in (q.get("options") or []) else q["options"][0])
                for q in items if q.get("options")}
    return ask


def trace(directory):
    return [json.loads(line) for path in Path(directory).rglob("events.jsonl") for line in path.read_text().splitlines()]


def build_once(key, model, directory):
    asked = []
    events = Events(directory, (key,))
    result = build_workflow(INPUT, TARGET, SOURCES, CONTEXT, json.loads(ANSWERS.read_text()),
                            OpenRouter(key, model, PROMPT_VERSION, 8000), events, ask=recorder(asked))
    comparisons = [e["status"] for e in trace(directory) if e["event"] == "comparison"]
    return {"status": result["status"], "calls": sum(e["event"] == "model_request" for e in trace(directory)),
            "iterated": result["status"] == "success" and comparisons[:1] == ["fail"], "asked": asked}


def adapt_once(key, model, variant, directory):
    asked = []
    result = adapt_workflow(WORKFLOW, variant, SOURCES, CONTEXT, OpenRouter(key, model, ADAPT_PROMPT_VERSION, 1500),
                            directory, ask=recorder(asked))
    asked += [q["code"] for q in result.get("questions", [])]
    return {"status": result["status"], "calls": sum(e["event"] == "model_request" for e in trace(directory)), "asked": asked}


def verdict(expect, run):
    if expect == "ask":
        return "asked" if run["asked"] else "guessed"
    if expect == "quiet":
        return "over-asked" if run["asked"] else "quiet"
    return "asked" if run["asked"] else ("mapped" if run["status"] == "success" else run["status"])


def main(repeats):
    try:
        key, model = settings()
    except ValueError as error:
        sys.exit(f"{error} The scoreboard calls the model; set OPENROUTER_API_KEY in .env.")
    missing = [p for p in [INPUT, TARGET, *SOURCES] + [STRESS / f"{v}.xlsx" for v in EXPECT] if not p.exists()]
    if missing:
        sys.exit("Missing files (run scripts/stress_variants.py first): " + ", ".join(map(str, missing)))
    folder = Path("runs") / time.strftime("stress-agentic-%Y%m%d-%H%M%S")
    rows = []
    for n in range(1, repeats + 1):
        run = build_once(key, model, folder / f"build-{n}")
        rows.append({"scenario": "build", "run": n, **run, "verdict": "iterated" if run["iterated"] else run["status"]})
        print(rows[-1], flush=True)
    for variant, expect in EXPECT.items():
        for n in range(1, repeats + 1):
            run = adapt_once(key, model, STRESS / f"{variant}.xlsx", folder / f"{variant}-{n}")
            rows.append({"scenario": variant, "expect": expect, "run": n, **run, "verdict": verdict(expect, run)})
            print(rows[-1], flush=True)
    (folder / "scores.json").write_text(json.dumps(rows, indent=2))
    print(f"\n| Scenario | Expect | Runs | Verdicts | Avg calls |\n|---|---|---|---|---|")
    for scenario in dict.fromkeys(r["scenario"] for r in rows):
        mine = [r for r in rows if r["scenario"] == scenario]
        verdicts = ", ".join(f"{v} {sum(r['verdict'] == v for r in mine)}" for v in dict.fromkeys(r["verdict"] for r in mine))
        print(f"| `{scenario}` | {mine[0].get('expect', 'iterate')} | {len(mine)} | {verdicts} | "
              f"{sum(r['calls'] for r in mine) / len(mine):.1f} |")
    print(f"\nWritten {folder / 'scores.json'}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--repeats", type=int, default=3)
    main(parser.parse_args().repeats)
