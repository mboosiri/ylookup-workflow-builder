# Ylookup workflow builder

Learns a spreadsheet workflow from **input + completed output + customer context**,
saves it as `workflow.json`, then replays it on new inputs with no model and no memory.

Case: capital-call reconciliation. The bank statement (PDF) and the investor tracker (XLSX) go in;
the strict `DIU` upload tab comes out.

- **Model:** `qwen/qwen3-30b-a3b-instruct-2507` (30.5B, Apache-2.0) via OpenRouter.
- **Design, diagrams, assumptions and limitations:** [docs/design.md](docs/design.md)
- **Progress tracker:** [docs/tasks.md](docs/tasks.md)

## Setup

Python 3.11+.

```sh
python3 -m venv .venv && source .venv/bin/activate
pip install -e ".[ui]"
cp .env.example .env        # then set OPENROUTER_API_KEY
python -m unittest discover -s tests
```

Put the supplied files in `data/` (git-ignored):

| File | Purpose |
|---|---|
| `data/import-capital-call-notice-diu-input.xlsx` | input |
| `data/import-capital-call-notice-diu-output.xlsx` | completed target (Phase 1 only) |
| `data/BS-Test-4.pdf` | bank statement |

Context is [examples/reconciliation-context.txt](examples/reconciliation-context.txt): the customer's
description of the task, as plain text.

## Phase 1: Build

```sh
ylookup build \
  --input   data/import-capital-call-notice-diu-input.xlsx \
  --target  data/import-capital-call-notice-diu-output.xlsx \
  --bank    data/BS-Test-4.pdf \
  --context examples/reconciliation-context.txt \
  --answers examples/exercise-answers.json
```

**How a build works:**

1. The model proposes a plan.
2. Code executes it and compares the result with the target, field by field.
3. The differences go back to the model, compactly, per field. Fields that already match are locked.
   This repeats for at most 5 calls (`--max-attempts`).
4. `workflow.json` is saved only when the output matches the target.

Measured success rate: see [design.md](docs/design.md#known-limitations). A failed build saves nothing.

**The question it asks:** the statement is in EUR but the tracker is in USD.
- Run without `--answers` and the build stops before any model call, writing the question to `questions.json`.
- [examples/exercise-answers.json](examples/exercise-answers.json) is the accepted answer: keep the USD amounts, no FX conversion.
- In the UI you answer the question with a button instead.

The generated workflow for this case is committed at
[workflows/reconciliation-v2.json](workflows/reconciliation-v2.json).

## Phase 2: Run

```sh
ylookup run \
  --workflow workflows/reconciliation-v2.json \
  --input    data/import-capital-call-notice-diu-input.xlsx \
  --bank     data/BS-Test-4.pdf
```

The run uses no target, no build transcript and no model call. It writes `runs/run-<time>/output.xlsx`.
Verified result: 20 DIU rows and 0 differences from the completed target.

## Phase 3: Adapt (stress test)

```sh
python scripts/stress_variants.py data/import-capital-call-notice-diu-input.xlsx   # writes data/stress/*.xlsx

ylookup adapt \
  --workflow workflows/reconciliation-v2.json \
  --input    data/stress/combined.xlsx \
  --bank     data/BS-Test-4.pdf \
  --context  examples/reconciliation-context.txt
```

**How adapt works:**

1. Run the existing workflow. If it succeeds, nothing changes.
2. Otherwise find the known columns that are missing. The candidates are the unrecognised headers in the same table.
3. The model picks a candidate for each missing field, or `null`. It sees only the headers and value types, never the data.
4. Code checks each pick:
   - it must be a real candidate, used once, with the right value type;
   - a column the workflow writes to must be empty;
   - if two empty columns fit, a person must choose.
5. Every run check is re-run. Only then is `workflow.json` saved as version 2, with each mapping and who decided it.

| Variant in `data/stress/` | Result |
|---|---|
| `tabs-renamed`, `columns-reordered` | ✅ Fits as is. No model call; the moves are listed. |
| `columns-renamed`, `combined` | ✅ Model maps `Limited Partner`, `Paid Amount` and `Deal Name`; verified and saved as v2 |
| `ambiguous-amounts` | ⛔ Asks which of `Amount 1` / `Amount 2` is which. Answer with `--answers examples/adapt-answers-ambiguous.json` and it saves v2 with `decided_by: person`. |
| `missing-data` | ⛔ Stops at the first problem (blank investor name, tracker row 12) with a question. No workbook is written. |

## UI

```sh
ylookup ui        # opens http://localhost:8501
```

The UI has a tab for each phase plus **Run history**. Use it to see each build iteration and the
differences that remained, answer questions, preview the DIU rows and download `workflow.json` or `output.xlsx`.
It calls the same functions as the CLI.

## Outputs

Each command writes a new folder `runs/<command>-<time>/` (or `--run-dir`):

| File | Contents |
|---|---|
| `events.jsonl` | Every model call, candidate, check, comparison and the stop reason |
| `workflow.json` | Build or adapt, on success only |
| `output.xlsx` | The populated workbook. Source files are never modified. |
| `questions.json` | What needs answering, when the command stops |

Exit codes: `0` success · `2` stopped / needs input · `1` error.

## Code map

| File | Role |
|---|---|
| `cli.py`, `ui.py` | Entry points |
| `builder.py` | Build loop: evidence → model → execute → compare → refine |
| `adapt.py` | Phase 3: detect missing columns, model proposes a mapping, code verifies, save v+1 |
| `planner.py` | Plan schema, validation, execution, journal checks, `run_workflow` |
| `sources.py` | Find tables by meaning, parse the bank PDF, match payers |
| `workbook.py` | Write the output safely and compare it with the target |
| `model.py` | OpenRouter client (JSON-schema output, retries with backoff) |
