# Ylookup workflow builder

Learns a spreadsheet workflow from **input + completed output + customer context**, saves it as
`workflow.json`, then replays it on new inputs with no model and no memory.

The builder knows nothing about this case. Tables, output headers, column meanings, the steps and
the checks are all learned from the files:

- **Outputs:** code diffs the input against the completed example to find the rows it appends and the columns it fills.
- **Meanings:** code searches the input for every value of the example and lists the expressions that reproduce it (a column, a receipt field, a counter, a lookup, a concatenation, a difference).
- **Plan:** the model reads that evidence together with the customer's own description and composes the steps. Code runs them, compares the result with the example and sends the differences back.
- **Checks:** code mines the conventions every example row follows and enforces them on every later run.

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

The context is [examples/reconciliation-context.txt](examples/reconciliation-context.txt): the
customer's description from the brief, word for word, and nothing else.

## Questions

Any step can ask a question: a currency mismatch, a payment it cannot match, a blank value, a
renamed column, or something the model finds ambiguous. Each question has a stable id, and options
when code can act on the answer.

- **At a terminal**, the command asks: pick a numbered option or type an answer, then it carries on.
- **With `--no-input`** (or no terminal), it stops with exit code 2 and writes `questions.json`.
  Answer with `--answers file.json` holding `{"question id": "answer"}`.
- **In the UI**, every question shows its options or a text box.

Policy answers (e.g. the currency decision) are recorded in `workflow.json`, so later runs do not ask
again. Answers about one input's data (a blank cell, an unmatched payment) are kept in that run's
`answers.json`.

## Phase 1: Build

```sh
ylookup build \
  --input   data/import-capital-call-notice-diu-input.xlsx \
  --target  data/import-capital-call-notice-diu-output.xlsx \
  --bank    data/BS-Test-4.pdf \
  --context examples/reconciliation-context.txt
```

It asks one question before any model call: the statement is in EUR but the tracker is in USD.
Answer `1` (post the amounts as they are, without FX conversion). To run it unattended, add
`--answers examples/answers.json --no-input`.

**How a build works:**

1. Code finds every table by its header row and diffs the input against the target. Here it finds
   three outputs: rows appended to `DIU` and `Bank Statement Staging`, and three columns filled in the tracker.
2. Code works out which iterations fit each appended table (e.g. 10 tracker records × 2 lines = 20 DIU rows).
   It then lists the verified candidate expressions for every output column, per line where the lines differ.
3. The model composes the plan from those tools and the customer's description. Each attempt is one call.
4. Code runs the plan, compares every cell with the target and sends back the differences with the
   verified candidates. Columns that already match are locked. This repeats for at most 5 calls (`--max-attempts`).
5. `workflow.json` is saved only when the written workbook matches the target and every learned check passes.

Private values (names, the fund, bank references, dates) are replaced by tokens before anything reaches the model.

**Measured:** 20 of 20 live builds on the final code, each in 1–3 model calls and about $0.002.
16 succeeded on the model's choices alone. In 4, code adopted a verified candidate the model kept
missing; see [design.md](docs/design.md#known-limitations). A failed build saves nothing.

The generated workflow for this case is committed at [workflows/reconciliation-v3.json](workflows/reconciliation-v3.json).

## Phase 2: Run

```sh
ylookup run \
  --workflow workflows/reconciliation-v3.json \
  --input    data/import-capital-call-notice-diu-input.xlsx \
  --bank     data/BS-Test-4.pdf
```

The run uses no target, no build transcript and no model call. It finds each table by its header
fingerprint, so renamed tabs and moved columns need no change. It writes `runs/run-<time>/output.xlsx`.
Verified result: 20 DIU rows and 0 differences from the completed target.

## Phase 3: Adapt (stress test)

```sh
python scripts/stress_variants.py data/import-capital-call-notice-diu-input.xlsx   # writes data/stress/*.xlsx

ylookup adapt \
  --workflow workflows/reconciliation-v3.json \
  --input    data/stress/constants-renamed.xlsx \
  --bank     data/BS-Test-4.pdf \
  --context  examples/reconciliation-context.txt
```

**How adapt works:**

1. Run the existing workflow. If it succeeds, nothing changes and any moves are listed.
2. Every header the run cannot find becomes a candidate set: the unrecognised headers of the same table.
   A table is found by most of its build-time headers or its sheet name, so renamed headers never pull in the wrong table.
3. The model picks a candidate for each missing column, or `null`. It sees the old header, how the
   workflow uses the column, the value types under each candidate, and sample values for public reference tables only.
4. Code checks each pick:
   - it must be a real candidate, used once in its table, with the value types the column held;
   - a column the workflow writes must be empty;
   - if two empty columns fit, a person must choose.
5. Rejected picks become questions. Every learned check is re-run before `workflow.json` is saved as
   v+1, with the new headers and each mapping's `decided_by`.

Live results with the model, on the committed workflow. "= target" means the adapted run's output
has 0 differences from the completed target:

| Variant in `data/stress/` | Without a person | After answering |
|---|---|---|
| `tabs-renamed`, `columns-reordered` | ✅ Fits as is, no model call; moves listed; = target | |
| `extra-investor` (11th investor) | ✅ Fits as is; 22 DIU rows | |
| `constants-renamed` (Legal Entity, Bank Account, Batch Type) | ✅ Model maps all three; verified; v2; = target | |
| `chart-renamed`, `chart-tab-renamed` (chart headers, then its tab too) | ✅ Right table found; model maps both; v2; = target | |
| `columns-renamed`, `combined` | ✅ Model maps 3 columns; v2; = target | |
| `distractor-column` (`Call #8 - Total Call` next to `Call #8 - Estimate`) | ✅ v2, or ⛔ asks which (model unsure in 1 of 3 runs) | ✅ v2, `decided_by: person`; = target |
| `tracker-reshuffled` (columns swapped, `Value Date` renamed) | ⛔ usually asks whether `Value Date` is the statement date (mapped in 1 of 4 runs) | ✅ v2; = target |
| `ambiguous-amounts` (`Amount 1`, `Amount 2`) | ⛔ Two empty columns fit; a person must choose | ✅ v2, `decided_by: person`; = target |
| `missing-data` (blank investor name, blank Deal) | ⛔ Asks for both values; no workbook | ✅ = target |
| `duplicate-name`, `unknown-payer` | ⛔ Asks which investor paid: the candidates or "leave it unmatched" | ✅ Uses the chosen investor |

## UI

```sh
ylookup ui        # opens http://localhost:8501
```

The UI has a tab for each phase plus **Run history**. Use it to see each build iteration and the
differences that remained, answer questions, preview the output tables and download `workflow.json` or `output.xlsx`.
It calls the same functions as the CLI.

## Outputs

Each command writes a new folder `runs/<command>-<time>/` (or `--run-dir`):

| File | Contents |
|---|---|
| `events.jsonl` | Every model call, candidate, check, comparison and the stop reason |
| `evidence.json` | Build only: the masked evidence the model saw |
| `workflow.json` | Build or adapt, on success only |
| `output.xlsx`, `outputs.json` | The populated workbook and its output rows. Source files are never modified. |
| `questions.json` / `answers.json` | What needs answering when the command stops / answers given to a run |

Exit codes: `0` success · `2` stopped / needs input · `1` error.

## Code map

| File | Role |
|---|---|
| `cli.py`, `ui.py` | Entry points; the terminal prompt for questions |
| `builder.py` | Learn outputs, candidates and checks from the target; the model loop; save |
| `planner.py` | Expression language, plan validation, execution, checks, `run_workflow` |
| `adapt.py` | Phase 3: model proposes mappings for missing headers, code verifies, save v+1 |
| `sources.py` | Find tables by header rows, bind them by fingerprint, parse the bank PDF, match names, questions |
| `workbook.py` | Write outputs into a copy, read them back, compare |
| `model.py` | OpenRouter client (JSON-schema output, retries with backoff) |
