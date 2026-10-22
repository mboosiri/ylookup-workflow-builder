# Ylookup workflow builder

Learns a spreadsheet workflow from **input + completed output + customer context**, saves it as
`workflow.json`, then replays it on new inputs with no model and no memory.

The builder knows nothing about this case, and expects no particular file. Every file is read as
tables: workbook sheets by their header rows, PDFs by a header line followed by dated rows. Which
file is the input, the completed example or a source, which table links to which, the output
headers, column meanings, the steps and the checks are all learned:

- **Roles:** the completed example is the workbook that adds rows or fills columns compared with another; that one is the input; every other file is a source.
- **Outputs:** code diffs the input against the completed example to find the rows it appends and the columns it fills.
- **Links:** code finds text columns whose values name the records of another table (here, statement descriptions name investors, by full or partial name).
- **Meanings:** code searches all tables for every value of the example and lists the expressions that reproduce it (a column, a linked record's column, a sum or value over linked rows, a counter, a lookup, a concatenation, a difference).
- **Plan:** the model reads that evidence together with the customer's own description and composes the steps. Code runs them, compares the result with the example and sends the differences back.
- **Checks:** code mines the conventions every example row follows and enforces them on every later run.

Case: capital-call reconciliation. The investor tracker (XLSX) and a bank statement (PDF) go in;
the strict `DIU` upload tab comes out. Nothing in the code refers to either.

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
| `data/BS-Test-4.pdf` | a source: the bank statement (`--sources` takes any workbooks or PDFs) |

The context is [examples/reconciliation-context.txt](examples/reconciliation-context.txt): the
customer's description from the brief, word for word, and nothing else.

## Questions

Any step can ask a question: a currency mismatch, a payment it cannot match, a blank value, a
renamed column, or something the model finds ambiguous. Each question has a stable id, and options
when code can act on the answer.

- **At a terminal**, the command asks: pick a numbered option or type an answer, then it carries on.
- **With `--no-input`** (or no terminal), it stops with exit code 2 and writes `questions.json`.
  Answer with `--answers file.json` holding `{"question id": "answer"}`.
- **In the chat UI**, a question arrives as a message with option buttons; click one or type any answer.

Policy answers (e.g. the currency decision) are recorded in `workflow.json`, so later runs do not ask
again. Answers about one input's data (a blank cell, an unmatched payment) are kept in that run's
`answers.json`.

## Phase 1: Build

```sh
ylookup build \
  --input   data/import-capital-call-notice-diu-input.xlsx \
  --target  data/import-capital-call-notice-diu-output.xlsx \
  --sources data/BS-Test-4.pdf \
  --context examples/reconciliation-context.txt
```

It asks one question before any model call: the statement is in EUR but the tracker is in USD.
Answer `1` (post the amounts as they are, without FX conversion). To run it unattended, add
`--answers examples/answers.json --no-input`.

**How a build works:**

1. Code reads every table of every file and diffs the input against the target. Here it finds
   three outputs: rows appended to `DIU` and `Bank Statement Staging`, and three columns filled in the tracker.
   It finds the statement's own columns (`Date`, `Description`, `Withdrawal`, `Deposit`, `Balance`) and
   that its `Description` names `Capital Call Tracker.Investor` (2 of 2 rows, by partial name).
2. Code works out which iterations fit each appended table (e.g. 10 tracker records × 2 lines = 20 DIU rows).
   It then lists the verified candidate expressions for every output column, per line where the lines differ
   (e.g. `Amount (Local)` = `{"sum": "BS-Test-4.Deposit"}` over the rows linked to each investor).
3. The model composes the plan from three tools (`link_rows`, `fill_columns`, `append_rows`) and the
   customer's description, or asks a question. Each attempt is one call.
4. Code runs the plan, compares every cell with the target and sends back the differences with the
   verified candidates. Columns that already match are locked. This repeats for at most 5 calls (`--max-attempts`).
5. `workflow.json` is saved only when the written workbook matches the target and every learned check passes.

Private values (names, the fund, bank references, dates) are replaced by tokens before anything reaches the model.

**Measured:** 30 of 30 unattended live builds on the generic code, each in 1–2 model calls (about
$0.002), none needing code to adopt a candidate; see [design.md](docs/design.md#known-limitations).
A failed build saves nothing.

The generated workflow for this case is committed at [workflows/reconciliation-v3.json](workflows/reconciliation-v3.json).

## Phase 2: Run

```sh
ylookup run \
  --workflow workflows/reconciliation-v3.json \
  --input    data/import-capital-call-notice-diu-input.xlsx \
  --sources  data/BS-Test-4.pdf
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
  --sources  data/BS-Test-4.pdf \
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
| `duplicate-name`, `unknown-payer` | ⛔ Asks which investor a statement row belongs to: the candidates or "leave it unlinked" | ✅ Uses the chosen investor |

## Stress: iteration and inquiry

Two behaviours, scored live on the given files. Expectations are per kind of variation, never per value:
a doubtful variant should ask (else it **guessed**), a harmless one should not (else it **over-asked**).

```sh
python scripts/stress_variants.py data/import-capital-call-notice-diu-input.xlsx
python scripts/stress_agentic.py --repeats 3      # writes runs/stress-agentic-<time>/scores.json
```

Measured (1 repeat, 22 October 2026):

| Scenario | Expect | Result | Model calls |
|---|---|---|---|
| build | iterate | ✅ attempt 1 differed (20 cells), attempt 2 matched | 2 |
| `duplicate-name`, `unknown-payer` | ask | ✅ asked (`link`), no model call; after "leave it unlinked" the run finishes | 0 |
| `missing-data` | ask | ✅ asked (4 `missing_value`) | 0 |
| `ambiguous-amounts` | ask | ✅ asked (`column_mapping`) | 1 |
| `distractor-column` | ask or map | ✅ asked | 1 |
| `tracker-reshuffled` | ask or map | ✅ mapped, verified | 1 |
| `tabs-renamed`, `columns-reordered`, `extra-investor` | no question | ✅ quiet | 0 |

None guessed, none over-asked. A row a person leaves unlinked (the customer's "leave it blank to chase up
later") is staged with a blank investor and credited to nobody; the learned checks excuse its link columns
and keep it out of linked totals, and still check everything else.

Offline, `tests/test_builder.py` and `tests/test_run.py` pin the same behaviours on a different template:
the loop stops at the first match, an exhausted budget saves nothing, each attempt is told only its own
differences, a clean job asks nothing, and an unclear payer is asked about, never guessed, and can be left unlinked.

## UI: chat

```sh
ylookup ui        # opens http://localhost:8501
```

One chat page. Attach files to a message (or add files already in `data/` from the sidebar) and talk:

1. **Attach the job's files** in any order. The assistant says which one is the input, the completed
   example and a source, from their content.
2. **Describe the job in your own words.** Once there is enough description it starts learning; you can
   also say **build**. The trace streams into the chat: what it learned, each plan, what still
   differed, and where code stepped in. Questions, whether from code (EUR vs USD) or from the
   model, arrive as messages with option buttons, or you type an answer. Answers are recorded in the workflow.
3. **Attach a new or changed input and say run.** It runs the workflow with no model. If headers
   changed, it adapts the workflow and shows each remapped column and who decided it.

Each result has its downloads (`workflow.json`, `output.xlsx`) and the full trace. The UI calls the same functions as the CLI.

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
| `cli.py`, `ui.py` | Entry points: the terminal prompt for questions; the chat UI |
| `builder.py` | File roles; learn outputs, links, candidates and checks from the target; the model loop; save |
| `planner.py` | Expression language, plan validation, execution, checks, `run_workflow` |
| `adapt.py` | Phase 3: model proposes mappings for missing headers, code verifies, save v+1 |
| `sources.py` | Read tables from workbooks and PDFs, bind them by fingerprint, link rows by name, questions |
| `workbook.py` | Write outputs into a copy, read them back, compare |
| `model.py` | OpenRouter client (JSON-schema output, retries with backoff) |
