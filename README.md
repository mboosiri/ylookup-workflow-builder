# Ylookup workflow builder

Give it an input workbook, a completed example of the same job and the customer's description. It learns
the workflow, saves it as `workflow.json`, and replays it on new inputs with no model.

Nothing in the code is about capital calls, bank statements or the `DIU` tab. Every file is read as
tables. Which file is which, how tables link, what each output column means, the steps and the checks
are all learned from the files and the description.

- **Model:** `qwen/qwen3-30b-a3b-instruct-2507` (30.5B, open weights) via OpenRouter. Why: [design decisions](docs/design.md#main-design-decisions).
- **Design, diagrams, assumptions, limitations:** [docs/design.md](docs/design.md)

## Submission checklist

| Item | Where |
|---|---|
| How to run the three steps | [Build](#phase-1-build), [Run](#phase-2-run), [Adapt](#phase-3-adapt-stress-test) |
| Model choice and design justification | [design.md: decisions](docs/design.md#main-design-decisions) |
| `.env.example` | [.env.example](.env.example) |
| The workflow the builder generated | [workflows/reconciliation-v3.json](workflows/reconciliation-v3.json) |
| The build that generated it | [workflows/reconciliation-v3.build/](workflows/reconciliation-v3.build/) |
| Sequence diagram: a new input through the saved workflow | [design.md](docs/design.md#how-a-saved-workflow-produces-the-output) |
| Assumptions | [design.md](docs/design.md#assumptions) |
| What I would improve with more time | [design.md](docs/design.md#with-more-time) |
| Main design decisions and alternatives | [design.md](docs/design.md#main-design-decisions) |

## Setup

Python 3.11+.

```sh
python3 -m venv .venv && source .venv/bin/activate
pip install -e ".[ui]"
cp .env.example .env        # then set OPENROUTER_API_KEY
python -m unittest discover -s tests
```

Put the supplied files in `data/` (git-ignored). The commands use the brief's file names, but any names
work: roles come from the content, and tables are found by their headers.

| File | Role |
|---|---|
| `data/capital-call-input.xlsx` | input |
| `data/capital-call-output.xlsx` | completed example (build only) |
| `data/bank-statement.pdf` | source |

The context is [examples/reconciliation-context.txt](examples/reconciliation-context.txt): the
customer's description from the brief, word for word.

## Phase 1: Build

```sh
ylookup build \
  --input   data/capital-call-input.xlsx \
  --target  data/capital-call-output.xlsx \
  --sources data/bank-statement.pdf \
  --context examples/reconciliation-context.txt
```

Before any model call it asks one question: the statement is in EUR and the tracker in USD. Answer `1`
(post as they are, no FX conversion). To run it unattended, add `--answers examples/answers.json --no-input`.

How a build works:

1. Code reads every table and compares the input with the example. It finds what the job produces:
   rows added to `DIU` and `Bank Statement Staging`, and three columns filled in the tracker.
2. Code finds links (the statement's `Description` names tracker investors, by full or partial name), the
   row counts that fit (10 investors × 2 lines = 20 DIU rows), and, for every output column, the
   expressions that reproduce the example.
3. The model reads this evidence and the description, and writes a plan from three tools: `link_rows`,
   `fill_columns` and `append_rows`. Or it asks a question.
4. Code runs the plan and compares every cell with the example. It sends the differences back and keeps
   the columns that already match. Up to 5 model calls.
5. `workflow.json` is saved only when the output matches the example and every learned check passes.

Names, the fund, bank references and dates are replaced by tokens before anything reaches the model.

The committed workflow took 3 calls: attempts 1 and 2 each had 40 cells wrong, then attempt 3 matched.
Its full trace (every call, plan, comparison and code step) is in
[workflows/reconciliation-v3.build/](workflows/reconciliation-v3.build/). Live builds: 3 of 3 on the
current prompt (1 to 3 calls), and 9 of 9 across three different descriptions on the previous one.

## Phase 2: Run

```sh
ylookup run \
  --workflow workflows/reconciliation-v3.json \
  --input    data/capital-call-input.xlsx \
  --sources  data/bank-statement.pdf
```

No target, no model, no memory of the build. Tables are found by their headers, so renamed tabs and moved
columns need no change. Output: `runs/run-<time>/output.xlsx`, with 0 differences from the completed example.

The result also lists what to follow up: Aaron Patel 5.00 and Daniel Brooks 20.00 still outstanding,
8 investors with no matching payment, and any statement line a person chose to leave unlinked.

If a computed number has a sign the example never showed (a negative outstanding after an overpayment, say),
the run asks before writing anything.

## Phase 3: Adapt (stress test)

```sh
python scripts/stress_variants.py data/capital-call-input.xlsx   # writes data/stress/*.xlsx

ylookup adapt \
  --workflow workflows/reconciliation-v3.json \
  --input    data/stress/constants-renamed.xlsx \
  --sources  data/bank-statement.pdf \
  --context  examples/reconciliation-context.txt
```

How adapt works:

1. Run the saved workflow. If it fits, nothing changes and any moved tabs or columns are listed.
2. For each header it can't find, the candidates are the unrecognised headers of the same table.
3. The model picks one for each, or none. It sees the old header, how the workflow uses it and the value
   types. Code then checks each pick: a real candidate, used once, same value types, and empty if the
   workflow writes it. If two empty columns fit, a person chooses.
4. Every check runs again before the workflow is saved as version 2, with who decided each mapping.

Results on the committed workflow ("= target" means 0 differences from the completed example):

| Variant | On its own | After a person answers |
|---|---|---|
| Tabs renamed, columns reordered | Fits as is, = target | |
| Extra investor | Fits as is, 22 DIU rows | |
| Renamed constants, chart headers, chart tab, tracker columns; all combined | Model maps them, v2, = target | |
| Tracker reshuffled, `Value Date` renamed | Mapped in 4 of 4 runs, v2, = target | |
| Distractor column next to the real one | Asks which (3 of 4 runs) or maps it | v2, = target |
| Two empty `Amount` columns | Asks which | v2, = target |
| Missing data | Asks for the missing values | = target |
| Duplicate name, unknown payer | Asks who paid, or "leave it unlinked" | Uses the answer |

## Asking, not guessing

Questions come from code (currency, an unmatched or doubtful payment, a missing value, a missing source
file, a renamed column) or from the model. Each has a stable id, and options when code can act on the answer.

- At a terminal, the command asks and carries on.
- With `--no-input`, it stops with exit code 2 and writes `questions.json`. Answer with `--answers file.json`.
- In the chat UI, questions appear as buttons.

Policy answers, like the currency one, are saved in `workflow.json`. Answers about one input's data stay
with that run.

## Testing

`python -m unittest discover -s tests` runs 58 offline tests. They use their own small workbooks, not the
brief's files.

Two scripts measure the builder live (about $0.002 per build):

- `scripts/stress_agentic.py` builds the workflow several times and adapts every stress variant. Each
  doubtful variant should ask, and each harmless one should not. Latest run: no guesses and no needless questions.
- `scripts/model_matters.py` builds each job twice, once by code alone (the first verified candidate
  everywhere) and once with the model:

| Job | Code alone | Model |
|---|---|---|
| Capital call (the brief's files) | builds | 3 of 3 |
| Test template where two columns each look up the other | fails | 3 of 3 |
| Same files, "partial names are fine" | fails | 3 of 3, and next month runs through a cut-short name |
| Same files, "full names only" | fails | 3 of 3, and next month asks about the cut-short name |
| Payroll: "tax is 20% of gross, net is gross minus tax" | builds | 3 of 3 |

On the brief's job, plain code can pick the right candidates. The model matters when the example can't
decide and only the description can, and when values need reading.

## Chat UI

```sh
ylookup ui        # http://localhost:8501, start it from the repo root
```

Attach the files, describe the job, say **build**. Questions appear as buttons and the trace streams in.
Then attach a new input and say **run**. **Try the sample job** adds the four sample files in one click.

## Outputs

Each command writes to `runs/<command>-<time>/`:

| File | Contents |
|---|---|
| `events.jsonl` | Every model call, plan, check, comparison and the stop reason |
| `evidence.json` | Build only: the masked evidence the model saw |
| `workflow.json` | Build or adapt, on success |
| `output.xlsx`, `outputs.json` | The filled workbook and its rows. Source files are never changed. |
| `questions.json`, `answers.json` | What needs answering, and the answers given |

Exit codes: `0` success, `2` stopped or needs input, `1` error.

## Code map

| File | Role |
|---|---|
| `cli.py`, `ui.py` | Command line and chat UI |
| `builder.py` | File roles, evidence, candidates, checks, the model loop, save |
| `planner.py` | Plan language, validation, execution, checks, `run_workflow` |
| `adapt.py` | Phase 3: model proposes mappings, code checks them, saves v2 |
| `sources.py` | Tables from workbooks and PDFs, finding tables by headers, name matching |
| `workbook.py` | Write outputs into a copy, read back, compare |
| `model.py` | OpenRouter client (JSON schema output, retries) |
| `baseline.py` | The code-alone plan, used only to measure the model |
