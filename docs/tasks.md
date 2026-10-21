# Tasks

✅ done · 🔨 next · ⏳ later.

## Done

| ID | Task | Evidence |
|---|---|---|
| ✅ B01 | Workbook inspection, package, CLI | `ylookup inspect` |
| ✅ S01 | Find tables by header meaning; parse the bank PDF; match payers | 10 investors, 2 receipts matched |
| ✅ W01 | Build loop: model plan → execute → compare → refine → save | 88 → 10 → 0 differences over three calls |
| ✅ W02 | Currency question (EUR vs USD); accepted answer saved in the workflow | `examples/exercise-answers.json` |
| ✅ R01 | `ylookup run`: saved workflow, no target or model | 20 rows, 0 differences from the target |
| ✅ R02 | Required DIU fields learned from the target and enforced at run time | Blank Deal / Domain now stop instead of passing |
| ✅ R03 | Report tab and column moves since build (`binding_changes`) | Renamed tabs: success, moves listed |
| ✅ U01 | Streamlit UI: Build / Run / Adapt / Run history | `ylookup ui` |
| ✅ D01 | README, design doc, sequence diagrams | [README](../README.md), [design](design.md) |
| ✅ C01 | Removed the legacy numeric experiment commands | 30 tests pass |
| ✅ R05 | Build reliability: retries with backoff, provider fallback, 5 calls, compact per-field feedback, field locking, type-preserving diffs | [design](design.md#known-limitations) |
| ✅ R06 | Measure: live rounds of 5 builds | 13 of 20 live builds since R05; 4 of 5 on the current prompt (v13) |
| ✅ A01 | `ylookup adapt`: model maps missing columns; code verifies and saves workflow v+1 | `combined.xlsx` → v2 |
| ✅ A02 | Stress variants and adapt tests | `scripts/stress_variants.py`, `tests/test_adapt.py` |

## Later

- [ ] ⏳ **A03** Inputs without cached formula results (recalculate, or ask).
- [ ] ⏳ **A04** General answer types beyond currency and column mapping; prompt on the CLI.

## Submission

- [x] ✅ **Q01** Clean-clone run of all three phases from the README: all results as documented.
- [x] ✅ **Q02** Credential and private-file scan of all history: clean.
