# Tasks

✅ done · 🔨 next · ⏳ later.

## Review round 1

| ID | Feedback | What changed | Evidence |
|---|---|---|---|
| ✅ V01 | Make the builder learn, not the code | Output tables, headers, filled columns, iterations, column meanings and checks are learned from input vs target. Steps are composed by the model from a tool catalogue. The DIU headers, alias lists, field vocabulary, three fixed steps and hand-written journal rules are gone. | `learn`, `survey`, `learn_checks`; tests use a different output template |
| ✅ V02 | Remove the answer hints | The system prompt describes only the tools and the evidence format. The context file is the brief's customer quote, word for word. | 10 of 10 live builds on the final prompt ([design](design.md#known-limitations)) |
| ✅ V03 | Adapt: renamed Bank Account / Legal Entity / Batch Type rejected | Each source column has one key, so mappings are one-to-one by construction. | `constants-renamed` → v2 (live) |
| ✅ V04 | Adapt: renamed chart headers pick the wrong table | Tables bind by their whole header fingerprint (or sheet name), best match first. | `chart-renamed`, `chart-tab-renamed` → v2 (live) |
| ✅ V05 | A wrong column mapping can hang the run | Cause: random-access reads in read-only mode past a missing `=SUM(` row, O(n²) over 16,406 rows. Sheets are now read in one pass, which stops after 50 blank rows, and rows are never deleted. | `test_a_wrong_column_mapping_fails_fast` |
| ✅ V06 | Harder adapt variants | 9 new variants: renamed constants, renamed chart (and its tab), reshuffled tracker, distractor column, duplicate name, unknown payer, extra investor | `scripts/stress_variants.py`, README table |
| ✅ V07 | General, interactive questions | Every question has a stable id and options when code can act. The terminal and UI prompt, `--answers` replays, and policy answers are recorded in `workflow.json`. | Interactive build and adapt via a pty; `test_a_model_question_is_asked_and_the_answer_is_used_and_recorded` |

## Review round 2

| ID | Feedback | What changed | Evidence |
|---|---|---|---|
| ✅ V08 | The agent should not expect a bank statement | No statement parser or receipt step: every file is read as tables (PDF header line + dated rows), and matching is a general `link_rows` between any text column and name column, chosen by the model from link evidence | `pdf_tables`, `link_rows`, `link_evidence`; 30/30 live builds |
| ✅ V09 | No fixed file roles | Roles from content (`guess_roles`); `--sources` takes any workbooks or PDFs | `test_file_roles_come_from_the_content_not_the_names` |
| ✅ V10 | UI: file upload, chat, workflow trace | One chat page: attach files, describe the job, answer questions inline (buttons or text), live trace, build → run → adapt in one conversation | Clicked through live: build with two inline questions, then adapt of `constants-renamed` to v2 |

## Done earlier

| ID | Task |
|---|---|
| ✅ B01 | Workbook inspection, package, CLI |
| ✅ W01 | Build loop: model plan → execute → compare → refine → save |
| ✅ R01 | `ylookup run`: saved workflow, no target or model |
| ✅ R03 | Report tab and column moves since build |
| ✅ U01 | Streamlit UI: Build / Run / Adapt / Run history |
| ✅ A01 | `ylookup adapt`: model maps missing columns; code verifies and saves workflow v+1 |

## Later

- [ ] ⏳ **A03** Inputs without cached formula results (recalculate instead of asking for each value).
- [ ] ⏳ **M01** Benchmark of several workflows and variants, scored per prompt or code change.
- [ ] ⏳ **M02** Several completed examples per build, to separate rules from coincidences in learned checks.

## Testing handover — 22 October 2026

- Direct chat attachments verified with a saved workflow, XLSX input and PDF source: 20 DIU rows, 0 differences from the completed target.
- Fixed chat runs dropping XLSX sources when PDFs were present, or excluding build-time XLSX sources; regression covers both combinations.
- Sidebar selection and Add are submitted together in a form; verified one-click addition in the browser.
- Each chat upload submission gets its own folder, preserving earlier files when a later input has the same name.
- Live model build/adapt measurements from the previous session have not been repeated in this handover. OCR, formula recalculation, broader benchmarks and multiple training examples remain later work.
- Chat routing: commands are short messages matched as whole words; the pasted customer description is kept even when it says "build" or "truncated"; out-of-scope requests (images, videos…) are declined and not stored as the description. `tests/test_ui.py`.
- PDF amounts not written as `1,234.56` (symbols, CR/DR, brackets, `1.234,56`) now ask a question instead of joining the description text. `test_amounts_in_another_form_ask_instead_of_joining_the_text`; `BS-Test-4.pdf` still reads 2 rows with no questions.
- Stress for iteration and inquiry: `scripts/stress_agentic.py` scores a live build and every stress variant (asked / guessed / over-asked); 5 offline tests pin the loop and the questions. 1-repeat result: build iterated (2 calls), every doubtful variant asked, no harmless variant asked. "Leave it unlinked" was then stopped by the learned checks; fixed: those rows are excused from link-reading checks and linked totals (`test_an_unclear_payer_is_asked_about_never_guessed`), rerun: both variants finish.
- Missing source: when the completed example holds varying values found in no attached file (e.g. the statement left out), the builder names the tables and columns and asks to attach the source or build anyway, before any model call. Verified in the chat: stop, attach the PDF, build again → only the currency question. Also: the model's questions now show real values instead of `V_n` tokens, and a lone workbook or two unrelated workbooks are described as such instead of as "the input" / "a source".
- Run without the source file: a saved workflow now says which separate file a table came from at build time and asks to attach it, instead of offering an unrelated table. Tables record their file; older workflows fall back to `built_from`.
- Matching rule: after a failed attempt the model is told how many rows its link rule linked and how many another rule links (evidence only). A model question that this evidence may answer goes back to the model once before a person; made-up tokens are dropped from options. Live: one-line description 3 of 3 (was 2 of 3); brief 3 of 3; all stress variants as expected.
- Link note names its knock-on: the differing columns that read across a failed link ("fix the link_rows step, not their expressions"). Prompt v22: the model asks people only about the business (data, rules, policies), never how to write the plan. Live, 3 repeats per description: brief 3/3, one-line 3/3, fund manager's 3/3; no person asked; stress variants unchanged.
- Payments the example never showed: each numeric output column learns the sign it kept in the example; a run value outside it (an overpayment or double payment giving a negative outstanding) asks "post it as computed" or "stop". Runs also list what to follow up: records whose own value minus their linked total is not zero (Aaron Patel 5.00, Daniel Brooks 20.00, 8 with nothing linked). Live: builds 3/3 (brief) and 3/3 (one-line); a fresh workflow learned 17 sign checks and raised no question on the real input or harmless variants.
