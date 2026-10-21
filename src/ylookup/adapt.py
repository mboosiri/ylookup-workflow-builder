"""Phase 3: update a saved workflow for a changed input. The model proposes column mappings; code verifies.

The run itself finds moved tables and columns by their header fingerprint. Only headers it cannot
find come here: the model picks a candidate from the unrecognised headers of the same table, code
checks each pick (a real candidate, used once, the value types the column held, a written column
must be empty, a tie between empty columns goes to a person), and every learned check is re-run
before workflow v+1 is saved with each mapping and who decided it.
"""

import json
from pathlib import Path

from .events import Events
from .builder import SMALL, walk
from .planner import run_workflow
from .sources import at, canon, find_tables, mapping_problem, normalise, sheet_rows


PROMPT_VERSION = "adapt-headers-v2"
SYSTEM = """A saved spreadsheet workflow no longer finds some input columns because the input file
changed (headers renamed, columns moved). For each missing column choose the candidate header with
the same business meaning, using the previous header, how the workflow uses the column, the value
types (and sample values, when shown) under each candidate, its role and the customer's description.
A column the workflow writes is empty in the input: an empty candidate is expected and is no reason
to doubt. Choose a candidate when it is the only one with that meaning. Return
null for a column when no candidate clearly has that meaning or two candidates could both be it,
and say in question what a person must confirm. Never choose by position alone; code checks every
choice and re-runs every learned check before anything is saved.
Return JSON: {"mappings": {column: header or null}, "question": text or null}."""
ROUNDS = 4


def usage(saved, key, name):
    """Where the workflow uses a column, so the model sees its role and not only its old name."""
    ref, found = f"{key}.{name}", []
    for step in saved["plan"]["steps"]:
        if step["op"] == "match_receipts":
            if (step["table"], step["name"]) == (key, name):
                found.append("bank deposits are matched to records by this name column")
            continue
        if step["table"] == key and name in step["columns"]:
            found.append(f"the workflow writes this column: {step['intent']}")
        for column, expression in step["columns"].items():
            lookups = [n["lookup"] for n in walk(expression) if "lookup" in n]
            if json.dumps(ref) in json.dumps(expression) or any(l["table"] == key and name in (l["where"], l["return"]) for l in lookups):
                found.append(f"feeds {step['table']}.{column}")
    for check in saved["checks"]:
        lookups = [n["lookup"] for n in walk(check.get("expression")) if "lookup" in n]
        for l in lookups:
            if l["table"] == key and name in (l["where"], l["return"]):
                role = "returns" if name == l["return"] else "is matched against"
                found.append(f"a learned check looks up {check['table']}.{check['column']} in this table: this column {role} "
                             f"{json.dumps(l['equals'])}")
    return sorted(set(found))


def samples(workbook, item):
    """A few values per candidate, only for large reference tables (a chart of accounts), never for people or entities."""
    table = next((t for t in find_tables(sheet_rows(workbook)) if f"{t['sheet']}!row {t['row']}" == item["where"]), None)
    if table is None or len(table["records"]) <= SMALL:
        return {}
    columns = {h: c for c, h in table["headers"].items()}
    return {h: sorted({canon(at(r["cells"], columns[h])) for r in table["records"][:200]} - {None})[:3] for h in item["options"]}


def proposal_schema(missing):
    return {"type": "object", "additionalProperties": False, "required": ["mappings", "question"], "properties": {
        "mappings": {"type": "object", "additionalProperties": False, "required": list(missing), "properties": {
            ref: {"anyOf": [{"type": "string", "enum": item["options"]}, {"type": "null"}]} for ref, item in missing.items()}},
        "question": {"type": ["string", "null"]}}}


def propose(saved, missing, context, client, directory, workbook):
    request = {"customer_context": Path(context).read_text() if context else "", "workflow_intent": saved["plan"]["intent"],
               "missing_columns": []}
    for ref, item in missing.items():
        examples = samples(workbook, item)
        written = item["column"] in saved["outputs"].get(item["table"], {}).get("columns", [])
        request["missing_columns"].append({
            "column": ref, "previous_header": item["previous_header"], "held_at_build": item["types"],
            "role": "the workflow fills this column, so in the input it is empty: pick the empty candidate whose header has this meaning"
                    if written else "the workflow reads this column: pick the candidate with this meaning and the value types it held at build",
            "used_for": usage(saved, item["table"], item["column"]), "table": item["where"],
            "candidates": [{"header": h, "value_types": item["candidate_types"][h], **({"sample_values": examples[h]} if h in examples else {})}
                           for h in item["options"]]})
    (Path(directory) / "adapt-request.json").write_text(json.dumps(request, indent=2))
    client.schema = proposal_schema(missing)
    reply = json.loads(client.complete([{"role": "system", "content": SYSTEM}, {"role": "user", "content": json.dumps(request)}],
                                       Events(Path(directory) / "model"), 1))
    return reply.get("mappings") or {}, reply.get("question")


def adapted_spec(saved, result, answers):
    """The workflow with the input's current sheets and headers, plus the record of what changed."""
    tables, mapped = {}, {}
    for key, table in saved["tables"].items():
        now = result["bindings"][key]
        columns = {}
        for name, meta in table["columns"].items():
            header = now["columns"][name]
            columns[name] = {**meta, "header": header}
            if normalise(header) != normalise(meta["header"]):
                by = answers.get(f"column_mapping:{key}.{name}", {}).get("by")
                mapped[f"{key}.{name}"] = {"from": meta["header"], "to": header,
                                          "decided_by": "model, verified by code" if by == "model" else "person"}
        if f"table_choice:{key}" in answers:
            mapped[key] = {"from": f"{table['sheet']} {table['headers'][:4]}", "to": f"{now['sheet']} {now['headers'][:4]}",
                           "decided_by": "person"}
        tables[key] = {**table, "sheet": now["sheet"], "headers": now["headers"], "columns": columns}
    return tables, mapped


def adapt_workflow(path, workbook, bank, context, client, directory, answers=None, ask=None):
    """answers: {question id: answer} from a person; ask(questions) -> answers prompts one when set."""
    directory = Path(directory)
    directory.mkdir(parents=True, exist_ok=False)
    saved = json.loads(Path(path).read_text())
    version = saved.get("workflow_version", 1)
    answers = {key: {**(value if isinstance(value, dict) else {"answer": value}), "by": "person"}
               for key, value in (answers or {}).items()}
    proposed, asked, notes = set(), set(), {}
    for round_ in range(1, ROUNDS + 1):
        result = run_workflow(path, workbook, bank, directory / f"run-{round_}", saved, answers)
        if result["status"] == "success":
            tables, mapped = adapted_spec(saved, result, answers)
            if not mapped:
                return {**result, "adapted": False, "workflow": str(path),
                        "note": "The existing workflow fits this input; no update needed."}
            updated = {**saved, "workflow_version": version + 1, "tables": tables,
                       "adaptations": saved.get("adaptations", []) + [{"from_version": version, "input": Path(workbook).name, "mapped": mapped}]}
            destination = directory / "workflow.json"
            destination.write_text(json.dumps(updated, indent=2))
            return {**result, "adapted": True, "mapping": mapped, "workflow": str(destination), "workflow_version": version + 1}
        questions = result.get("questions", [])
        if not any(q["code"] in ("column_mapping", "table_choice") for q in questions):
            return {**result, "adapted": False, "note": "Nothing renamed or moved is left to adapt; the run stopped for the reasons listed."}
        missing = {q["id"].split(":", 1)[1]: q for q in questions if q["code"] == "column_mapping"}
        gone = [ref for ref, q in missing.items() if not q["options"]]
        if gone:
            return {**result, "status": "stopped", "adapted": False,
                    "reason": "No unrecognised header is left for " + ", ".join(gone) + "; add the column back or say where it moved."}
        fresh = {ref: q for ref, q in missing.items() if ref not in proposed}
        if fresh and client is not None and not any(q["code"] == "table_choice" for q in questions):
            proposed.update(fresh)
            mappings, model_question = propose(saved, fresh, context, client, directory / f"run-{round_}", workbook)
            taken = {}
            for ref, header in mappings.items():
                if ref not in fresh or header is None:
                    continue
                key, name = fresh[ref]["table"], fresh[ref]["column"]
                problem = mapping_problem(saved, key, name, header, fresh[ref], by_person=False)
                if problem is None and (key, header) in taken:
                    problem = f"'{header}' was also chosen for {taken[(key, header)]}"
                if problem:
                    notes[ref] = f" The model chose '{header}', rejected by code: {problem}"
                else:
                    taken[(key, header)] = ref
                    answers[f"column_mapping:{ref}"] = {"answer": header, "by": "model"}
            for ref in fresh:
                if model_question and f"column_mapping:{ref}" not in answers:
                    notes[ref] = notes.get(ref, "") + f" The model was unsure: {model_question}"
            if taken:
                continue
        pending = [{**q, "question": q["question"] + notes.get(q["id"].split(":", 1)[1], "")} for q in questions]
        fresh_items = [q for q in pending if q["id"] not in asked]
        if ask is None or not fresh_items:
            (directory / "questions.json").write_text(json.dumps(pending, indent=2, default=str))
            return {**result, "status": "needs_input", "adapted": False, "questions": pending}
        asked.update(q["id"] for q in fresh_items)
        for key, reply in (ask(fresh_items) or {}).items():
            answers[key] = {**(reply if isinstance(reply, dict) else {"answer": reply}), "by": "person"}
    return {"status": "stopped", "adapted": False, "reason": f"Still unresolved after {ROUNDS} rounds.", "run_dir": str(directory)}
