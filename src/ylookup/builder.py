"""Phase 1: learn a workflow from the input, a completed target and the customer's description.

Code does what code can be sure of. It diffs the input against the target to learn what the job
produces, searches the input for every target value (candidate expressions), and mines the
conventions every target row follows (checks). The model reads that evidence with the customer's
description and composes the plan. Code runs it, compares with the target and sends back the
differences until nothing differs.
"""

from datetime import date
import hashlib
from itertools import product
import json
from pathlib import Path
import re

from openpyxl.utils import get_column_letter

from .planner import (ARITHMETIC, MAX_LINES, POLICIES, VARS, check_outputs, contexts, currency_questions, entries,
                      environment, evaluate, execute_plan, keyed, learn_checks, receipts_total, references,
                      statement_with_answers, total_checks)
from .sources import (answered, bind_tables, canon, column_types, equivalent, find_tables, kind, match_names,
                      match_receipts, normalise, number, question, sheet_rows)
from .workbook import compare_outputs, plain, read_outputs, write_outputs


PROMPT_VERSION = "workflow-plan-v20"
MAX_ATTEMPTS = 5
# ponytail: tables with up to 50 records are treated as private (people, entities); larger ones as
# reference lists. Replace with an explicit privacy setting if a real roster is bigger.
SMALL = 50
SYSTEM = """You design a reusable spreadsheet workflow. The user message holds evidence that code
gathered from an input workbook, a bank statement and a completed example of the same job:
- customer_context: how the customer describes the job, in their words.
- input_tables: the tables found in the input, with the value types of each column.
- outputs: what the completed example adds: rows appended under an empty header (append_rows) or
  values written into existing records (fill_columns). For appended tables, fits lists the
  iterations whose row count matches the example. For each output column, code lists candidate
  expressions it verified against the example, either for "all lines"/"all rows" or for one line
  of each entry ("line 1", "line 2", ...). Where none was found it shows examples instead.
- name_matching: how many deposits each name column identifies, by full name or by first name
  plus a unique surname prefix.
- recorded_answers: decisions a person has made. Follow them.

Reply with JSON, either {"status":"ready","workflow":PLAN,"question":null,"options":null} or
{"status":"needs_input","workflow":null,"question":"...","options":[...] or null} when something
essential is ambiguous and neither the evidence nor the customer's description settles it.

PLAN = {"intent": text, "steps": [STEP, ...]}. Each step is one tool call with its own intent.
Intents say what the job does for the business, never how an earlier attempt was fixed.
- match_receipts {table, name, policy}: matches each statement deposit to one record of table by
  its name column. policy is full_name_only or full_or_unique_prefix.
- fill_columns {table, columns}: one value per record of table for each column.
- append_rows {table, for_each, only_received, lines, columns}: appends rows under an empty header.
  for_each is a table (one entry per record) or "statement" (one entry per matched deposit);
  only_received keeps only entries with a receipt; lines is the number of rows per entry.
Produce every output exactly once and order the steps so each comes after what it uses.

An expression is null, a literal, or one JSON object:
{"col":"Table.Column"}  the current record's value, or the value of a one-row table
{"var":NAME}  receipt.amount (0 without a receipt), receipt.date, receipt.description,
              entry, line, row (1-based counters in the output table), statement.currency
{"unit":"Table.Column"}  the unit label printed above that column
{"out":"Column"}  another column of the same output row
{"line":[e1,e2,...]}  one expression per line of an entry, in order; outermost only
{"concat":[e,...]}  joins text; no part may be null
{"add"|"sub"|"mul"|"div":[a,b]}
{"lookup":{"table":T,"where":"Column","equals":e,"return":"Column"}}  the row of T whose Column equals e
Each candidate list holds alternatives: pick ONE expression from it. Where a list is given for all
lines, use that one expression without line; where lists are given per line, combine one pick per
line with line. Never put a whole candidate list inside line. Prefer an expression that reads the input over a literal copied from the
example, so the workflow still works when the values change. Use the customer's description to
choose between candidates and fits. V_ and DATE_ tokens stand for anonymised input values; never
write them as literals. Code runs your plan, compares it with the example and sends back what differs."""


def walk(node):
    if isinstance(node, dict):
        yield node
        for value in node.values():
            yield from walk(value)
    elif isinstance(node, list):
        for value in node:
            yield from walk(value)


def table_keys(found):
    """The sheet name for the widest table on a sheet, sheet@row for the others."""
    keys = []
    for table in found:
        rivals = [t for t in found if t["sheet"] == table["sheet"]]
        widest = max(rivals, key=lambda t: (t["width"], -t["row"]))
        keys.append(table["sheet"] if table is widest else f"{table['sheet']}@{table['row']}")
    return keys


def table_spec(table):
    return {"sheet": table["sheet"], "headers": list(table["headers"].values()),
            "columns": {h: {"header": h, "types": column_types(table, c), "position": get_column_letter(c + 1)}
                        for c, h in table["headers"].items()}}


def learn_outputs(spec, bound, targets):
    """What the target changed: rows appended under empty headers, and columns filled in existing records."""
    tbound, _ = bind_tables(spec, targets)
    outputs, target, helpers = {}, {}, {}
    for key, table in bound.items():
        other = tbound.get(key)
        if other is None:
            continue
        if not table["rows"] and other["all"]:
            outputs[key] = {"op": "append_rows", "columns": list(spec["tables"][key]["columns"])}
            target[key] = [r["values"] for r in other["all"]]
        elif table["records"]:
            by_key = {canon(r["key"]) or r["row"]: r["values"] for r in other["records"]}
            pairs = [(canon(r["key"]) or r["row"], r) for r in table["records"]]
            changed = [name for name in spec["tables"][key]["columns"]
                       if any(k in by_key and not equivalent(r["values"][name], by_key[k][name]) for k, r in pairs)]
            if changed:
                outputs[key] = {"op": "fill_columns", "columns": changed}
                target[key] = {k: {name: by_key.get(k, {}).get(name) for name in changed} for k, _ in pairs}
                known = {normalise(h) for h in spec["tables"][key]["headers"]}
                helpers[key] = [h for h in other["headers"].values() if normalise(h) not in known]
    return outputs, target, helpers


def name_evidence(env):
    deposits = [t for t in env["statement"]["transactions"] if t["kind"] == "deposit"]
    found = []
    for key, table in env["tables"].items():
        if len(table["records"]) < 2:
            continue
        for name in table["columns"]:
            names = {i: r["values"][name] for i, r in enumerate(table["records"])
                     if isinstance(r["values"][name], str) and normalise(r["values"][name])}
            if len(names) < 2:
                continue
            full = sum(len(match_names(d["description"], names, "full_name_only")[0]) == 1 for d in deposits)
            either = sum(len(match_names(d["description"], names)[0]) == 1 for d in deposits)
            if either:
                found.append({"table": key, "column": name, "deposits": len(deposits),
                              "matched_full_name_only": full, "matched_full_or_unique_prefix": either})
    return sorted(found, key=lambda e: (-e["matched_full_or_unique_prefix"], -e["matched_full_name_only"]))[:3]


def fits(env, rows):
    """Iterations whose entry count divides the target's row count into 1-9 lines per entry."""
    options = [(key, False) for key, t in env["tables"].items()
               if t["records"] and env["spec"]["outputs"].get(key, {}).get("op") != "append_rows"]
    if env.get("match"):
        options += [(env["match"], True), ("statement", False)]
    result = []
    for for_each, only in options:
        count = len(entries(env, for_each, only))
        if count and rows % count == 0 and rows // count <= MAX_LINES:
            result.append((for_each, only, rows // count))
    return result


def probe(expression, ctxs):
    values = []
    for ctx in ctxs:
        try:
            values.append(evaluate(expression, ctx))
        except (ValueError, TypeError, KeyError, ArithmeticError, IndexError):
            return None
    return values


def simple_refs(env, iterated):
    found, units = [], set()
    for ref, (key, name) in env["refs"].items():
        out = env["spec"]["outputs"].get(key, {})
        if out.get("op") == "append_rows" or name in out.get("columns", []):
            continue
        table = env["tables"][key]
        if key == iterated or len(table["records"]) == 1:
            found.append({"col": ref})
        if table["units"].get(name) and key not in units:
            units.add(key)
            found.append({"unit": ref})
    return found + [{"var": name} for name in VARS]


def value_index(env, key):
    cache = env.setdefault("value_index", {})
    if key not in cache:
        cache[key] = {}
        for i, record in enumerate(env["tables"][key]["records"]):
            for name, value in record["values"].items():
                cache[key].setdefault(name, {}).setdefault(canon(value), set()).add(i)
    return cache[key]


def lookups(env, idx, want, target, column, iterated):
    """{"lookup"} candidates: a reference table row selected by another output column of the same row."""
    found = []
    for key, table in env["tables"].items():
        if len(table["records"]) < 2 or key == iterated or key in env["spec"]["outputs"]:
            continue
        index = value_index(env, key)
        returns = [name for name in index if all(canon(want[i]) in index[name] for i in idx)]
        if not returns:
            continue
        for other in target[0]:
            keys = [canon(target[i].get(other)) for i in idx]
            if other == column or any(k is None for k in keys):
                continue
            for where in index:
                if any(len(rows) > 1 for value, rows in index[where].items() if value is not None):
                    continue  # a lookup key must identify one row of the table
                hits = [index[where].get(k) for k in keys]
                if not all(hits):
                    continue
                for returned in returns:
                    values = [{canon(table["records"][j]["values"][returned]) for j in rows} for rows in hits]
                    if returned != where and all(v == {canon(want[i])} for v, i in zip(values, idx)):
                        found.append({"lookup": {"table": key, "where": where, "equals": {"out": other}, "return": returned}})
    return found


def concats(probed, idx, want, private):
    found = []
    if not all(isinstance(want[i], str) for i in idx):
        return found
    for expression, values in probed:
        parts = set()
        for i in idx:
            value = values[i]
            piece = None if kind(value) is None else (value.isoformat() if kind(value) == "date" else str(value))
            if piece is None or len(piece) < 3 or piece not in want[i] or piece == want[i]:
                break
            start = want[i].index(piece)
            parts.add((want[i][:start], want[i][start + len(piece):]))
        else:
            if len(parts) == 1:
                prefix, suffix = parts.pop()
                pieces = [p.strip() for p in (prefix, suffix) if len(p.strip()) >= 4]
                if not any(p in value or value in p for p in pieces for value in private):
                    found.append({"concat": [p for p in (prefix, expression, suffix) if p != ""]})
    return found


def arithmetic(probed, idx, want):
    if not all(kind(want[i]) == "number" for i in idx):
        return []
    numeric = [(e, v) for e, v in probed if all(kind(v[i]) == "number" for i in idx)]
    found = []
    for (a, av), (b, bv) in product(numeric, numeric):
        if a is b:
            continue
        if all(equivalent(number(av[i]) - number(bv[i]), want[i]) for i in idx):
            found.append({"sub": [a, b]})
        if id(a) < id(b) and all(equivalent(number(av[i]) + number(bv[i]), want[i]) for i in idx):
            found.append({"add": [a, b]})
    return found[:2]


def literal(idx, want, private):
    if len(idx) < 2:
        return []  # one row cannot show that a value is constant
    first = want[idx[0]]
    if kind(first) not in ("text", "number") or isinstance(first, bool) or canon(first) in private:
        return []
    if not all(equivalent(want[i], first) for i in idx):
        return []
    if kind(first) == "number":
        return [int(first) if number(first) == int(number(first)) else float(first)]
    return [first]


def explain(env, ctxs, target, column, lines, probed, private, mask):
    """Candidate expressions reproducing one target column, for all lines or line by line."""
    want = [row.get(column) for row in target]
    iterated = ctxs[0]["table"] if ctxs else None

    def candidates(idx):
        if all(kind(want[i]) is None for i in idx):
            return [None]
        found = [e for e, v in probed if any(kind(v[i]) for i in idx) and all(equivalent(v[i], want[i]) for i in idx)]
        found += lookups(env, idx, want, target, column, iterated) + concats(probed, idx, want, private)
        if not found:
            found += arithmetic(probed, idx, want)
        found = [e for e in found if '"out"' not in json.dumps(e)] + literal(idx, want, private) + \
            [e for e in found if '"out"' in json.dumps(e)]  # anchors first, lookups through other columns last
        return found[:5]

    everything = list(range(len(want)))
    common = candidates(everything)
    if lines == 1:
        return {"all rows": common or {"examples": mask(examples(want, everything))}}
    result = {"all lines": common} if common else {}
    if common and not all('"out"' in json.dumps(e) for e in common):
        return result
    # Nothing, or only lookups through other output columns, fits all lines: keep each line's options
    # too; anchored() drops them again when the column a lookup reads is itself anchored.
    for line in range(1, lines + 1):
        idx = everything[line - 1::lines]
        result[f"line {line}"] = candidates(idx) or {"examples": mask(examples(want, idx))}
    return result


def anchored(columns):
    """Drop per-line alternatives of a column whose all-lines lookup reads an anchored column (no cycle)."""
    def free(column, seen=()):
        groups = columns.get(column, {})
        options = [e for g in groups.values() if isinstance(g, list) for e in g]
        if any('"out"' not in json.dumps(e) for e in groups.get("all lines", [])):
            return True
        if any(k.startswith("line ") for k in groups) and all(
                isinstance(g, list) and any('"out"' not in json.dumps(e) for e in g) for k, g in groups.items() if k.startswith("line ")):
            return True
        return any(column not in seen and all(free(n["out"], seen + (column,)) for n in walk(e) if "out" in n)
                   for e in options if '"out"' in json.dumps(e))
    for column, groups in columns.items():
        common = groups.get("all lines")
        if isinstance(common, list) and common and any(k.startswith("line ") for k in groups) and \
                all(free(n["out"], (column,)) for e in common for n in walk(e) if "out" in n):
            columns[column] = {"all lines": common}
    return columns


def examples(want, idx):
    return [plain(want[i]) for i in idx][:4]


def explained(columns):
    """Columns with a candidate for every row, either for all lines or for each line."""
    def covered(groups):
        return isinstance(groups.get("all lines", groups.get("all rows")), list) or \
            all(isinstance(g, list) for k, g in groups.items() if k.startswith("line ")) and any(k.startswith("line ") for k in groups)
    return sum(covered(groups) for groups in columns.values())


def masker(env, sources):
    """Replace private input values and dates with tokens. Private: text in the small tables the outputs
    read (people, entities, accounts), unless a large reference table (a chart of accounts) also has it.
    Only values are masked, never keys, so table and column names stay valid.
    """
    public = {canon(v) for t in env["tables"].values() if len(t["records"]) > SMALL
              for r in t["records"] for v in r["values"].values() if kind(v) == "text"}
    tokens = {}
    for key in sorted(k for k in sources if k):
        table = env["tables"][key]
        if len(table["records"]) <= SMALL:
            for record in table["records"]:
                for value in record["values"].values():
                    text = canon(value) if kind(value) == "text" else None
                    if text and len(text) >= 3 and text not in public and not re.fullmatch(r"[A-Z]{3}", text):
                        tokens.setdefault(text, f"V_{len(tokens) + 1}")
    for transaction in env["statement"]["transactions"]:  # bank references carry names and account numbers
        tokens.setdefault(canon(transaction["description"]), f"V_{len(tokens) + 1}")
    dates = {}
    ordered = sorted((t for t in tokens if len(t) >= 6), key=len, reverse=True)

    def mask(value):
        if isinstance(value, dict):
            return {k: mask(v) for k, v in value.items()}
        if isinstance(value, (list, tuple)):
            return [mask(v) for v in value]
        if kind(value) == "date":
            return dates.setdefault(canon(value), f"DATE_{len(dates) + 1}")
        if isinstance(value, str):
            if canon(value) in tokens:
                return tokens[canon(value)]
            if canon(value) in dates:
                return dates[canon(value)]
            for original in ordered:
                value = value.replace(original, tokens[original])
            return value
        return plain(value)

    return mask, set(tokens)


def survey(env, outputs, expected, helpers, private, mask):
    """For each output: the iteration that fits, candidate expressions per column, and learned checks."""
    view, checks, shapes, unexplained = {}, [], {}, {}
    for key, out in outputs.items():
        rows = expected[key] if out["op"] == "append_rows" else list(expected[key].values())
        options = fits(env, len(rows)) if out["op"] == "append_rows" else [(key, False, 1)]
        scored = []
        for shape in options:
            ctxs = contexts(env, *shape)
            if len(ctxs) != len(rows):
                continue
            iterated = ctxs[0]["table"] if ctxs else key
            probed = [(e, v) for e in simple_refs(env, iterated) if (v := probe(e, ctxs)) is not None]
            columns = anchored({c: explain(env, ctxs, rows, c, shape[2], probed, private, mask) for c in out["columns"]})
            scored.append((explained(columns), shape, columns, ctxs))
        if not scored:
            view[key] = {"op": out["op"], "target_rows": len(rows), "fits": [], "columns": {}}
            continue
        top = max(score for score, *_ in scored)
        if top == 0 and out["op"] == "fill_columns":
            # Edits nothing in the input can explain (e.g. retyped notes) are reported, not reproduced.
            unexplained[key] = out["columns"]
            continue
        best = [s for s in scored if s[0] == top]
        score, shape, columns, ctxs = best[0]
        shapes[key] = shape
        view[key] = {"op": out["op"], "target_rows": len(rows),
                     "columns_explained": f"{score} of {len(out['columns'])}", "columns": columns}
        if out["op"] == "append_rows":
            view[key]["fits"] = [{"for_each": s[0], "only_received": s[1], "lines": s[2]} for _, s, *_ in best]
        if helpers.get(key):
            view[key]["target_helper_columns_not_reproduced"] = helpers[key]
        checks += learn_checks(key, out["columns"], rows, shape[2])
        checks += total_checks(key, out["columns"], rows, shape[2], receipts_total(env))
        for column, groups in columns.items():
            for group, found_ in groups.items():
                relation = next((e for e in found_ if isinstance(e, dict) and set(e) & {"lookup", *ARITHMETIC}), None) \
                    if isinstance(found_, list) else None
                line = int(group.split()[1]) if group.startswith("line ") else None
                if relation and not any(c.get("expression") == relation and c["column"] == column and c["table"] == key for c in checks):
                    checks.append({"table": key, "kind": "expression", "column": column, "line": line, "expression": relation})
                elif isinstance(found_, list) and found_ and all(e is not None and not isinstance(e, dict) for e in found_):
                    # Only a literal explains it: a constant of the output template (e.g. a debit flag), not input data.
                    checks.append({"table": key, "kind": "constant", "column": column, "line": line, "value": found_[0]})
    return view, checks, shapes, unexplained


def learn(workbook, target, bank, answers):
    """Everything code can establish before the model is asked: outputs, evidence, candidates, checks."""
    found = find_tables(sheet_rows(workbook))
    keys = table_keys(found)
    spec = {"tables": {key: table_spec(t) for key, t in zip(keys, found)}, "outputs": {}}
    bound, _ = bind_tables(spec, found)
    outputs, expected, helpers = learn_outputs(spec, bound, find_tables(sheet_rows(target)))
    if not outputs:
        raise ValueError("The target adds no rows and fills no columns compared with the input; nothing to learn.")
    spec["outputs"] = outputs
    statement, questions = statement_with_answers(bank, answers)
    env = environment(spec, bound, statement, answers)
    env["probe"] = True
    names = name_evidence(env)
    if names:
        best = names[0]
        result = match_receipts(statement, bound[best["table"]], best["column"], "full_or_unique_prefix", {})
        env.update(match=best["table"], receipts=result["receipts"], deposits=result["deposits"])
        questions += currency_questions(env, best["table"])
    settled = [q for q in questions if answered(q, answers) is not None]
    questions = [q for q in questions if answered(q, answers) is None]
    # Pass 1 finds which tables the outputs read; only their values are private (pass 2 masks them).
    first = survey(env, outputs, expected, helpers, set(), lambda value: value)[0]
    sources = {env["refs"][ref][0] for ref in env["refs"] if json.dumps(ref) in json.dumps(first)} | {env.get("match")}
    mask, private = masker(env, sources)
    view, checks, shapes, unexplained = survey(env, outputs, expected, helpers, private, mask)
    for key in unexplained:
        del outputs[key], expected[key]
    used = {ref for ref in env["refs"] if json.dumps(ref) in json.dumps(view)}
    relevant = {env["refs"][ref][0] for ref in used} | set(outputs) | {n["table"] for n in names}
    relevant |= {node["lookup"]["table"] for node in walk(view) if "lookup" in node}
    inputs = [{"table": key, "records": len(bound[key]["records"]),
               "columns": {name: "|".join(meta["types"]) for name, meta in spec["tables"][key]["columns"].items()},
               **({"units": {k: v for k, v in bound[key]["units"].items() if v}} if any(bound[key]["units"].values()) else {})}
              for key in spec["tables"] if key in relevant and outputs.get(key, {}).get("op") != "append_rows"]
    profile = {"input_tables": inputs,
               "statement": {"currency": statement["currency"], "deposits": sum(t["kind"] == "deposit" for t in statement["transactions"]),
                             "withdrawals": sum(t["kind"] != "deposit" for t in statement["transactions"])},
               "name_matching": names, "outputs": [{"table": key, **item} for key, item in view.items()],
               **({"target_changes_not_reproduced": unexplained} if unexplained else {})}
    refs = sorted(ref for ref, (key, name) in env["refs"].items() if key in relevant
                  and outputs.get(key, {}).get("op") != "append_rows" and name not in outputs.get(key, {}).get("columns", []))
    schema_view = {"refs": refs, "units": sorted(r for r in refs if bound[env["refs"][r][0]]["units"].get(env["refs"][r][1])),
                   "labels": sorted({env["refs"][r][1] for r in refs}),
                   "headers": sorted({c for o in outputs.values() for c in o["columns"]}),
                   "tables": sorted(k for k in relevant if bound[k]["records"] and outputs.get(k, {}).get("op") != "append_rows"),
                   "lookup_tables": sorted(k for k in relevant if len(bound[k]["records"]) > 1 and k not in outputs),
                   "outputs": outputs}
    return {"spec": spec, "bound": bound, "statement": statement, "expected": expected, "checks": checks,
            "fits": {key: item.get("fits", []) for key, item in view.items()},
            "profile": profile, "schema_view": schema_view, "mask": mask, "questions": questions, "settled": settled,
            "shapes": shapes}


def reply_schema(view):
    def obj(properties):
        return {"type": "object", "properties": properties, "required": list(properties), "additionalProperties": False}

    def enum(values):
        return {"type": "string", "enum": sorted(values)}

    ref = lambda name: {"$ref": "#/$defs/" + name}
    text = {"type": "string"}
    atom = {"anyOf": [{"type": "null"}, text, {"type": "number"}, {"type": "boolean"}, obj({"col": enum(view["refs"])}),
                      obj({"var": enum(VARS)}), obj({"out": enum(view["headers"])})]
            + ([obj({"unit": enum(view["units"])})] if view["units"] else [])}
    pair = {"type": "array", "items": ref("atom"), "minItems": 2, "maxItems": 2}
    part = {"anyOf": [ref("atom")] + [obj({op: pair}) for op in ARITHMETIC]
            + ([obj({"lookup": obj({"table": enum(view["lookup_tables"]), "where": enum(view["labels"]), "equals": ref("atom"),
                                    "return": enum(view["labels"])})})] if view["lookup_tables"] else [])}
    value = {"anyOf": [ref("part"), obj({"concat": {"type": "array", "items": ref("part"), "minItems": 1}})]}
    expression = {"anyOf": [ref("value"), obj({"line": {"type": "array", "items": ref("value"), "minItems": 1, "maxItems": MAX_LINES}})]}
    variants = []
    if view["tables"]:
        variants.append(obj({"op": enum(["match_receipts"]), "intent": text, "table": enum(view["tables"]),
                             "name": enum(view["labels"]), "policy": enum(POLICIES)}))
    for key, out in view["outputs"].items():
        properties = {"op": enum([out["op"]]), "intent": text, "table": enum([key])}
        if out["op"] == "append_rows":
            properties.update(for_each=enum(view["tables"] + ["statement"]), only_received={"type": "boolean"},
                              lines={"type": "integer"})
        properties["columns"] = obj({column: ref("expression") for column in out["columns"]})
        variants.append(obj(properties))
    plan = obj({"intent": text, "steps": {"type": "array", "items": {"anyOf": variants}, "minItems": 1}})
    root = obj({"status": enum(["needs_input", "ready"]), "workflow": {"anyOf": [plan, {"type": "null"}]},
                "question": {"type": ["string", "null"]}, "options": {"anyOf": [{"type": "array", "items": text}, {"type": "null"}]}})
    root["$defs"] = {"atom": atom, "part": part, "value": value, "expression": expression}
    return root


def unpaste(plan, evidence):
    """A line holding exactly an all-lines candidate list is that list pasted in: use its first candidate."""
    for step in plan.get("steps", []) if isinstance(plan, dict) else []:
        known = evidence.get(step.get("table") if isinstance(step, dict) else None, {})
        for column, expression in (step.get("columns") or {}).items() if isinstance(step.get("columns"), dict) else []:
            options = known.get(column, {}).get("all lines")
            if isinstance(options, list) and len(options) > 1 and expression == {"line": options[:len(expression.get("line", []))]} \
                    if isinstance(expression, dict) else False:
                step["columns"][column] = options[0]


def adopt(plan, previous, wrong, evidence, fits):
    """When the model resends an expression (or an iteration) that already failed, and code verified
    exactly one for that column (or table) on every row, use that one. Logged, so the decision stays visible."""
    if not previous:
        return []
    before = {(s.get("table"), c): e for s in previous["steps"] for c, e in s.get("columns", {}).items()}
    iterations = {s.get("table"): {k: s.get(k) for k in ("for_each", "only_received", "lines")}
                  for s in previous["steps"] if s.get("op") == "append_rows"}
    adopted = []
    for step in plan.get("steps", []):
        if not isinstance(step, dict) or not isinstance(step.get("columns"), dict):
            continue
        table, now = step.get("table"), {k: step.get(k) for k in ("for_each", "only_received", "lines")}
        if (table, "row_count") in wrong and iterations.get(table) == now and len(fits.get(table, [])) == 1 and fits[table][0] != now:
            step.update(fits[table][0])
            adopted.append(f"{table} iteration")
        for column, expression in step["columns"].items():
            key = (step.get("table"), column)
            only = evidence.get(key[0], {}).get(column, {})
            only = only.get("all lines", only.get("all rows")) if len(only) == 1 else None
            if key in wrong and before.get(key) == expression and isinstance(only, list) and len(only) == 1 and only[0] != expression:
                step["columns"][column] = only[0]
                adopted.append(f"{key[0]}.{column}")
    return adopted


def restore(plan, locked):
    """Put back expressions that already matched the target, so a fix elsewhere cannot break them."""
    restored = []
    for step in plan.get("steps", []) if isinstance(plan, dict) else []:
        if isinstance(step, dict) and isinstance(step.get("columns"), dict):
            for (table, column), expression in locked.items():
                if table != step.get("table"):
                    continue
                if column is None:
                    step.update(expression)
                elif step["columns"].get(column) != expression:
                    step["columns"][column] = expression
                    restored.append(f"{table}.{column}")
    return restored


def matching(plan, outputs, differences, errors):
    """(table, column) -> expression for every output column that reproduced the target on every row."""
    broken = {d["table"] for d in differences if "check" in d}
    wrong = {(d["table"], d.get("column")) for d in differences} | {(e["table"], e["column"]) for e in errors}
    locked = {}
    for step in plan["steps"]:
        if step["op"] == "match_receipts" or step["table"] in broken or step["table"] not in outputs:
            continue
        locked[(step["table"], None)] = {k: step[k] for k in ("for_each", "only_received", "lines") if k in step}
        locked.update({(step["table"], c): e for c, e in step["columns"].items() if (step["table"], c) not in wrong})
    return locked


def summarise(differences, plan, errors, fits, evidence):
    lines = {s["table"]: s.get("lines", 1) for s in plan["steps"] if s["op"] != "match_receipts"}
    expressions = {(s["table"], c): e for s in plan["steps"] for c, e in s.get("columns", {}).items()}
    miscounted = {d["table"] for d in differences if d.get("check") == "row_count"}
    groups = {}
    for d in differences:
        if d.get("check") == "row_count":
            step = next(s for s in plan["steps"] if s.get("table") == d["table"] and s["op"] == "append_rows")
            groups[(d["table"], "row_count", None)] = {
                "table": d["table"], "problem": f"your step produced {d['actual']} rows; the example has {d['expected']}. "
                "Row-by-row differences are hidden until the count matches.",
                "your_iteration": {k: step[k] for k in ("for_each", "only_received", "lines")},
                "iterations_that_fit": fits.get(d["table"], [])}
            continue
        if "check" in d or d["table"] in miscounted:
            if "check" in d:
                groups[(d["table"], d["check"], None)] = d
            continue
        line = (d["row"] - 1) % lines.get(d["table"], 1) + 1 if isinstance(d["row"], int) and lines.get(d["table"], 1) > 1 else None
        known = evidence.get(d["table"], {}).get(d["column"], {})
        verified = known.get(f"line {line}") or known.get("all lines") or known.get("all rows")
        group = groups.setdefault((d["table"], d["column"], line), {
            "table": d["table"], "column": d["column"], **({"line": line} if line else {}), "rows_wrong": 0,
            "examples": [], "your_expression": expressions.get((d["table"], d["column"])),
            **({"verified_candidates": verified} if isinstance(verified, list) else {})})
        group["rows_wrong"] += 1
        example = {"expected": d["expected"], "actual": d["actual"]}
        if example not in group["examples"] and len(group["examples"]) < 3:
            group["examples"].append(example)
    seen = {}
    for e in errors:
        seen.setdefault((e["table"], e["column"], e["error"]), e)
    errors_by_column = {}
    for (table, column, _), e in seen.items():
        known = evidence.get(table, {}).get(column, {})
        verified = known.get(f"line {e['line']}") or known.get("all lines") or known.get("all rows")
        if isinstance(verified, list):
            errors_by_column[f"{table}.{column}"] = verified
    lines_ = [f"{t}.{c} (line {e['line']}): {msg}" for (t, c, msg), e in list(seen.items())[:10]]
    lines_ += [f"verified candidates for {name}: {json.dumps(found)}" for name, found in list(errors_by_column.items())[:10]]
    return list(groups.values()), lines_


def prune(spec, plan, checks):
    """Keep only the tables and columns the plan and checks use, with each table's header fingerprint."""
    text = json.dumps([plan, checks])
    keep = {}
    for ref, (key, name) in references(spec).items():
        if json.dumps(ref) in text:
            keep.setdefault(key, set()).add(name)
    for step in plan["steps"]:
        if step["op"] == "match_receipts":
            keep.setdefault(step["table"], set()).add(step["name"])
        else:
            keep.setdefault(step["table"], set()).update(spec["outputs"][step["table"]]["columns"])
            if step.get("for_each") not in (None, "statement"):
                keep.setdefault(step["for_each"], set())
    for node in walk([plan, checks]):
        if "lookup" in node:
            keep.setdefault(node["lookup"]["table"], set()).update({node["lookup"]["where"], node["lookup"]["return"]})
    return {key: {**spec["tables"][key], "columns": {n: m for n, m in spec["tables"][key]["columns"].items() if n in names}}
            for key, names in keep.items()}


def ask_until_settled(items, answers, ask):
    """Ask a person each question once. Returns the questions still open."""
    if not items or ask is None:
        return items
    for key, reply in (ask(items) or {}).items():
        answers[key] = reply
    return [item for item in items if answered(item, answers) is None]


def build_workflow(workbook, target, bank, context, answers, client, events, max_attempts=MAX_ATTEMPTS, ask=None):
    """answers: {question id: answer} given up front; ask(questions) -> answers prompts a person when set."""
    if type(max_attempts) is not int or not 1 <= max_attempts <= MAX_ATTEMPTS:
        raise ValueError(f"Build budget must be between one and {MAX_ATTEMPTS} model calls.")
    if Path(context).suffix.lower() == ".pdf":
        raise ValueError("Provide the customer context as a plain-text file, not a PDF.")
    answers = dict(answers or {})
    asked = {}
    events.emit("run_start", phase="build", model=client.model, prompt_version=PROMPT_VERSION,
                max_attempts=max_attempts, max_tokens_per_call=client.max_tokens)
    try:
        learned = learn(workbook, target, bank, answers)
        asked.update({q["id"]: q for q in learned["questions"] + learned["settled"]})
        if ask_until_settled(learned["questions"], answers, ask) != learned["questions"]:
            learned = learn(workbook, target, bank, answers)
        if learned["questions"]:
            return needs_input(events, learned["questions"])
        spec, mask = learned["spec"], learned["mask"]
        profile = {"customer_context": Path(context).read_text(),
                   "recorded_answers": [{"question": asked[k]["question"], "answer": answered(asked[k], answers)}
                                        for k in answers if k in asked],
                   **learned["profile"]}
        (events.directory / "evidence.json").write_text(json.dumps(mask(profile), indent=2))
        events.emit("learned", outputs=spec["outputs"], checks=len(learned["checks"]), shapes=learned["shapes"])
        client.schema = reply_schema(learned["schema_view"])
        messages = [{"role": "system", "content": SYSTEM}, {"role": "user", "content": json.dumps(mask(profile))}]
        locked, tried, previous_plan, previous_wrong = {}, [], None, set()
        for attempt in range(1, max_attempts + 1):
            content = client.complete(messages, events, attempt)
            messages.append({"role": "assistant", "content": content})
            folder = events.directory / f"candidate-{attempt}"
            folder.mkdir()
            try:
                reply = json.loads(content)
                if not isinstance(reply, dict):
                    raise ValueError("Expected a JSON object.")
                if reply.get("status") == "needs_input":
                    if not isinstance(reply.get("question"), str) or not reply["question"].strip():
                        raise ValueError("A needs_input reply must carry a nonempty question.")
                    item = question("model_question", reply["question"], hashlib.sha1(reply["question"].encode()).hexdigest()[:10],
                                    options=reply.get("options") or None, scope="workflow")
                    asked[item["id"]] = item
                    events.emit("inquiry", attempt=attempt, question=item)
                    if answered(item, answers) is None and ask_until_settled([item], answers, ask):
                        return needs_input(events, [item])
                    feedback = {"answer": answered(item, answers), "instruction": "A person answered your question. Use it and return the workflow."}
                else:
                    if reply.get("status") != "ready" or not isinstance(reply.get("workflow"), dict):
                        raise ValueError("Expected status ready with a workflow, or needs_input with a question.")
                    plan = reply["workflow"]
                    evidence = {o["table"]: o["columns"] for o in learned["profile"]["outputs"]}
                    unpaste(plan, evidence)
                    restored = restore(plan, locked)
                    adopted = adopt(plan, previous_plan, previous_wrong, evidence, learned["fits"])
                    if adopted:
                        events.emit("adopted_candidates", attempt=attempt, columns=adopted)
                    repeat = next((n for n, earlier in enumerate(tried, 1) if earlier == plan), None)
                    tried.append(json.loads(json.dumps(plan)))
                    if restored:
                        events.emit("locked_fields", attempt=attempt, restored=restored)
                    (folder / "plan.json").write_text(json.dumps(plan, indent=2))
                    events.emit("candidate", attempt=attempt, plan=plan)
                    env = environment(spec, learned["bound"], learned["statement"], answers)
                    outputs, errors = execute_plan(plan, spec, env, events)
                    differences = compare_outputs({key: keyed(o) for key, o in outputs.items()}, learned["expected"])
                    failures = check_outputs(learned["checks"], outputs, env)
                    (folder / "differences.json").write_text(json.dumps(differences, indent=2, default=str))
                    events.emit("comparison", attempt=attempt, status="fail" if differences or errors else "pass",
                                differences=differences, expression_errors=len(errors), check_failures=failures)
                    if not differences and not errors and not env["questions"] and not failures:
                        output = write_outputs(workbook, folder / "output.xlsx", learned["bound"], outputs)
                        written = compare_outputs(read_outputs(output, spec), learned["expected"])
                        if written:
                            raise ValueError(f"The written workbook differs from the target in {len(written)} cells.")
                        return save(events, spec, plan, learned["checks"], asked, answers, attempt, output, outputs, workbook, target)
                    miscounted = {d["table"] for d in differences if "check" in d}  # cells misalign; not evidence
                    previous_plan = plan
                    previous_wrong = {(d["table"], d.get("column")) for d in differences if d["table"] not in miscounted} | \
                        {(e["table"], e["column"]) for e in errors} | {(t, "row_count") for t in miscounted}
                    summary, error_lines = summarise(differences, plan, errors, learned["fits"],
                                                     {o["table"]: o["columns"] for o in learned["profile"]["outputs"]})
                    feedback = {"differences": summary, "total_differences": len(differences)}
                    if error_lines:
                        feedback["expression_errors"] = error_lines
                    if failures:
                        feedback["check_failures"] = failures[:8]
                    if env["questions"]:
                        feedback["questions_your_plan_raised"] = [q["question"] for q in env["questions"].values()][:5]
                    if repeat:
                        feedback["repeated"] = f"This plan is identical to attempt {repeat}. Change the reported columns; use their verified_candidates."
                    locked = matching(plan, outputs, differences, errors)
            except (ValueError, TypeError, KeyError, ArithmeticError, IndexError, AttributeError) as error:
                feedback = {"validation_error": str(error)}
                events.emit("validation", attempt=attempt, status="fail", **feedback)
            if locked:
                feedback["locked_columns"] = sorted(f"{t}.{c}" for t, c in locked if c)
            messages.append({"role": "user", "content": json.dumps(mask({
                "instruction": "The candidate failed. Fix every reported issue and return the complete corrected workflow. "
                               "Columns in locked_columns already match the example and are kept as they are.",
                **feedback}))})
        events.emit("run_end", status="budget_exhausted", reason="No candidate matched the target within the call budget.")
        return {"status": "budget_exhausted", "run_dir": str(events.directory)}
    except Exception as error:
        events.emit("run_end", status="error", reason=str(error))
        raise


def needs_input(events, questions):
    path = events.directory / "questions.json"
    path.write_text(json.dumps(questions, indent=2, default=str))
    events.emit("run_end", status="needs_input", questions=questions)
    return {"status": "needs_input", "questions": questions, "question_file": str(path), "run_dir": str(events.directory)}


def save(events, spec, plan, checks, asked, answers, attempt, output, outputs, workbook, target):
    recorded = [{"id": key, "question": item["question"], "answer": answered(item, answers), "by": "person"}
                for key, item in asked.items() if item.get("scope") == "workflow" and answered(item, answers) is not None]
    saved = {"schema_version": 3, "workflow_version": 1, "plan": plan, "tables": prune(spec, plan, checks),
             "outputs": spec["outputs"], "checks": checks, "answers": recorded,
             "built_from": {"input": Path(workbook).name, "target": Path(target).name, "prompt_version": PROMPT_VERSION}}
    path = events.directory / "workflow.json"
    with path.open("x", encoding="utf-8") as stream:
        json.dump(saved, stream, indent=2)
    events.emit("run_end", status="success", attempts=attempt, workflow="workflow.json", output=str(output))
    return {"status": "success", "attempts": attempt, "workflow": str(path), "output": str(output),
            "rows": {key: len(o["rows"]) for key, o in outputs.items()}, "run_dir": str(events.directory)}
