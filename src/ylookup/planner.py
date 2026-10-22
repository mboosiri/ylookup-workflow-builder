"""Workflow plans: a small expression language, plan validation, execution, and checks learned from a target.

Nothing here knows the business. Tables, columns, output headers and conventions all come from the
saved workflow, which the builder learned from the input, the completed target and the context.
"""

from datetime import date
from decimal import Decimal
import json
from pathlib import Path
import re

from .events import Events
from .sources import (answered, bind_tables, binding_changes, canon, equivalent, kind, link_rows, number, question,
                      read_sources)
from .workbook import compare_outputs, plain, read_outputs, write_outputs


VARS = ["entry", "line", "row"]
AGGREGATES = ["sum", "first"]
ARITHMETIC = ["add", "sub", "mul", "div"]
POLICIES = ["full_name_only", "full_or_unique_prefix"]
MAX_LINES = 9
TOKEN = re.compile(r"\b(?:V|DATE)_\d+\b")


def references(spec):
    return {f"{key}.{name}": (key, name) for key, table in spec["tables"].items() for name in table["columns"]}


# --- expressions -----------------------------------------------------------------------------------

def source_value(env, ref, record):
    """A source cell; a blank where the build saw none becomes a question (a recorded answer fills it)."""
    key, name = env["refs"][ref]
    value = record["values"].get(name)
    types = env["spec"]["tables"][key]["columns"][name]["types"]
    if kind(value) is not None or "blank" in types or env.get("probe"):
        return value
    formula = record["formulas"].get(name)
    item = question("missing_value", f"{ref} is blank in {env['tables'][key]['sheet']} row {record['row']}"
                    + (f" (formula {formula} has no saved result)" if formula else "") + ". What value should be used?",
                    f"{ref} row {record['row']}", table=key, column=name, row=record["row"])
    reply = answered(item, env["answers"])
    if reply is None:
        env["questions"][item["id"]] = item
        return None
    return number(reply) if types == ["number"] else reply


def text(value):
    return value.isoformat() if kind(value) == "date" else str(value)


def lookup(env, table_key, where, wanted, returned):
    index = env.setdefault("index", {}).get((table_key, where))
    if index is None:
        index = {}
        for record in env["tables"][table_key]["records"]:
            index.setdefault(canon(record["values"][where]), []).append(record)
        env["index"][(table_key, where)] = index
    values = {canon(r["values"][returned]): r["values"][returned] for r in index.get(canon(wanted), [])}
    if len(values) != 1:
        raise ValueError(f"lookup in {table_key}: {len(values)} distinct {returned} values where {where} = {wanted!r}")
    return next(iter(values.values()))


def related(ctx, key):
    """The record of table `key` a row reads: its own, the one it is linked to, or a one-row table's."""
    env = ctx["env"]
    if key == ctx.get("table"):
        return ctx["record"]
    link = env["links"].get(ctx.get("table"))
    if link and link["to"] == key:
        index = link["rows"].get(ctx["index"])
        return None if index is None else env["tables"][key]["records"][index]
    if len(env["tables"][key]["records"]) == 1:
        return env["tables"][key]["records"][0]
    raise ValueError(f"{key} has {len(env['tables'][key]['records'])} rows and this step neither iterates it nor is linked "
                     "to it; use lookup, or sum/first over linked rows.")


def linked_rows(ctx, key):
    back = ctx["env"]["linked"].get(ctx.get("table"))
    if not back or back["from"] != key:
        raise ValueError(f"sum/first of {key} needs a link_rows step linking {key} to {ctx.get('table')} earlier in the plan.")
    return [ctx["env"]["tables"][key]["records"][j] for j in back["records"].get(ctx["index"], [])]


def aggregate(op, arg, ctx):
    """sum: total over the rows linked to this record (0 without any); first: their value, asked when they differ."""
    env = ctx["env"]
    key, name = env["refs"][arg]
    rows = linked_rows(ctx, key)
    values = [r["values"].get(name) for r in rows if kind(r["values"].get(name))]
    if op == "sum":
        return sum((number(v) for v in values), Decimal(0))
    distinct = {canon(v): v for v in values}
    if len(distinct) <= 1:
        return next(iter(distinct.values()), None)
    where = f"{ctx['table']} row {ctx['record']['row']}"
    item = question("linked_values", f"{where} is linked to {len(rows)} {key} rows with different {name}. Which one applies?",
                    f"{arg} for {where}", options=sorted(text(v) for v in distinct.values()))
    reply = answered(item, env["answers"])
    if reply is None:
        env["questions"][item["id"]] = item
        return None
    return next((v for v in distinct.values() if text(v) == reply), None)


def evaluate(expression, ctx):
    if not isinstance(expression, dict):
        return expression
    (op, arg), = expression.items()
    env = ctx["env"]
    if op == "col":
        key, _ = env["refs"][arg]
        record = related(ctx, key)
        return None if record is None else source_value(env, arg, record)
    if op in AGGREGATES:
        return aggregate(op, arg, ctx)
    if op == "unit":
        key, name = env["refs"][arg]
        return env["tables"][key]["units"].get(name)
    if op == "var":
        return ctx[arg]
    if op == "out":
        return ctx["out"](arg)
    if op == "line":
        return evaluate(arg[ctx["line"] - 1], ctx)
    if op == "concat":
        parts = [evaluate(item, ctx) for item in arg]
        if any(part is None for part in parts):
            raise ValueError(f"concat got a blank part from {json.dumps(arg)}; put line outermost if one line is blank.")
        return "".join(text(part) for part in parts)
    if op in ARITHMETIC:
        left, right = (number(evaluate(item, ctx)) for item in arg)
        if op == "div":
            if right == 0:
                raise ValueError("division by zero")
            return left / right
        return {"add": left + right, "sub": left - right, "mul": left * right}[op]
    if op == "lookup":
        return lookup(env, arg["table"], arg["where"], evaluate(arg["equals"], ctx), arg["return"])
    raise ValueError(f"unknown expression {op}")


def emit(columns, ctx):
    """One output row; {"out": header} reads another column of the same row, evaluated on demand."""
    values, busy, errors = {}, set(), []

    def out(name):
        if name not in values:
            if name in busy:
                raise ValueError(f"circular reference: {name} depends on itself through out/lookup; give one column of "
                                 "the pair a value that does not read the other (for example its per-line candidate)")
            busy.add(name)
            values[name] = evaluate(columns[name], ctx)
        return values[name]

    ctx["out"] = out
    for name in columns:
        try:
            out(name)
        except (ValueError, TypeError, KeyError, ArithmeticError, IndexError) as error:
            values[name] = None
            errors.append((name, str(error)))
    return {name: values[name] for name in columns}, errors


def entries(env, for_each, only_linked):
    """(table, record, index) for each entry of an output step; only_linked keeps records with a link."""
    link, back = env["links"].get(for_each), env["linked"].get(for_each)
    result = []
    for index, record in enumerate(env["tables"][for_each]["records"]):
        if only_linked and not ((link and link["rows"].get(index) is not None) or (back and back["records"].get(index))):
            continue
        result.append((for_each, record, index))
    return result


def contexts(env, for_each, only_linked, lines):
    rows = []
    for entry, (table, record, index) in enumerate(entries(env, for_each, only_linked), 1):
        for line in range(1, lines + 1):
            rows.append({"env": env, "table": table, "record": record, "index": index,
                         "entry": entry, "line": line, "row": len(rows) + 1})
    return rows


# --- plans -----------------------------------------------------------------------------------------

def validate_expression(expression, refs, headers, lines, depth=0):
    if depth > 6:
        raise ValueError("expression nesting is too deep")
    if expression is None or type(expression) in {str, int, float, bool}:
        if isinstance(expression, str) and OPERATION.match(expression):
            raise ValueError(f"'{expression[:60]}' is a string; expressions must be JSON objects such as {{\"col\": \"Table.Column\"}}")
        if isinstance(expression, str) and TOKEN.search(expression):
            raise ValueError(f"'{expression}' contains an anonymised token; reference the input column instead of copying it")
        return
    if not isinstance(expression, dict) or len(expression) != 1:
        raise ValueError(f"expected null, a literal or a one-key object, got {json.dumps(expression)[:80]}")
    (op, arg), = expression.items()
    inner = lambda item: validate_expression(item, refs, headers, lines, depth + 1)
    if op in ("col", "unit", *AGGREGATES):
        if arg not in refs:
            raise ValueError(f"unknown column {arg!r}")
    elif op == "var":
        if arg not in VARS:
            raise ValueError(f"unknown var {arg!r}; use one of {VARS}")
    elif op == "out":
        if arg not in headers:
            raise ValueError(f"unknown output column {arg!r}")
    elif op == "line":
        if not isinstance(arg, list) or len(arg) != lines:
            raise ValueError(f"line needs exactly {lines} expressions, one per line of the entry")
        for item in arg:
            if isinstance(item, dict) and "line" in item:
                raise ValueError("line cannot be nested")
            inner(item)
    elif op == "concat":
        if not isinstance(arg, list) or not arg:
            raise ValueError("concat needs a nonempty list")
        for item in arg:
            inner(item)
    elif op in ARITHMETIC:
        if not isinstance(arg, list) or len(arg) != 2:
            raise ValueError(f"{op} needs two expressions")
        for item in arg:
            inner(item)
    elif op == "lookup":
        if not isinstance(arg, dict) or set(arg) != {"table", "where", "equals", "return"}:
            raise ValueError("lookup needs exactly table, where, equals, return")
        for part in ("where", "return"):
            if f"{arg['table']}.{arg[part]}" not in refs:
                raise ValueError(f"lookup {part}: unknown column {arg['table']}.{arg[part]}")
        inner(arg["equals"])
    else:
        raise ValueError(f"unknown expression operation {op!r}")


PSEUDO = re.compile(r"^\{?\s*['\"]?(col|var|unit|out|sum|first)['\"]?\s*:\s*['\"]?([^\[\]{},]+?)['\"]?\s*\}?$")
OPERATION = re.compile(r"^\s*\{?\s*['\"]?(col|var|unit|out|sum|first|line|concat|lookup|add|sub|mul|div)['\"]?\s*:")


def unwrap(expression, lines):
    """Repair two common small-model slips without changing meaning; the result is still validated.

    A line choice that cannot vary means its first item ({"line": [e]}, [e, e], or any line in a one-line table),
    and a simple reference written as a string ("{col: Table.Column}") means that reference.
    """
    if isinstance(expression, str) and PSEUDO.match(expression.strip()):
        op, arg = PSEUDO.match(expression.strip()).groups()
        return {op: arg}
    if isinstance(expression, list):
        return [unwrap(item, lines) for item in expression]
    if not isinstance(expression, dict) or len(expression) != 1:
        return expression
    (op, arg), = expression.items()
    if op == "line" and isinstance(arg, list) and arg and (len(arg) == 1 or lines == 1 or all(item == arg[0] for item in arg)):
        return unwrap(arg[0], lines)
    if op == "lookup" and isinstance(arg, dict):
        return {op: {**arg, "equals": unwrap(arg.get("equals"), lines)}}
    return {op: unwrap(arg, lines) if isinstance(arg, list) else arg}


def normalise_plan(plan):
    """Apply unwrap() to every column expression, in place; returns the plan."""
    for step in plan.get("steps", []) if isinstance(plan, dict) and isinstance(plan.get("steps"), list) else []:
        if isinstance(step, dict) and isinstance(step.get("columns"), dict):
            step["columns"] = {name: unwrap(e, step.get("lines", 1)) for name, e in step["columns"].items()}
    return plan


def validate_plan(plan, spec):
    """Structure, references and dependencies; every learned output covered exactly once."""
    if not isinstance(plan, dict) or set(plan) != {"intent", "steps"} or not isinstance(plan["steps"], list):
        raise ValueError("Plan validation failures: expected exactly intent and steps.")
    normalise_plan(plan)
    refs, errors, covered, matched = references(spec), [], set(), False
    for number_, step in enumerate(plan["steps"], 1):
        where = f"step {number_}"
        if not isinstance(step, dict) or not isinstance(step.get("intent"), str) or not step["intent"].strip():
            errors.append(f"{where}: needs an op and a nonempty intent")
            continue
        op, key = step.get("op"), step.get("table")
        if op == "link_rows":
            if set(step) != {"op", "intent", "table", "text", "to", "name", "policy"} or f"{key}.{step['text']}" not in refs \
                    or f"{step['to']}.{step['name']}" not in refs or key == step["to"]:
                errors.append(f"{where}: link_rows needs table and text (its column holding names), to and name "
                              "(the other table and its name column) and policy")
            elif step["policy"] not in POLICIES:
                errors.append(f"{where}: policy must be one of {POLICIES}")
            matched = True
            continue
        output = spec["outputs"].get(key)
        if op not in ("fill_columns", "append_rows") or output is None or output["op"] != op:
            errors.append(f"{where}: {op} on {key!r} is not one of the learned outputs "
                          f"{ {k: v['op'] for k, v in spec['outputs'].items()} }")
            continue
        if key in covered:
            errors.append(f"{where}: {key} is produced twice")
        covered.add(key)
        allowed = {"op", "intent", "table", "columns"} | ({"for_each", "only_linked", "lines"} if op == "append_rows" else set())
        if set(step) != allowed:
            errors.append(f"{where}: {op} takes exactly {sorted(allowed)}")
            continue
        lines = step.get("lines", 1)
        if op == "append_rows":
            if step["for_each"] not in spec["tables"]:
                errors.append(f"{where}: for_each must be a table name")
            if type(step["only_linked"]) is not bool or type(lines) is not int or not 1 <= lines <= MAX_LINES:
                errors.append(f"{where}: only_linked must be boolean and lines an integer 1-{MAX_LINES}")
                continue
        columns = step["columns"]
        if not isinstance(columns, dict) or set(columns) != set(output["columns"]):
            errors.append(f"{where}: columns must be exactly {output['columns']}")
            continue
        needs_link = step.get("only_linked") or re.search(r'"(sum|first)":', json.dumps(columns))
        if needs_link and not matched:
            errors.append(f"{where}: uses linked rows, so a link_rows step must come before it")
        for name, expression in columns.items():
            try:
                validate_expression(expression, refs, columns, lines)
            except (ValueError, TypeError) as error:
                errors.append(f"{key}.{name}: {error}")
    for key in sorted(set(spec["outputs"]) - covered):
        errors.append(f"no step produces {key} ({spec['outputs'][key]['op']})")
    if errors:
        raise ValueError("Plan validation failures: " + " | ".join(errors))
    return plan


def environment(spec, bound, answers):
    return {"spec": spec, "refs": references(spec), "tables": bound, "answers": answers or {}, "questions": {},
            "links": {}, "linked": {}}


def unit_questions(env, source, target):
    left = sorted({u for u in env["tables"][source]["units"].values() if u})
    right = sorted({u for u in env["tables"][target]["units"].values() if u})
    if not left or not right or left == right:
        return []
    stop = "Stop: the amounts need converting first"
    return [question("currency_conflict", f"{source} is in {', '.join(left)} but {target} states its amounts in "
                     f"{', '.join(right)}. How should the amounts be posted?", f"{'/'.join(left)} vs {'/'.join(right)}",
                     options=["Post the amounts as they are, without FX conversion", stop], stop=stop,
                     scope="workflow", source_units=left, target_units=right)]


def run_link(env, source, text_, target, name, policy):
    result = link_rows(env["tables"][source], text_, env["tables"][target], name, policy, env["answers"], source, target)
    env["links"][source] = {"to": target, "rows": result["rows"], "left": result["left"]}
    env["linked"][target] = {"from": source, "records": result["records"]}
    for item in result["questions"] + [q for q in unit_questions(env, source, target) if answered(q, env["answers"]) is None]:
        env["questions"][item["id"]] = item
    return result


def execute_plan(plan, spec, env, events):
    """Run each step on bound sources. Returns ({table: output}, expression errors)."""
    validate_plan(plan, spec)
    outputs, errors = {}, []
    for step in plan["steps"]:
        events.emit("operation_start", operation=step["op"], table=step["table"], intent=step["intent"])
        if step["op"] == "link_rows":
            result = run_link(env, step["table"], step["text"], step["to"], step["name"], step["policy"])
            events.emit("linked", rows=len(result["rows"]), linked=sum(r is not None for r in result["rows"].values()),
                        questions=len(result["questions"]))
        else:
            lines = step.get("lines", 1)
            rows = []
            ctxs = contexts(env, step.get("for_each", step["table"]), step.get("only_linked", False), lines)
            for ctx in ctxs:
                values, problems = emit(step["columns"], ctx)
                rows.append(values)
                errors += [{"table": step["table"], "row": ctx["row"], "line": ctx["line"], "column": name, "error": message}
                           for name, message in problems]
            outputs[step["table"]] = {"op": step["op"], "rows": rows, "contexts": ctxs, "lines": lines, "columns": step["columns"]}
        events.emit("operation_end", operation=step["op"], table=step["table"])
    return outputs, errors


def keyed(output):
    """Output rows in the shape read_outputs returns: a list to append, or fill values by record key."""
    if output["op"] == "append_rows":
        return output["rows"]
    return {canon(c["record"]["key"]) or c["record"]["row"]: row for c, row in zip(output["contexts"], output["rows"])}


# --- checks learned from the target ----------------------------------------------------------------

def learn_checks(key, columns, target, lines):
    """Conventions every target row of an output table follows, to be enforced on every future run.

    Generic kinds only: blank/required per line, one value per line, equal across an entry's lines and
    row/entry/line counters. The builder adds totals that tie to linked rows, template constants and
    relations (lookups, arithmetic) found in the target.
    """
    checks, n = [], len(target)
    entries_ = n // lines if lines else 0
    for column in columns:
        values = [row.get(column) for row in target]
        for line in range(1, lines + 1):
            part = values[line - 1::lines]
            if all(kind(v) is None for v in part):
                checks.append({"table": key, "kind": "blank", "column": column, "line": line})
            elif all(kind(v) for v in part):
                checks.append({"table": key, "kind": "required", "column": column, "line": line})
                if entries_ >= 2 and all(equivalent(v, part[0]) for v in part):
                    checks.append({"table": key, "kind": "uniform", "column": column, "line": line})
        if lines > 1 and entries_ >= 2 and any(kind(v) for v in values) and \
                all(equivalent(values[i], values[i - i % lines]) for i in range(n)):
            checks.append({"table": key, "kind": "same_in_entry", "column": column})
        if n >= 2 and all(kind(v) == "number" for v in values):
            numbers = [number(v) for v in values]
            for name, expected in [("sequence", range(1, n + 1)), ("entry_counter", [i // lines + 1 for i in range(n)]),
                                   ("line_counter", [i % lines + 1 for i in range(n)])]:
                if lines == 1 and name != "sequence":
                    continue
                if numbers == list(expected):
                    checks.append({"table": key, "kind": name, "column": column})
                    break
    return checks


def link_totals(env):
    """Per linked table and numeric column: the total over the rows that were linked."""
    totals = {}
    for source, link in env["links"].items():
        records = env["tables"][source]["records"]
        rows = [records[i] for i, record in link["rows"].items() if record is not None]
        for name in env["tables"][source]["columns"]:
            values = [r["values"].get(name) for r in records]  # numeric by the whole column, so no links means 0
            if any(kind(v) == "number" for v in values) and all(kind(v) in (None, "number") for v in values):
                totals[f"{source}.{name}"] = sum((number(r["values"].get(name)) for r in rows if kind(r["values"].get(name))), Decimal(0))
    return totals


def total_checks(key, columns, target, lines, totals):
    checks = []
    for column in columns:
        for line in range(1, lines + 1):
            part = [row.get(column) for row in target][line - 1::lines]
            if not all(kind(v) in (None, "number") for v in part):
                continue
            got = sum((number(v) for v in part if kind(v)), Decimal(0))
            source = next((ref for ref, total in totals.items() if total and total == got), None)
            if source:
                checks.append({"table": key, "kind": "total", "column": column, "line": line, "of": source})
    return checks


def left_unlinked(output, column, env):
    """Output rows built from a source row a person chose to leave unlinked, when this column reads through
    that link. The example only shows linked rows, so its conventions do not apply to these."""
    expression = json.dumps(output.get("columns", {}).get(column))
    rows = set()
    for i, ctx in enumerate(output["contexts"]):
        link = env["links"].get(ctx["table"])
        # ponytail: "reads through the link" = the expression names the linked table; a lookup keyed by
        # another output column that itself reads the link is not traced.
        if link and ctx["index"] in link.get("left", ()) and f'"{link["to"]}.' in expression:
            rows.add(i)
    return rows


def check_outputs(checks, outputs, env):
    """Every learned convention, re-checked on this run's rows. Returns plain-language failures.

    Rows a person left unlinked are excused from checks on columns that read the link, and from linked totals."""
    failures, totals = [], link_totals(env)
    for check in checks:
        output = outputs.get(check["table"])
        if output is None:
            continue
        rows, lines, column, line = output["rows"], output["lines"], check["column"], check.get("line")
        values = [row.get(column) for row in rows]
        excused = left_unlinked(output, column, env)
        if check["kind"] == "total":  # a total ties to linked rows only, so rows left unlinked are not part of it
            excused = {i for i, ctx in enumerate(output["contexts"])
                       if ctx["index"] in env["links"].get(ctx["table"], {}).get("left", ())}
        indexes = [i for i in (list(range(len(rows)))[line - 1::lines] if line else range(len(rows))) if i not in excused]
        part = [values[i] for i in indexes]
        name = f"{check['table']}.{column}" + (f" (line {line})" if line else "")
        kind_ = check["kind"]
        problem = None
        if kind_ == "required" and any(kind(v) is None for v in part):
            problem = f"{name} is blank on {sum(kind(v) is None for v in part)} rows; the completed example always fills it"
        elif kind_ == "blank" and any(kind(v) for v in part):
            problem = f"{name} holds values; the completed example always leaves it blank"
        elif kind_ == "uniform" and any(not equivalent(v, part[0]) for v in part):
            problem = f"{name} differs between rows ({sorted({str(v) for v in part})[:4]}); the example holds one value"
        elif kind_ == "constant" and any(not equivalent(v, check["value"]) for v in part):
            problem = f"{name} must be {check['value']!r} on every row, as in the example; found {sorted({str(v) for v in part})[:4]}"
        elif kind_ == "same_in_entry" and any(not equivalent(values[i], values[i - i % lines]) for i in range(len(rows))):
            problem = f"{name} differs between the lines of one entry; the example keeps the lines equal"
        elif kind_ in ("sequence", "entry_counter", "line_counter"):
            expected = {"sequence": lambda i: i + 1, "entry_counter": lambda i: i // lines + 1,
                        "line_counter": lambda i: i % lines + 1}[kind_]
            if any(kind(v) != "number" or number(v) != expected(i) for i, v in enumerate(values)):
                problem = f"{name} is not the {kind_.replace('_', ' ')} 1, 2, 3... the example uses"
        elif kind_ == "total":
            got, total = sum((number(v) for v in part if kind(v) == "number"), Decimal(0)), totals.get(check["of"])
            if got != total:
                problem = f"{name} totals {got}, but the linked {check['of']} total {total}"
        elif kind_ == "expression":
            for i in indexes:
                ctx = {**output["contexts"][i], "out": lambda h, row=rows[i]: row.get(h)}
                try:
                    ok = equivalent(evaluate(check["expression"], ctx), values[i])
                except (ValueError, TypeError, KeyError, ArithmeticError):
                    ok = False
                if not ok:
                    problem = f"{name} breaks the relation {json.dumps(check['expression'])} seen in the example (row {i + 1})"
                    break
        if problem:
            failures.append(problem)
    return failures


# --- run -------------------------------------------------------------------------------------------

def run_workflow(path, workbook, sources, directory, saved=None, answers=None):
    """Phase 2: execute a saved workflow on the input workbook plus any other source files (workbooks, PDFs).
    No target, transcript or model."""
    saved = saved if saved is not None else json.loads(Path(path).read_text())
    if not isinstance(saved, dict) or saved.get("schema_version") != 3:
        raise ValueError("Expected a version-3 workflow produced by `ylookup build`.")
    events = Events(directory)
    given = answers or {}
    answers = {**{a["id"]: a for a in saved.get("answers", [])}, **given}
    events.emit("run_start", phase="run", workflow=str(path), workflow_version=saved.get("workflow_version"))
    base = {"run_dir": str(directory)}
    try:
        found, source_issues = read_sources(workbook, sources)
        bound, problems = bind_tables(saved, found, answers)
        layout = {key: {"sheet": t["sheet"], "headers": list(t["headers"].values()),
                        "columns": {name: t["headers"][c] for name, c in t["columns"].items()}} for key, t in bound.items()}
        base.update(binding_changes=binding_changes(saved, bound), bindings=layout)
        if problems:
            return stop(events, directory, {**base, "status": "needs_input", "questions": problems,
                                            "reason": "Some workflow tables or columns are not in this input."})
        open_questions = [q for q in source_issues if answered(q, answers) is None]
        env = environment(saved, bound, answers)
        outputs, errors = execute_plan(saved["plan"], saved, env, events)
        questions = open_questions + list(env["questions"].values())
        if questions:
            return stop(events, directory, {**base, "status": "needs_input", "questions": questions})
        if errors:
            return stop(events, directory, {**base, "status": "stopped", "reason": "; ".join(
                f"{e['table']} row {e['row']} {e['column']}: {e['error']}" for e in errors[:10])})
        failures = check_outputs(saved["checks"], outputs, env)
        if failures:
            return stop(events, directory, {**base, "status": "stopped", "check_failures": failures,
                                            "reason": "Learned checks failed; no workbook was written: " + "; ".join(failures[:5])})
        output = write_outputs(workbook, Path(directory) / "output.xlsx", bound, outputs)
        expected = {key: keyed(o) for key, o in outputs.items()}
        now = {**saved, "tables": {key: {**table, "headers": layout[key]["headers"], "columns": {
            name: {**meta, "header": layout[key]["columns"][name]} for name, meta in table["columns"].items()}}
            for key, table in saved["tables"].items()}}
        if compare_outputs(read_outputs(output, now), expected):
            raise ValueError("The written workbook differs from the validated in-memory values.")
    except ValueError as error:
        return stop(events, directory, {**base, "status": "stopped", "reason": str(error)})
    (Path(directory) / "outputs.json").write_text(json.dumps({k: [plain(r) for r in o["rows"]] for k, o in outputs.items()}, indent=2))
    if given:
        # Answers about this input's data (a missing value, an unmatched payment) belong to the run, not the workflow.
        (Path(directory) / "answers.json").write_text(json.dumps(given, indent=2))
    events.emit("run_end", status="success", output=str(output), binding_changes=base["binding_changes"])
    return {**base, "status": "success", "output": str(output),
            "rows": {key: len(o["rows"]) for key, o in outputs.items()}}


def stop(events, directory, result):
    if result.get("questions"):
        (Path(directory) / "questions.json").write_text(json.dumps(result["questions"], indent=2, default=str))
    events.emit("run_end", **{k: v for k, v in result.items() if k != "bindings"})
    return result
