"""Discover and verify full reconciliation plans from local example evidence."""

from datetime import date, datetime
import json
from pathlib import Path

from .events import Events
from .planner import (CHECKS, FIELD_DOCS, FIELDS, SYSTEM, accepted_questions, execute_plan, proposal_schema,
                      required_fields, validate_plan)
from .sources import money, prepare_sources, read_workbook
from .workbook import business_tables, compare_tables, fund_values, observed_labels


def model_evidence(workbook, target, source, context):
    fund = fund_values(workbook, source)
    target_source = read_workbook(target)
    target_tables = business_tables(target, target_source)
    if not target_tables["journal"]:
        raise ValueError("Target contains no journal example; supply a completed reference workbook.")
    replacements = {}
    for row in source["roster"]:
        replacements[row["name"]] = "PERSON_" + row["id"]
    for row in source["tracker"]:
        if row["receipt_date"]:
            replacements[row["receipt_date"]] = "DATE_" + row["id"]
    for key, value in fund.items():
        if isinstance(value, str) and value:
            replacements[value] = "CONST_" + key
    types = {r["Transaction Type"] for r in target_tables["journal"]}
    relevant_accounts = [row for row in source["accounts"] if row["transaction_type"] in types]
    if types - {r["transaction_type"] for r in relevant_accounts}:
        raise ValueError("Resolve target transaction types missing from the current source chart before planning.")
    if {r["Account"] for r in target_tables["journal"]} - {r["account"] for r in relevant_accounts}:
        raise ValueError("Resolve target account codes inconsistent with the current source chart before planning.")
    for row in relevant_accounts:
        if row["account"] not in replacements:
            replacements[row["account"]] = "ACCOUNT_" + str(len(replacements))

    def mask(value):
        if isinstance(value, datetime):
            value = value.date().isoformat()
        elif isinstance(value, date):
            value = value.isoformat()
        if isinstance(value, str):
            if value in replacements:
                return replacements[value]
            for original, surrogate in sorted(replacements.items(), key=lambda item: -len(item[0])):
                if len(original) >= 3:
                    value = value.replace(original, surrogate)
            return value
        if isinstance(value, dict):
            return {key: mask(item) for key, item in value.items()}
        if isinstance(value, list):
            return [mask(item) for item in value]
        return value

    example_rows, included_unpaid = [], False
    for offset in range(0, len(target_tables["journal"]), 2):
        pair = target_tables["journal"][offset:offset + 2]
        paid = any(money(row["Amount (Local)"]) != "0.00" for row in pair)
        if paid or not included_unpaid:
            example_rows.extend(pair)
        if not paid:
            included_unpaid = True
    evidence = {"context": context, "allowed_fields": sorted(FIELDS),
                "tool_fields": FIELD_DOCS, "required_checks": sorted(CHECKS), "numeric_inputs": ["called", "received"],
                "numeric_outputs": ["outstanding"], "fund": fund,
                "input_tracker": [{"id": row["id"], "investor": row["name"], **{k: row[k] for k in ["called", "received", "receipt_date"]}} for row in source["tracker"]],
                "receipt_evidence": [{"amount": row["amount"], "date": row["date"],
                                      "investor_id": row["investor_id"], "match_evidence": row["match"]["method"]} for row in source["transactions"]],
                "accounts": [{k: row[k] for k in ["transaction_type", "account"]} for row in relevant_accounts],
                "statement_currency": source["statement_currency"], "tracker_currency": source["tracker_currency"],
                "target": {"journal": {"row_count": len(target_tables["journal"]),
                                      "example_rows": example_rows,
                                      "index_sequence": [{key: row[key] for key in ["JE Index", "Trans Index", "Transaction GL Reference"]} for row in target_tables["journal"]]},
                           "tracker": target_tables["tracker"],
                           "staging": [{k: row[k] for k in ["date", "deposit", "investor"]} for row in target_tables["staging"]]}}
    return mask(evidence), mask, target_tables


MAX_ATTEMPTS = 5


def matching_fields(plan, differences):
    """Journal field expressions that already reproduce the target on every row."""
    if any(d["table"] == "journal" and d.get("check") == "row_count" for d in differences):
        return {}
    wrong = {d.get("field") for d in differences if d["table"] == "journal"}
    return {name: expression for name, expression in plan["journal"]["fields"].items() if name not in wrong}


def summarise(differences, plan):
    """One entry per wrong field: target vs candidate examples and the expression that produced them."""
    groups = {}
    for d in differences:
        key = (d["table"], d.get("field") or d.get("check"))
        group = groups.setdefault(key, {"table": key[0], "field": key[1], "rows_wrong": 0, "examples": []})
        group["rows_wrong"] += 1
        example = {"expected": d.get("expected"), "actual": d.get("actual")}
        if d["table"] == "journal" and "row" in d:
            example["line"] = 2 - d["row"] % 2
        if example not in group["examples"] and len(group["examples"]) < 3:
            group["examples"].append(example)
    for group in groups.values():
        if group["table"] == "journal" and plan and group["field"] in plan["journal"]["fields"]:
            group["your_expression"] = plan["journal"]["fields"][group["field"]]
    return list(groups.values())


def build_workbook(workbook, target, bank, context, answers, client, events, max_attempts=MAX_ATTEMPTS):
    if type(max_attempts) is not int or not 1 <= max_attempts <= MAX_ATTEMPTS:
        raise ValueError(f"Full build budget must be between one and {MAX_ATTEMPTS} calls.")
    if Path(context).suffix.lower() == ".pdf":
        raise ValueError("Provide the customer context as a plain-text file, not a PDF.")
    events.emit("run_start", phase="build_workbook", model=client.model, workflow_version=2,
                max_attempts=max_attempts, max_tokens_per_call=client.max_tokens,
                prompt_version=client.prompt_version)
    try:
        child = Events(events.directory / "inspection")
        inspection = prepare_sources(workbook, bank, None, child)
        source = json.loads(Path(inspection["source_evidence"]).read_text())
        blockers = accepted_questions(source, answers)
        if blockers:
            events.emit("inquiry", questions=blockers)
            (events.directory / "questions.json").write_text(json.dumps(blockers, indent=2))
            events.emit("run_end", status="needs_input", reason="Source policies require explicit answers.")
            return {"status": "needs_input", "questions": blockers, "question_file": str(events.directory / "questions.json")}
        if not source["tracker"]:
            raise ValueError("No complete source records are available for planning.")
        for code, answer in answers.items():
            events.emit("accepted_answer", code=code, answer=answer, origin="explicit_answers_file")
        profile, mask, target_tables = model_evidence(workbook, target, source, Path(context).read_text())
        profile["accepted_answers"] = answers
        required = required_fields(target_tables["journal"])
        client.schema = proposal_schema(sorted({row["transaction_type"] for row in profile["accounts"]}))
        messages = [{"role": "system", "content": SYSTEM}, {"role": "user", "content": json.dumps(profile)}]
        locked = {}
        for attempt in range(1, max_attempts + 1):
            content = client.complete(messages, events, attempt)
            messages.append({"role": "assistant", "content": content})
            candidate = events.directory / f"candidate-{attempt}"
            candidate.mkdir()
            plan = None
            try:
                reply = json.loads(content)
                if not isinstance(reply, dict):
                    raise ValueError("Expected a ready workflow or a specific question.")
                for optional in ["workflow", "question"]:
                    if optional in reply and reply[optional] is None:
                        del reply[optional]
                if reply.get("status") == "needs_input":
                    if set(reply) != {"status", "question"} or not isinstance(reply["question"], str) or not reply["question"].strip():
                        raise ValueError("Inquiry requires a nonempty question.")
                    events.emit("inquiry", attempt=attempt, question=reply["question"])
                    events.emit("run_end", status="needs_input", reason=reply["question"])
                    return {"status": "needs_input", "question": reply["question"]}
                if set(reply) != {"status", "workflow"} or reply["status"] != "ready":
                    raise ValueError("Expected exact ready/workflow fields.")
                plan = validate_plan(reply["workflow"])
                # Fields already proven against the target cannot regress in a later proposal.
                restored = [name for name, expression in locked.items() if plan["journal"]["fields"][name] != expression]
                plan["journal"]["fields"].update(locked)
                if restored:
                    events.emit("locked_fields", attempt=attempt, restored=restored)
                (candidate / "plan.json").write_text(json.dumps(plan, indent=2))
                events.emit("candidate", attempt=attempt, candidate_version=attempt, plan=plan)
                result = execute_plan(plan, workbook, bank, candidate, answers, events, required)
                if result["status"] == "needs_input":
                    feedback = {"source_questions": result["questions"]}
                else:
                    written = business_tables(result["output"], result["evidence"])
                    differences = compare_tables(written, target_tables)
                    (candidate / "differences.json").write_text(json.dumps(differences, indent=2))
                    events.emit("comparison", attempt=attempt, status="fail" if differences else "pass", differences=differences)
                    if not differences:
                        saved = {"schema_version": 2, "plan": plan, "accepted_answers": answers,
                                 "checks": {"required_fields": required},
                                 "bindings": result["evidence"]["bindings"],
                                 "headers": observed_labels(workbook, result["evidence"]),
                                 "workflow_version": 1}
                        path = events.directory / "workflow.json"
                        with path.open("x", encoding="utf-8") as stream:
                            json.dump(saved, stream, indent=2)
                        events.emit("run_end", status="success", attempts=attempt,
                                    workflow="workflow.json", output=result["output"],
                                    fidelity="business values; helper columns/formula text not cloned")
                        return {"status": "success", "attempts": attempt, "workflow": str(path),
                                "output": result["output"], "journal_rows": len(written["journal"]),
                                "fidelity": "business values and exact DIU headers; helper columns/formula text not cloned"}
                    feedback = {"differences": summarise(differences, plan), "total_differences": len(differences)}
                    locked = matching_fields(plan, differences)
            except (ValueError, TypeError, KeyError, ArithmeticError) as error:
                feedback = {"validation_error": str(error)}
                events.emit("validation", attempt=attempt, status="fail", **feedback)
                journal_file = candidate / "journal.json"
                if journal_file.exists():
                    actual = {"journal": json.loads(journal_file.read_text()), "tracker": [], "staging": []}
                    expected = {"journal": target_tables["journal"], "tracker": [], "staging": []}
                    differences = compare_tables(actual, expected)
                    (candidate / "differences.json").write_text(json.dumps(differences, indent=2))
                    feedback.update(differences=summarise(differences, plan), total_differences=len(differences))
                    locked = matching_fields(plan, differences)
                    events.emit("comparison", attempt=attempt, status="fail", differences=differences, output_written=False)
            if locked:
                feedback["locked_fields"] = sorted(locked)
            messages.append({"role": "user", "content": json.dumps({
                "instruction": "The candidate failed. Revise it to resolve every reported issue and return a corrected complete workflow. "
                               "Fields in locked_fields already match the target and are kept as they are; change only the others. "
                               "Do not repeat the failed candidate.",
                **mask(feedback)})})
        events.emit("run_end", status="budget_exhausted", reason="No passing workbook within call budget.")
        return {"status": "budget_exhausted"}
    except Exception as error:
        events.emit("run_end", status="error", reason=str(error))
        raise
