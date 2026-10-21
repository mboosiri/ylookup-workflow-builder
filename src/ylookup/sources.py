"""Local source evidence for reconciliation; no target access or posting approval."""

from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
import hashlib
import json
from pathlib import Path
import re
import unicodedata

from openpyxl import load_workbook
from pypdf import PdfReader

from .workflow import execute, number, validate_workflow


ALIASES = {
    "id": {"id", "investor id", "partner id", "lp id"},
    "investor": {"investor", "investor name", "lp name", "partner name"},
    "called": {"called", "amount called", "capital call amount", "total call", "amount due"},
    "received": {"received", "amount received", "cash collected", "money received"},
    "outstanding": {"outstanding", "amount outstanding", "remaining", "balance due"},
    "transaction_type": {"trans type", "transaction type"},
    "account": {"gl account code", "account code"},
    "entity": {"legal entity", "fund name"},
    "bank_account": {"bank account", "bank account name"},
    "batch_type": {"batch type"},
}
ROLES = {"roster": ("id", "investor", "called"),
         "accounts": ("transaction_type", "account"),
         "constants": ("entity", "bank_account", "batch_type")}


def normalise(value):
    text = unicodedata.normalize("NFKD", str(value))
    text = "".join(c for c in text if not unicodedata.combining(c))
    return " ".join(re.findall(r"[a-z0-9]+", text.lower()))


def meaning(value):
    label = normalise(value)
    if re.fullmatch(r"call(?: \d+)? total call", label):
        return "called"
    matches = [key for key, names in ALIASES.items() if label in names]
    return matches[0] if len(matches) == 1 else None


def money(value):
    return str(number(value).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def question(code, detail, **evidence):
    return {"code": code, "question": detail, **evidence}


def read_workbook(path):
    book = load_workbook(path, read_only=True, data_only=True)
    raw = load_workbook(path, read_only=True, data_only=False)
    try:
        candidates = {role: [] for role in ROLES}
        for sheet in book:
            for index, row in enumerate(sheet.iter_rows(values_only=True), 1):
                columns = {}
                for column, value in enumerate(row, 1):
                    role = meaning(value) if value is not None else None
                    if role:
                        columns.setdefault(role, []).append(column)
                for role, required in ROLES.items():
                    if index < sheet.max_row and all(field in columns for field in required):
                        candidates[role].append({"sheet": sheet.title, "header_row": index,
                                                 "columns": {k: columns[k] for k in required}})
        bindings, issues = {}, []
        for role, choices in candidates.items():
            if len(choices) != 1 or any(len(v) != 1 for choice in choices for v in choice["columns"].values()):
                issues.append(question("binding", f"Which table/columns supply the {role} role?",
                                       required_fields=list(ROLES[role]), candidates=choices))
            else:
                selected = choices[0]
                bindings[role] = {**selected, "columns": {k: v[0] for k, v in selected["columns"].items()}}
        result = {"bindings": bindings, "questions": issues, "roster": [], "accounts": [],
                  "constants": [], "currency": None}
        for role, binding in bindings.items():
            sheet = book[binding["sheet"]]
            columns = binding["columns"]
            started = False
            for row_index, row in enumerate(sheet.iter_rows(min_row=binding["header_row"] + 1, values_only=True), binding["header_row"] + 1):
                values = {key: row[col - 1] if col <= len(row) else None for key, col in columns.items()}
                provenance = {"sheet": sheet.title, "row": row_index}
                if role == "roster":
                    if values["id"] is None and values["investor"] is None:
                        if not started and isinstance(values["called"], str) and re.fullmatch(r"[A-Z]{3}", values["called"]):
                            result["currency"] = values["called"]
                        elif started:
                            formula = raw[sheet.title].cell(row_index, columns["called"]).value
                            if isinstance(formula, str) and re.match(r"=SUM\(", formula, re.I):
                                # An explicit summary ends this table, not the workbook's used range.
                                break
                            if values["called"] is not None:
                                issues.append(question("roster_identity", "Resolve an amount row without investor identity or an unlabelled total.", **provenance))
                        continue
                    started = True
                    if any(values[key] is None for key in columns):
                        formula = raw[sheet.title].cell(row_index, columns["called"]).value
                        issues.append(question("roster_value", "Resolve missing investor ID/name/called amount or unavailable formula result.",
                                               **provenance, called_formula=str(formula) if isinstance(formula, str) and formula.startswith("=") else None))
                        continue
                    if (not isinstance(values["id"], str) or not values["id"].strip() or
                            not isinstance(values["investor"], str) or not normalise(values["investor"])):
                        issues.append(question("roster_identity", "Provide text investor IDs/names without losing identifier formatting.", **provenance))
                        continue
                    try:
                        called = money(values["called"])
                    except (ValueError, ArithmeticError):
                        issues.append(question("called_amount", "Provide a finite evaluated called amount.", **provenance))
                        continue
                    result["roster"].append({"id": values["id"], "name": values["investor"],
                                               "called": called, "source": provenance})
                else:
                    if all(value is None for value in values.values()):
                        continue
                    if any(value is None for value in values.values()):
                        issues.append(question("lookup_value", f"Resolve missing values in the {role} table.", **provenance))
                        continue
                    result[role].append({**{k: str(v) for k, v in values.items()}, "source": provenance})
        ids = [row["id"] for row in result["roster"]]
        if not ids or len(ids) != len(set(ids)):
            issues.append(question("roster_ids", "Resolve an empty roster or duplicate investor IDs."))
        if not result["currency"]:
            issues.append(question("tracker_currency", "Which currency applies to the called amounts?"))
        return result
    finally:
        book.close()
        raw.close()


def parse_statement(pages):
    transactions, issues, currencies = [], [], set()
    for page_number, text in pages:
        currencies.update(re.findall(r"\bCurrency\s*:\s*([A-Z]{3})\b", text))
        positions, current = None, None
        for line_number, line in enumerate(text.splitlines(), 1):
            labels = {label: re.search(r"\b" + label + r"\b", line, re.I)
                      for label in ["date", "description", "withdrawal", "deposit", "balance"]}
            if all(labels.values()):
                positions = {k: (v.start() + v.end()) / 2 for k, v in labels.items()}
                if list(positions.values()) != sorted(positions.values()):
                    issues.append(question("bank_layout", "Provide a supported ordered transaction table.", page=page_number))
                    positions = None
                current = None
                continue
            if positions is None:
                continue
            date_match = re.match(r"^\s*(\d{1,2} [A-Za-z]{3} \d{4})\b", line)
            if date_match:
                current = None
                try:
                    date = datetime.strptime(date_match.group(1), "%d %b %Y").date().isoformat()
                except ValueError:
                    issues.append(question("receipt_date", "Resolve the transaction date.", page=page_number, line=line_number))
                    continue
                boundary = (positions["description"] + positions["withdrawal"]) / 2
                amounts = {key: [] for key in ["withdrawal", "deposit", "balance"]}
                for match in re.finditer(r"(?<![\w.])-?\d[\d,]*\.\d{2}(?![\w.])", line):
                    centre = (match.start() + match.end()) / 2
                    if centre >= boundary:
                        nearest = min(amounts, key=lambda key: abs(positions[key] - centre))
                        amounts[nearest].append(match.group())
                if any(len(values) > 1 for values in amounts.values()) or len(amounts["balance"]) != 1 or sum(bool(amounts[k]) for k in ["withdrawal", "deposit"]) != 1:
                    issues.append(question("receipt_amount", "Resolve deposit/withdrawal/balance columns; no amount will be guessed.", page=page_number, line=line_number))
                    continue
                kind = "deposit" if amounts["deposit"] else "withdrawal"
                amount = money(amounts[kind][0].replace(",", ""))
                if number(amount) < 0:
                    issues.append(question("receipt_sign", "Resolve a negative transaction amount.", page=page_number, line=line_number))
                    continue
                current = {"date": date, "kind": kind, "amount": amount,
                           "balance": money(amounts["balance"][0].replace(",", "")),
                           "description": line[date_match.end():int(boundary)].strip(),
                           "source": {"page": page_number, "line": line_number}}
                transactions.append(current)
            elif line.strip():
                boundary = (positions["description"] + positions["withdrawal"]) / 2
                monetary_text = line[int(boundary):]
                if re.search(r"(?<![\w.])-?\d[\d,]*\.\d{2}(?![\w.])", monetary_text):
                    issues.append(question("receipt_date", "Resolve an undated or unsupported transaction row; monetary values will not be attached to another receipt.", page=page_number, line=line_number))
                    current = None
                    continue
                if re.search(r"Thank you for|Generated on", line, re.I):
                    current = None
                elif current:
                    current["description"] += " " + line.strip()
        if positions is None:
            issues.append(question("bank_layout", "Provide a readable text-layer statement with Date/Description/Withdrawal/Deposit/Balance headers.", page=page_number))
    for transaction in transactions:
        reference_currencies = set(re.findall(r"\b([A-Z]{3})\|\d+(?:\.\d+)?\|", transaction["description"]))
        if reference_currencies and (len(currencies) != 1 or reference_currencies != currencies):
            issues.append(question("receipt_currency", "Resolve conflicting currency evidence in the transaction reference.",
                                   source=transaction["source"], reference_currencies=sorted(reference_currencies)))
    if len(currencies) != 1:
        issues.append(question("statement_currency", "Resolve the statement currency.", candidates=sorted(currencies)))
    if not transactions:
        issues.append(question("bank_records", "No supported transaction records were extracted; confirm whether this is an empty or unsupported statement."))
    return {"transactions": transactions, "currency": next(iter(currencies)) if len(currencies) == 1 else None,
            "questions": issues}


def read_statement(path):
    # shortcut: aligned text-layer tables only; add an OCR/layout adapter for other statements.
    reader = PdfReader(path)
    return parse_statement([(index, page.extract_text(extraction_mode="layout"))
                            for index, page in enumerate(reader.pages, 1)])


def match_investor(description, roster, allow_partial=True):
    text = normalise(description)
    full = [row for row in roster if re.search(r"\b" + re.escape(normalise(row["name"])) + r"\b", text)]
    if full or not allow_partial:
        return full, "full_name"
    tokens = text.split()
    matches = []
    for row in roster:
        parts = normalise(row["name"]).split()
        if len(parts) >= 2 and parts[0] in tokens and any(
                token == parts[-1] or (len(token) >= 3 and parts[-1].startswith(token)) for token in tokens):
            matches.append(row)
    return matches, "first_name_and_surname_prefix"


def prepare_sources(workbook, bank, workflow, events, matching_policy="full_or_unique_prefix"):
    if matching_policy not in {"full_name_only", "full_or_unique_prefix"}:
        raise ValueError("Unsupported source matching policy.")
    if workflow is not None:
        validate_workflow(workflow)
        source_fields = {workflow[key]: meaning(workflow[key]) for key in ["left", "right"]}
        if set(source_fields.values()) != {"called", "received"} or meaning(workflow["output"]) != "outstanding":
            raise ValueError("The numeric rule must bind called/received source meanings and an outstanding output.")
    events.emit("run_start", phase="sources", workflow_version=workflow["version"] if workflow else None,
                source_sha256={str(path): hashlib.sha256(Path(path).read_bytes()).hexdigest() for path in [workbook, bank]})
    try:
        extracted = read_workbook(workbook)
        statement = read_statement(bank)
        issues = extracted["questions"] + statement["questions"]
        events.emit("source_extraction", bindings=extracted["bindings"], roster_count=len(extracted["roster"]),
                    transactions=statement["transactions"], questions=issues)
        matched, used = {}, set()
        coverage_complete = not issues
        for transaction in statement["transactions"]:
            signature = (transaction["date"], normalise(transaction["description"]), transaction["kind"], transaction["amount"])
            if signature in used:
                issues.append(question("duplicate_receipt", "Resolve a repeated transaction before counting it.", source=transaction["source"]))
                coverage_complete = False
                continue
            used.add(signature)
            if transaction["kind"] != "deposit":
                issues.append(question("withdrawal_policy", "How should this withdrawal affect capital-call reconciliation?", source=transaction["source"]))
                coverage_complete = False
                continue
            candidates, method = match_investor(transaction["description"], extracted["roster"], matching_policy == "full_or_unique_prefix")
            transaction["match"] = {"method": method, "candidates": [row["id"] for row in candidates]}
            if len(candidates) != 1:
                issues.append(question("investor_match", "Identify the investor for this receipt; no unique supported match exists.", source=transaction["source"], candidates=transaction["match"]["candidates"]))
                coverage_complete = False
                continue
            investor = candidates[0]
            transaction["investor_id"] = investor["id"]
            if investor["id"] in matched:
                issues.append(question("multiple_receipts", "Resolve the grouping/date policy for repeated investor receipts.", investor_id=investor["id"]))
                coverage_complete = False
            else:
                matched[investor["id"]] = transaction
            events.emit("receipt_match", source=transaction["source"], investor_id=investor["id"], method=method)
        tracker = []
        if coverage_complete:
            numeric_input = []
            for investor in extracted["roster"]:
                receipt = matched.get(investor["id"])
                received = receipt["amount"] if receipt else "0.00"
                row = {**investor, "received": received, "receipt_date": receipt["date"] if receipt else None}
                tracker.append(row)
                if workflow is not None:
                    numeric_input.append({"id": investor["id"], **{field: row[role] for field, role in source_fields.items()}})
            if workflow is not None:
                by_id = {row["id"]: row[workflow["output"]] for row in execute(workflow, numeric_input)}
                for row in tracker:
                    row["outstanding"] = money(by_id[row["id"]])
        if extracted["currency"] and statement["currency"] and extracted["currency"] != statement["currency"]:
            issues.append(question("currency_conflict", "Statement and tracker currencies differ; confirm a currency policy before posting.", statement_currency=statement["currency"], tracker_currency=extracted["currency"]))
        for issue in issues:
            events.emit("inquiry", **issue)
        result = {"status": "needs_input" if issues else "success", "bindings": extracted["bindings"],
                  "statement_currency": statement["currency"], "tracker_currency": extracted["currency"],
                  "roster": extracted["roster"], "transactions": statement["transactions"], "tracker": tracker,
                  "accounts": extracted["accounts"], "constants": extracted["constants"], "questions": issues}
        with (events.directory / "sources.json").open("x", encoding="utf-8") as stream:
            json.dump(result, stream, indent=2)
        events.emit("run_end", status=result["status"], source_evidence="sources.json")
        return {"status": result["status"], "source_evidence": str(events.directory / "sources.json"),
                "investors": len(result["roster"]), "transactions": len(result["transactions"]),
                "tracker_records": len(tracker), "questions": issues}
    except Exception as error:
        events.emit("run_end", status="error", reason=str(error))
        raise
