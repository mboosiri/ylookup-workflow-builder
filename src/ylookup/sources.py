"""General source tools: header-row tables in a workbook, a bank statement PDF, name matching.

No business vocabulary lives here. A table is whatever a workbook header row says it is; which
tables and columns matter is learned by the builder and stored in the workflow.
"""

from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import re
import unicodedata

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from pypdf import PdfReader


# ponytail: a sheet ends after 50 empty rows; Excel's used range is often inflated far past the data
# (the tracker reports 16,406 rows). Raise it if real files leave bigger gaps inside a table.
BLANK_RUN = 50


def normalise(value):
    text = unicodedata.normalize("NFKD", str(value))
    text = "".join(c for c in text if not unicodedata.combining(c))
    return " ".join(re.findall(r"[a-z0-9]+", text.lower()))


def label(value):
    return " ".join(str(value).split())


def number(value):
    if isinstance(value, bool) or not isinstance(value, (str, int, float, Decimal)):
        raise ValueError("Expected a finite numeric value, not missing data.")
    try:
        result = Decimal(str(value))
    except InvalidOperation as error:
        raise ValueError("Invalid numeric value.") from error
    if not result.is_finite():
        raise ValueError("Expected a finite numeric value.")
    return result


def money(value):
    return str(number(value).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def kind(value):
    if value is None or isinstance(value, str) and not value.strip():
        return None
    if isinstance(value, (datetime, date)):
        return "date"
    if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
        return "number"
    return "text"


def day(value):
    return value.date() if isinstance(value, datetime) else value


def canon(value):
    """Comparable form of a cell value, for indexes and equality."""
    k = kind(value)
    if k == "number":
        return money(value)
    if k == "date":
        return day(value).isoformat()
    return None if k is None else " ".join(str(value).split())


def equivalent(actual, expected):
    a, b = kind(actual), kind(expected)
    if a is None or b is None:
        return a is None and b is None
    if a != b or (a == "text") != (isinstance(actual, str) and isinstance(expected, str)):
        return False
    x, y = canon(actual), canon(expected)
    if x == y:
        return True
    # ponytail: long free text (bank references) is compared on letters and digits only, because PDF
    # extraction moves spaces and separators; short codes and labels stay exact.
    return a == "text" and min(len(x), len(y)) > 60 and re.sub(r"\W", "", x) == re.sub(r"\W", "", y)


def question(code, text, key, options=None, stop=None, **evidence):
    """Something only a person can settle. The key makes the id stable, so a recorded answer applies again.

    options: answers the code can act on (free text is accepted when absent).
    stop: the option meaning "do not continue"; choosing it stops the command with that reason.
    """
    return {"id": f"{code}:{key}", "code": code, "question": text, "options": options, "stop": stop, **evidence}


def answered(item, answers):
    """The recorded answer text for a question, or None. Choosing its stop option raises."""
    entry = (answers or {}).get(item["id"])
    text = entry.get("answer") if isinstance(entry, dict) else entry
    if text is not None and item.get("stop") and text == item["stop"]:
        raise ValueError(f"Stopped by a recorded answer: {item['question']} -> {text}")
    return text


def sheet_rows(path):
    """{sheet: [(row number, values, formulas)]} for non-empty rows, read in one streaming pass."""
    values = load_workbook(path, read_only=True, data_only=True)
    formulas = load_workbook(path, read_only=True, data_only=False)
    try:
        result = {}
        for sheet in values:
            rows, blank = [], 0
            pairs = zip(sheet.iter_rows(values_only=True), formulas[sheet.title].iter_rows(values_only=True))
            for index, (row, raw) in enumerate(pairs, 1):
                raw = [getattr(f, "text", f) for f in raw]
                raw = [f if isinstance(f, str) and f.startswith("=") else None for f in raw]
                if any(kind(v) for v in row) or any(raw):
                    rows.append((index, list(row), raw))
                    blank = 0
                else:
                    blank += 1
                    if blank >= BLANK_RUN:
                        break
            result[sheet.title] = rows
        return result
    finally:
        values.close()
        formulas.close()


def at(cells, column):
    return cells[column] if column < len(cells) else None


def is_header(values):
    cells = [v for v in values if kind(v)]
    texts = [normalise(v) for v in cells if isinstance(v, str)]
    return len(cells) >= 2 and len(texts) == len(cells) and all(texts) and len(set(texts)) == len(texts)


def make_table(sheet, header, block):
    """A header row plus the contiguous rows under it, split into records and the rows around them.

    The key column is the first one holding distinct text on at least a third of the rows (an ID or a name). Rows without
    a key before the first record are units rows (e.g. USD), after the last one totals; rows without a
    key between records are orphans a person must explain.
    """
    row_number, values, _ = header
    headers = {c: label(v) for c, v in enumerate(values) if kind(v)}
    key = None
    for column in sorted(headers):
        found = [at(r[1], column) for r in block if kind(at(r[1], column))]
        enough = len(found) >= 2 and 3 * len(found) >= len(block) or len(found) == len(block) == 1
        if enough and all(isinstance(v, str) for v in found) and len(set(found)) == len(found):
            key = column
            break
    filled = [r for r in block if any(kind(at(r[1], c)) or at(r[2], c) for c in headers)]
    if key is None:
        records, pre, post, orphans = filled, [], [], []
    else:
        keyed = [i for i, r in enumerate(block) if kind(at(r[1], key))]
        first, last = (keyed[0], keyed[-1]) if keyed else (len(block), len(block))
        records = [block[i] for i in keyed]
        pre, post = block[:first], block[last + 1:]
        orphans = [r for r in block[first:last + 1] if not kind(at(r[1], key)) and r in filled]
    units = {headers[c]: at(r[1], c).strip() for r in pre for c in headers
             if isinstance(at(r[1], c), str) and re.fullmatch(r"[A-Z]{3}", at(r[1], c).strip())}
    return {"sheet": sheet, "row": row_number, "headers": headers, "key": headers.get(key), "width": len(headers),
            "records": [{"row": r[0], "cells": r[1], "formulas": r[2]} for r in records],
            "all": [{"row": r[0], "cells": r[1], "formulas": r[2]} for r in filled],
            "orphans": [r[0] for r in orphans], "post": [{"row": r[0], "formulas": r[2]} for r in post],
            "units": units, "rows": len(filled)}


def find_tables(rows_by_sheet):
    tables = []
    for sheet, rows in rows_by_sheet.items():
        index = 0
        while index < len(rows):
            if not is_header(rows[index][1]):
                index += 1
                continue
            end = index + 1
            while end < len(rows) and rows[end][0] == rows[end - 1][0] + 1:
                end += 1
            tables.append(make_table(sheet, rows[index], rows[index + 1:end]))
            index = end
    return tables


def column_types(table, column):
    return sorted({kind(at(r["cells"], column)) or "blank" for r in table["records"]}) or ["blank"]


def rekey(table, columns):
    """A found table with its records keyed by the workflow's column labels instead of positions."""
    key_column = next((c for c, h in table["headers"].items() if h == table["key"]), None)
    keyed = lambda rows: [{"row": r["row"], "key": None if key_column is None else at(r["cells"], key_column),
                           "values": {label: at(r["cells"], c) for label, c in columns.items()},
                           "formulas": {label: at(r["formulas"], c) for label, c in columns.items()}} for r in rows]
    return {"sheet": table["sheet"], "row": table["row"], "columns": columns, "rows": table["rows"],
            "units": {label: table["units"].get(table["headers"][c]) for label, c in columns.items()},
            "orphans": table["orphans"], "post": table["post"], "headers": table["headers"],
            "records": keyed(table["records"]), "all": keyed(table["all"])}


def table_score(want, table):
    labels = {normalise(h) for h in table["headers"].values()}
    return sum(normalise(h) in labels for h in want["headers"]), table["sheet"] == want["sheet"]


def describe(table):
    return f"{table['sheet']}!row {table['row']}: " + ", ".join(list(table["headers"].values())[:8])


def bind_tables(spec, found, answers=None):
    """Find each workflow table in the current workbook by its header fingerprint, then its columns.

    A table binds when at least half its build-time headers are still there, or its sheet name is.
    Best matches are assigned first, so a strong match is never stolen by a weak one. Returns
    (bound tables, problems); a problem names what is missing and the candidates a person can choose.
    """
    pairs = sorted(((table_score(want, table), key, i) for key, want in spec["tables"].items()
                    for i, table in enumerate(found)), reverse=True)
    chosen, claimed = {}, set()
    for (hits, same_sheet), key, i in pairs:
        if key in chosen or i in claimed:
            continue
        if 2 * hits >= len(spec["tables"][key]["headers"]) or same_sheet:
            ties = [j for (h, s), k, j in pairs if k == key and (h, s) == (hits, same_sheet) and j not in claimed]
            if len(ties) == 1:
                chosen[key] = i
                claimed.add(i)
    bound, problems = {}, []
    for key, want in spec["tables"].items():
        if key not in chosen:
            item = question("table_choice", f"Which table is now the '{key}' table (previously {want['sheet']!r} "
                            f"with headers {want['headers'][:6]})?", key,
                            options=[describe(t) for i, t in enumerate(found) if i not in claimed], table=key)
            reply = answered(item, answers)
            index = next((i for i, t in enumerate(found) if describe(t) == reply and i not in claimed), None)
            if index is None:
                problems.append(item)
                continue
            chosen[key] = index
            claimed.add(index)
        table = found[chosen[key]]
        current = {normalise(h): c for c, h in table["headers"].items()}
        known = {normalise(h) for h in want["headers"]} | {normalise(m["header"]) for m in want["columns"].values()}
        columns = {name: current[normalise(m["header"])] for name, m in want["columns"].items() if normalise(m["header"]) in current}
        for name, meta in want["columns"].items():
            if name in columns:
                continue
            item = question(
                "column_mapping", f"Which column of {table['sheet']} now holds '{name}' (previously '{meta['header']}')?",
                f"{key}.{name}", table=key, column=name, previous_header=meta["header"], types=meta["types"],
                where=f"{table['sheet']}!row {table['row']}",
                options=[h for c, h in table["headers"].items() if normalise(h) not in known and c not in columns.values()],
                candidate_types={h: column_types(table, c) for c, h in table["headers"].items() if normalise(h) not in known})
            reply = answered(item, answers)
            if reply is not None:
                entry = answers[item["id"]]
                problem = (f"'{reply}' is already used for another column" if current.get(normalise(reply)) in columns.values()
                           else mapping_problem(spec, key, name, reply, item, isinstance(entry, str) or entry.get("by") != "model"))
                if problem is None:
                    columns[name] = current[normalise(reply)]
                    continue
                item = {**item, "question": item["question"] + f" The answer '{reply}' was rejected: {problem}"}
            problems.append(item)
        bound[key] = rekey(table, columns)
    return bound, problems


def mapping_problem(spec, key, name, header, item, by_person):
    """Why a proposed column for a renamed header cannot be right, or None.

    A column the workflow writes must be empty in the input, and when several empty columns fit, the
    values give no evidence, so only a person can choose. A column it reads must hold the value types
    it held at build time.
    """
    if header not in (item["options"] or []):
        return f"'{header}' is not one of the unrecognised headers {item['options']}"
    types = set(item["candidate_types"][header]) - {"blank"}
    if name in spec["outputs"].get(key, {}).get("columns", []):
        if types:
            return f"'{header}' already holds {sorted(types)} values; the workflow writes {key}.{name} and would overwrite them"
        empty = [h for h in item["options"] if not set(item["candidate_types"][h]) - {"blank"}]
        if len(empty) > 1 and not by_person:
            return f"{len(empty)} empty columns could hold {key}.{name} ({', '.join(empty)}); the values give no evidence, so a person must choose"
        return None
    before = set(item["types"]) - {"blank"}
    if types and before and not types <= before:
        return f"'{header}' holds {sorted(types)} values but {key}.{name} held {sorted(before)}"
    return None


def binding_changes(spec, bound):
    changes = []
    for key, want in spec["tables"].items():
        table = bound.get(key)
        if table is None:
            continue
        if table["sheet"] != want["sheet"]:
            changes.append(f"{key}: tab '{want['sheet']}' -> '{table['sheet']}'")
        for name, column in table["columns"].items():
            before = want["columns"][name].get("position")
            if before and before != get_column_letter(column + 1):
                changes.append(f"{key}.{name}: column {before} -> {get_column_letter(column + 1)}")
    return changes


def parse_statement(pages):
    transactions, issues, currencies = [], [], set()
    for page_number, text in pages:
        currencies.update(re.findall(r"\bCurrency\s*:\s*([A-Z]{3})\b", text))
        positions, current = None, None
        for line_number, line in enumerate(text.splitlines(), 1):
            where = f"page {page_number} line {line_number}"
            labels = {name: re.search(r"\b" + name + r"\b", line, re.I)
                      for name in ["date", "description", "withdrawal", "deposit", "balance"]}
            if all(labels.values()):
                positions = {k: (v.start() + v.end()) / 2 for k, v in labels.items()}
                if list(positions.values()) != sorted(positions.values()):
                    issues.append(question("bank_layout", "Provide a supported ordered transaction table.", f"page {page_number}"))
                    positions = None
                current = None
                continue
            if positions is None:
                continue
            date_match = re.match(r"^\s*(\d{1,2} [A-Za-z]{3} \d{4})\b", line)
            if date_match:
                current = None
                try:
                    when = datetime.strptime(date_match.group(1), "%d %b %Y").date().isoformat()
                except ValueError:
                    issues.append(question("receipt_date", "Resolve the transaction date.", where))
                    continue
                boundary = (positions["description"] + positions["withdrawal"]) / 2
                amounts = {key: [] for key in ["withdrawal", "deposit", "balance"]}
                for match in re.finditer(r"(?<![\w.])-?\d[\d,]*\.\d{2}(?![\w.])", line):
                    centre = (match.start() + match.end()) / 2
                    if centre >= boundary:
                        nearest = min(amounts, key=lambda key: abs(positions[key] - centre))
                        amounts[nearest].append(match.group())
                if any(len(values) > 1 for values in amounts.values()) or len(amounts["balance"]) != 1 or sum(bool(amounts[k]) for k in ["withdrawal", "deposit"]) != 1:
                    issues.append(question("receipt_amount", "Resolve deposit/withdrawal/balance columns; no amount will be guessed.", where))
                    continue
                kind_ = "deposit" if amounts["deposit"] else "withdrawal"
                amount = money(amounts[kind_][0].replace(",", ""))
                if number(amount) < 0:
                    issues.append(question("receipt_sign", "Resolve a negative transaction amount.", where))
                    continue
                current = {"date": when, "kind": kind_, "amount": amount,
                           "balance": money(amounts["balance"][0].replace(",", "")),
                           "description": line[date_match.end():int(boundary)].strip(),
                           "source": {"page": page_number, "line": line_number}}
                transactions.append(current)
            elif line.strip():
                boundary = (positions["description"] + positions["withdrawal"]) / 2
                if re.search(r"(?<![\w.])-?\d[\d,]*\.\d{2}(?![\w.])", line[int(boundary):]):
                    issues.append(question("receipt_date", "Resolve an undated or unsupported transaction row; monetary values will not be attached to another receipt.", where))
                    current = None
                    continue
                if re.search(r"Thank you for|Generated on", line, re.I):
                    current = None
                elif current:
                    current["description"] += " " + line.strip()
        if positions is None:
            issues.append(question("bank_layout", "Provide a readable text-layer statement with Date/Description/Withdrawal/Deposit/Balance headers.", f"page {page_number}"))
    for transaction in transactions:
        reference_currencies = set(re.findall(r"\b([A-Z]{3})\|\d+(?:\.\d+)?\|", transaction["description"]))
        if reference_currencies and (len(currencies) != 1 or reference_currencies != currencies):
            issues.append(question("receipt_currency", "Resolve conflicting currency evidence in the transaction reference.",
                                   f"{transaction['date']} {transaction['amount']}", reference_currencies=sorted(reference_currencies)))
    if len(currencies) != 1:
        issues.append(question("statement_currency", "Which currency is this statement in?", "statement",
                               options=sorted(currencies) or None))
    if not transactions:
        issues.append(question("bank_records", "No supported transaction records were extracted; confirm whether this is an empty or unsupported statement.", "statement"))
    return {"transactions": transactions, "currency": next(iter(currencies)) if len(currencies) == 1 else None,
            "questions": issues}


def read_statement(path):
    # ponytail: aligned text-layer tables only; add an OCR/layout adapter for other statements.
    reader = PdfReader(path)
    return parse_statement([(index, page.extract_text(extraction_mode="layout"))
                            for index, page in enumerate(reader.pages, 1)])


def match_names(description, names, policy="full_or_unique_prefix"):
    """Indexes of names found in a bank reference: full names, or first name plus a surname prefix."""
    text = normalise(description)
    full = [i for i, name in names.items() if re.search(r"\b" + re.escape(normalise(name)) + r"\b", text)]
    if full or policy == "full_name_only":
        return full, "full_name"
    tokens = text.split()
    partial = []
    for i, name in names.items():
        parts = normalise(name).split()
        if len(parts) >= 2 and parts[0] in tokens and any(
                token == parts[-1] or (len(token) >= 3 and parts[-1].startswith(token)) for token in tokens):
            partial.append(i)
    return partial, "first_name_and_surname_prefix"


LEAVE = "Leave it unmatched for follow-up"


def match_receipts(statement, table, name, policy, answers):
    """Assign each deposit to one record by its name column. Doubt becomes a question, never a guess.

    Returns per-record receipts {record index: {amount, date, description}}, the kept deposits with
    their record (or None when a person chose to leave one unmatched), and the open questions.
    """
    names = {i: r["values"][name] for i, r in enumerate(table["records"])
             if isinstance(r["values"].get(name), str) and normalise(r["values"][name])}
    display = {i: n if list(names.values()).count(n) == 1 else f"{n} (row {table['records'][i]['row']})"
               for i, n in names.items()}
    deposits, issues, seen = [], [], set()
    for t in statement["transactions"]:
        key = f"{t['date']} {t['kind']} {t['amount']}"
        signature = (t["date"], normalise(t["description"]), t["kind"], t["amount"])
        if signature in seen:
            item = question("duplicate_transaction", f"The statement repeats a {t['kind']} of {t['amount']} on {t['date']}. Count it twice?",
                            key, options=["Count it twice", "Count it once"], source=t["source"])
            if answered(item, answers) is None:
                issues.append(item)
            if answered(item, answers) != "Count it twice":
                continue
        seen.add(signature)
        if t["kind"] != "deposit":
            item = question("withdrawal", f"The statement has a withdrawal of {t['amount']} on {t['date']}. How should it be treated?",
                            key, options=["Leave it out of this reconciliation", "Stop for review"],
                            stop="Stop for review", source=t["source"])
            if answered(item, answers) is None:
                issues.append(item)
            continue
        found, method = match_names(t["description"], names, policy)
        record = found[0] if len(found) == 1 else None
        if record is None:
            item = question("receipt_match", f"Which {name} paid {t['amount']} on {t['date']}? "
                            f"{'No' if not found else len(found)} {'name matches' if not found else 'names match'} the bank reference.",
                            key, options=[display[i] for i in (found or names)] + [LEAVE],
                            source=t["source"], method=method)
            reply = answered(item, answers)
            if reply is None:
                issues.append(item)
                continue
            if reply != LEAVE:
                record = next((i for i, n in display.items() if n == reply), None)
                if record is None:
                    issues.append({**item, "question": item["question"] + f" (the answer '{reply}' is not one of the options)"})
                    continue
        deposits.append({**t, "record": record})
    receipts = {}
    for t in deposits:
        if t["record"] is None:
            continue
        entry = receipts.setdefault(t["record"], {"amount": Decimal(0), "dates": [], "description": []})
        entry["amount"] += number(t["amount"])
        entry["dates"].append(t["date"])
        entry["description"].append(t["description"])
    for index, entry in receipts.items():
        dates = sorted(set(entry.pop("dates")))
        entry["description"] = " ".join(entry["description"])
        if len(dates) > 1:
            item = question("receipt_date", f"{display[index]} paid on {len(dates)} dates. Which date should the entry use?",
                            f"{display[index]} {' '.join(dates)}", options=dates)
            reply = answered(item, answers)
            if reply not in dates:
                issues.append(item)
                reply = dates[0]
            dates = [reply]
        entry["date"] = date.fromisoformat(dates[0])
    return {"receipts": receipts, "deposits": deposits, "questions": issues}
