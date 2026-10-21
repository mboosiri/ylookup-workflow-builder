"""Strict business tables and safe workbook output with unchanged formula caches."""

from contextlib import contextmanager
from datetime import date, datetime
from decimal import Decimal
import os
from pathlib import Path
import posixpath
import tempfile
import xml.etree.ElementTree as ET
import zipfile

from openpyxl import load_workbook

from .sources import ALIASES, money, normalise


HEADERS = ["Legal Entity", "Legal Entity Domain", "Deal", "Position", "Batch Index",
           "Batch Type", "JE Index", "Trans Index", "GL Date", "Effective Date",
           "Transaction Currency", "Allocation Rule", "Transaction Type", "Account",
           "Is Debit", "Amount (Local)", "Amount (LE)", "Transaction Comments",
           "Batch Comments", "Bank Account", "Investor", "Investor Domain", "Vehicle",
           "Investor Amount (LE)", "Investor Amount (Local)", "Transaction GL Reference"]
STAGING = {"date": {"date", "transaction date"},
           "description": {"line item description", "description", "bank reference"},
           "withdrawal": {"withdrawal", "debit"}, "deposit": {"deposit", "credit"},
           "investor": {"matched investor", "matched investor name"}}
TRACKER = {"received": {"amount received", "cash collected", "money received"},
           "outstanding": {"amount outstanding", "outstanding", "balance due"},
           "receipt_date": {"date as per bank statement", "receipt date", "payment date"}}
FUND = {"entity": {"legal entity", "fund name"}, "entity_domain": {"legal entity domain"},
        "deal": {"deal"}, "position": {"position"}, "batch_type": {"batch type"},
        "bank_account": {"bank account", "bank account name"}}


# Workflow header names are "table.field"; each maps onto the alias sets that recognise it.
TABLES = {"roster": ALIASES, "accounts": ALIASES, "constants": ALIASES,
          "tracker": TRACKER, "fund": FUND, "staging": STAGING}


def alias_sets(name):
    table, field = name.split(".", 1)
    sets = [TABLES[table][field]]
    if table in {"tracker", "fund"} and field in ALIASES:
        sets.append(ALIASES[field])
    return sets


@contextmanager
def learned_headers(headers):
    """Recognise the header names a workflow has learned, for the duration of one run."""
    # mutates module-level alias sets, so one run per process at a time; pass aliases explicitly if runs go concurrent.
    added = []
    for name, label in headers.items():
        if not isinstance(label, str) or not normalise(label):
            continue
        for names in alias_sets(name):
            if normalise(label) not in names:
                names.add(normalise(label))
                added.append((names, normalise(label)))
    try:
        yield
    finally:
        for names, label in added:
            names.discard(label)


def observed_labels(path, evidence):
    """Header text seen for each bound field at build time, stored in the workflow."""
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        labels = {}
        for role, binding in evidence["bindings"].items():
            values = next(book[binding["sheet"]].iter_rows(min_row=binding["header_row"], max_row=binding["header_row"], values_only=True))
            for field, column in binding["columns"].items():
                labels[f"{role}.{field}"] = values[column - 1]
            extra = {"roster": ("tracker", TRACKER), "constants": ("fund", FUND)}.get(role)
            if extra:
                for field, aliases in extra[1].items():
                    found = [v for v in values if v is not None and normalise(v) in aliases]
                    if len(found) == 1:
                        labels[f"{extra[0]}.{field}"] = found[0]
        sheet, row, mapping = find_table(book, meanings=STAGING)
        values = next(sheet.iter_rows(min_row=row, max_row=row, values_only=True))
        labels.update({f"staging.{field}": values[column - 1] for field, column in mapping.items()})
        return labels
    finally:
        book.close()


def columns(values, meanings):
    result = {}
    for role, aliases in meanings.items():
        choices = [i for i, value in enumerate(values, 1) if value is not None and normalise(value) in aliases]
        if len(choices) != 1:
            raise ValueError(f"Cannot find exactly one '{role}' column (found {len(choices)}); "
                             f"recognised headers: {sorted(aliases)}.")
        result[role] = choices[0]
    return result


def find_table(book, meanings=None, exact=None):
    candidates = []
    for sheet in book:
        for index, row in enumerate(sheet.iter_rows(values_only=True), 1):
            if exact is not None:
                if list(row[:len(exact)]) == exact and not any(v is not None for v in row[len(exact):]):
                    candidates.append((sheet, index, {name: i for i, name in enumerate(exact, 1)}))
            else:
                try:
                    mapped = columns(row, meanings)
                except ValueError:
                    continue
                candidates.append((sheet, index, mapped))
    if len(candidates) != 1:
        raise ValueError("Missing or competing output table; a unique supported header is required.")
    return candidates[0]


def fund_values(path, evidence):
    if len(evidence["constants"]) != 1:
        raise ValueError("Choose a unique fund constants record before generating journals.")
    binding = evidence["bindings"]["constants"]
    row_number = evidence["constants"][0]["source"]["row"]
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = book[binding["sheet"]]
        headers = next(sheet.iter_rows(min_row=binding["header_row"], max_row=binding["header_row"], values_only=True))
        values = next(sheet.iter_rows(min_row=row_number, max_row=row_number, values_only=True))
        result = {}
        for key, aliases in FUND.items():
            matches = [i for i, value in enumerate(headers) if value is not None and normalise(value) in aliases]
            if len(matches) > 1:
                raise ValueError(f"Competing fund field {key}.")
            value = values[matches[0]] if matches else None
            if value is not None and not isinstance(value, str):
                raise ValueError(f"Fund field {key} requires a text value; clarify numeric identifiers before planning.")
            result[key] = value
        return result
    finally:
        book.close()


def business_tables(path, evidence):
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet, header, _ = find_table(book, exact=HEADERS)
        journal = [dict(zip(HEADERS, row[:len(HEADERS)]))
                   for row in sheet.iter_rows(min_row=header + 1, values_only=True)
                   if any(value is not None for value in row)]
        sheet, header, mapping = find_table(book, meanings=STAGING)
        staging = [{key: row[col - 1] for key, col in mapping.items()}
                   for row in sheet.iter_rows(min_row=header + 1, values_only=True)
                   if any(row[col - 1] is not None for col in mapping.values())]
        binding = evidence["bindings"]["roster"]
        sheet = book[binding["sheet"]]
        values = next(sheet.iter_rows(min_row=binding["header_row"], max_row=binding["header_row"], values_only=True))
        mapping = columns(values, TRACKER)
        tracker = []
        for investor in evidence["roster"]:
            row = next(sheet.iter_rows(min_row=investor["source"]["row"], max_row=investor["source"]["row"], values_only=True))
            tracker.append({"id": investor["id"], **{key: row[col - 1] for key, col in mapping.items()}})
        return {"journal": journal, "staging": staging, "tracker": tracker}
    finally:
        book.close()


def sheet_parts(archive):
    main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {r.attrib["Id"]: r.attrib["Target"] for r in rels}
    book = ET.fromstring(archive.read("xl/workbook.xml"))
    result = {}
    for sheet in book.find(f"{{{main}}}sheets"):
        target = targets[sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]]
        result[sheet.attrib["name"]] = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
    return result


def restore_caches(original, generated, destination):
    main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    tag = lambda name: f"{{{main}}}{name}"
    with zipfile.ZipFile(original) as source, zipfile.ZipFile(generated) as edited:
        old_parts, new_parts = sheet_parts(source), sheet_parts(edited)
        replacements = {}
        for name in old_parts.keys() & new_parts.keys():
            old = ET.fromstring(source.read(old_parts[name]))
            new = ET.fromstring(edited.read(new_parts[name]))
            cached = {cell.attrib["r"]: cell for cell in old.iter(tag("c")) if cell.find(tag("f")) is not None}
            for cell in new.iter(tag("c")):
                prior = cached.get(cell.attrib["r"])
                formula = cell.find(tag("f"))
                if prior is None or formula is None or formula.text != prior.find(tag("f")).text:
                    continue
                value = prior.find(tag("v"))
                if value is None or value.text is None:
                    continue
                current = cell.find(tag("v"))
                if current is None:
                    current = ET.SubElement(cell, tag("v"))
                current.text = value.text
                if "t" in prior.attrib:
                    cell.attrib["t"] = prior.attrib["t"]
                else:
                    cell.attrib.pop("t", None)
            replacements[new_parts[name]] = ET.tostring(new)
        with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as result:
            for item in edited.infolist():
                result.writestr(item, replacements.get(item.filename, edited.read(item.filename)))


def write_workbook(source, destination, evidence, journal):
    destination = Path(destination)
    if destination.exists() or destination.resolve() == Path(source).resolve():
        raise ValueError("Output must be a new path; original workbooks are never overwritten.")
    book = load_workbook(source, data_only=False)
    try:
        sheet, header, mapping = find_table(book, meanings=STAGING)
        if sheet.max_row > header:
            sheet.delete_rows(header + 1, sheet.max_row - header)
        names = {row["id"]: row["name"] for row in evidence["roster"]}
        for offset, receipt in enumerate(evidence["transactions"], header + 1):
            values = {"date": date.fromisoformat(receipt["date"]), "description": receipt["description"],
                      "withdrawal": None, "deposit": Decimal(receipt["amount"]),
                      "investor": names[receipt["investor_id"]]}
            for key, column in mapping.items():
                cell = sheet.cell(offset, column, values[key])
                if isinstance(values[key], str):
                    cell.data_type = "s"
        binding = evidence["bindings"]["roster"]
        sheet = book[binding["sheet"]]
        mapping = columns([c.value for c in sheet[binding["header_row"]]], TRACKER)
        for row in evidence["tracker"]:
            for key, column in mapping.items():
                value = date.fromisoformat(row[key]) if key == "receipt_date" and row[key] else None if key == "receipt_date" else Decimal(row[key])
                cell = sheet.cell(row["source"]["row"], column)
                cell.value = value
                cell.number_format = "yyyy-mm-dd" if key == "receipt_date" else "#,##0.00"
        total_row = max(row["source"]["row"] for row in evidence["tracker"]) + 1
        summary = sheet.cell(total_row, binding["columns"]["called"]).value
        if isinstance(summary, str) and summary.upper().startswith("=SUM("):
            for key in ["received", "outstanding"]:
                sheet.cell(total_row, mapping[key], sum(Decimal(row[key]) for row in evidence["tracker"]))
        sheet, header, _ = find_table(book, exact=HEADERS)
        if sheet.max_row > header:
            sheet.delete_rows(header + 1, sheet.max_row - header)
        for index, row in enumerate(journal, header + 1):
            for column, field in enumerate(HEADERS, 1):
                value = row[field]
                if field in {"GL Date", "Effective Date"} and value:
                    value = date.fromisoformat(value)
                elif "Amount" in field and value is not None:
                    value = Decimal(str(value))
                cell = sheet.cell(index, column, value)
                if isinstance(value, str):
                    cell.data_type = "s"
        with tempfile.TemporaryDirectory(dir=destination.parent) as directory:
            raw, ready = Path(directory) / "raw.xlsx", Path(directory) / "ready.xlsx"
            book.save(raw)
            restore_caches(source, raw, ready)
            os.link(ready, destination)
    finally:
        book.close()
    return destination


def equivalent(actual, expected, field):
    if actual in (None, "") or expected in (None, ""):
        return actual in (None, "") and expected in (None, "")
    if "amount" in field.lower() or field in {"received", "outstanding", "deposit", "withdrawal"}:
        try:
            return money(actual) == money(expected)
        except (ValueError, ArithmeticError):
            return False
    if "date" in field.lower():
        def day(value):
            return value.date().isoformat() if isinstance(value, datetime) else value.isoformat() if isinstance(value, date) else str(value)
        return day(actual) == day(expected)
    if field == "description":
        # Bank export formatting may split references; retain every letter/digit.
        return "".join(c for c in normalise(actual) if c.isalnum()) == "".join(c for c in normalise(expected) if c.isalnum())
    return actual == expected


def plain(value):
    """JSON value that keeps the type visible, so 1 and "1" stay distinguishable in feedback."""
    return value if value is None or isinstance(value, (bool, int, float, str)) else str(value)


def compare_tables(actual, target):
    differences = []
    for table in ["journal", "tracker", "staging"]:
        if len(actual[table]) != len(target[table]):
            differences.append({"table": table, "check": "row_count", "actual": len(actual[table]), "expected": len(target[table])})
        for index, (got, expected) in enumerate(zip(actual[table], target[table]), 1):
            for field in sorted(got.keys() | expected.keys()):
                if not equivalent(got.get(field), expected.get(field), field):
                    differences.append({"table": table, "row": index, "field": field,
                                        "actual": plain(got.get(field)), "expected": plain(expected.get(field))})
    return differences
