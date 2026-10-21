"""Version-1 numeric rule used inside the full plan, e.g. outstanding = called - received."""

from decimal import Decimal, InvalidOperation, ROUND_HALF_UP


OPERATIONS = {"add", "subtract", "multiply", "divide"}


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


def records(rows):
    if not isinstance(rows, list) or not rows:
        raise ValueError("Expected a nonempty list of records.")
    ids = []
    for row in rows:
        if not isinstance(row, dict) or not isinstance(row.get("id"), str) or not row["id"]:
            raise ValueError("Each record needs a nonempty string id.")
        ids.append(row["id"])
    if len(ids) != len(set(ids)):
        raise ValueError("Duplicate record id.")
    return rows


def validate_workflow(workflow):
    required = {"version", "intent", "operation", "left", "right", "output"}
    if not isinstance(workflow, dict) or set(workflow) != required:
        raise ValueError("Workflow has missing or unsupported fields.")
    if type(workflow["version"]) is not int or workflow["version"] != 1:
        raise ValueError("Unsupported workflow version.")
    if not isinstance(workflow["operation"], str) or workflow["operation"] not in OPERATIONS:
        raise ValueError("Unsupported arithmetic operation.")
    for key in ["intent", "left", "right", "output"]:
        if not isinstance(workflow[key], str) or not workflow[key].strip():
            raise ValueError(f"Workflow {key} must be nonempty text.")
    if "id" in [workflow["left"], workflow["right"], workflow["output"]]:
        raise ValueError("Record identifiers cannot be used as amounts or overwritten.")
    return workflow


def execute(workflow, rows):
    validate_workflow(workflow)
    result = []
    for row in records(rows):
        for field in [workflow["left"], workflow["right"]]:
            if field not in row:
                raise ValueError(f"Record {row['id']} is missing required field {field}.")
        left, right = number(row[workflow["left"]]), number(row[workflow["right"]])
        op = workflow["operation"]
        if op == "add":
            value = left + right
        elif op == "subtract":
            value = left - right
        elif op == "multiply":
            value = left * right
        else:
            if right == 0:
                raise ValueError(f"Record {row['id']}: division by zero.")
            value = left / right
        result.append({"id": row["id"], workflow["output"]: str(value)})
    return result


def compare(actual, expected):
    actual_by_id = {row["id"]: row for row in records(actual)}
    expected_by_id = {row["id"]: row for row in records(expected)}
    differences = []
    cents = Decimal("0.01")
    for record_id in sorted(actual_by_id.keys() | expected_by_id.keys()):
        got, wanted = actual_by_id.get(record_id), expected_by_id.get(record_id)
        if got is None or wanted is None:
            differences.append({"id": record_id, "check": "record_coverage",
                                "expected": wanted, "actual": got})
            continue
        for field in sorted((got.keys() | wanted.keys()) - {"id"}):
            if (field not in got or field not in wanted or
                    number(got[field]).quantize(cents, rounding=ROUND_HALF_UP) !=
                    number(wanted[field]).quantize(cents, rounding=ROUND_HALF_UP)):
                differences.append({"id": record_id, "field": field,
                                    "expected": wanted.get(field), "actual": got.get(field)})
    return differences
