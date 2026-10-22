"""The plan a script would write from the evidence alone, with no model: the strongest link, the first iteration
that fits and the first verified candidate for every column. The builder never runs it; it exists to measure
what the model's reading of the customer's description adds (scripts/model_matters.py, tests/test_model_matters.py)."""

from .planner import POLICIES


def first_candidate_plan(profile):
    first = lambda found: found[0] if isinstance(found, list) and found else None
    steps = []
    for link in profile["links"][:1]:
        steps.append({"op": "link_rows", "intent": "link", "table": link["table"], "text": link["text"], "to": link["to"],
                      "name": link["name"], "policy": max(POLICIES, key=lambda p: link[f"linked_{p}"])})
    for output in sorted(profile["outputs"], key=lambda o: o["op"] == "append_rows"):  # fills before appends
        columns = {}
        for column, groups in output["columns"].items():
            whole = groups.get("all lines", groups.get("all rows"))
            per_line = sorted((k for k in groups if k.startswith("line ")), key=lambda k: int(k.split()[1]))
            columns[column] = first(whole) if isinstance(whole, list) or not per_line else {"line": [first(groups[k]) for k in per_line]}
        step = {"op": output["op"], "intent": "first candidate", "table": output["table"], "columns": columns}
        if output["op"] == "append_rows" and output.get("fits"):
            step.update(output["fits"][0])
        steps.append(step)
    return {"intent": "First verified candidate everywhere", "steps": steps}
