"""Local Streamlit UI for the three evaluation phases. Same functions as the CLI; no extra logic."""

import json
from pathlib import Path
import time

import streamlit as st

from ylookup.adapt import PROMPT_VERSION as ADAPT_PROMPT_VERSION, adapt_workflow
from ylookup.builder import MAX_ATTEMPTS, PROMPT_VERSION, build_workflow
from ylookup.events import Events
from ylookup.model import OpenRouter, settings
from ylookup.planner import run_workflow

RUNS = Path("runs")
UPLOADS = Path("data/uploads")
OTHER = "Other (type below)"


def local_files(suffix):
    found = [p for folder in ["data", "examples", "workflows"] for p in Path(folder).glob("**/*" + suffix)]
    if suffix == ".json":
        found = [p for p in found if p.parent.name == "workflows"] + list(RUNS.glob("*/workflow.json"))
    found = [p for p in found if not p.name.startswith("~$")]
    return sorted(found, key=lambda p: p.stat().st_mtime, reverse=True)


def pick(label, suffix, hint, key):
    options = local_files(suffix)
    if not options:
        st.warning(f"No {suffix} files found. Add one in the sidebar.")
        return None
    default = next((i for i, p in enumerate(options) if hint in p.name), 0)
    return st.selectbox(label, options, index=default, key=key, format_func=str)


def new_run(kind):
    return RUNS / f"ui-{kind}-{time.strftime('%Y%m%d-%H%M%S')}"


def read_events(run_dir):
    path = Path(run_dir) / "events.jsonl"
    return [json.loads(line) for line in path.read_text().splitlines()] if path.exists() else []


def status_banner(result):
    status = result.get("status")
    text = {"success": "Success", "needs_input": "Stopped: needs your input",
            "stopped": "Stopped: could not resolve", "budget_exhausted": "Stopped: call budget used up"}.get(status, status)
    {"success": st.success, "needs_input": st.warning}.get(status, st.error)(text)


def download(label, path, key):
    path = Path(path)
    if path.exists():
        st.download_button(label, path.read_bytes(), file_name=path.name, key=key)


def answer_questions(questions, key):
    """Any question: pick an option or type an answer. Answers are kept by question id and sent on the next run."""
    answers = st.session_state.setdefault("answers", {})
    for item in questions:
        st.write(f"**{item['question']}**")
        options = item.get("options") or []
        choice = st.radio("Answer", options + [OTHER], key=f"{key}-{item['id']}", label_visibility="collapsed") if options else OTHER
        text = st.text_input("Your answer", key=f"{key}-{item['id']}-text") if choice == OTHER else choice
        if text:
            answers[item["id"]] = {"question": item["question"], "answer": text}
    if questions:
        st.caption("Answers are kept for this session; run again to use them.")


def attempts_table(run_dir):
    rows = {}
    for event in read_events(run_dir):
        attempt = event.get("attempt")
        if attempt is None:
            continue
        row = rows.setdefault(attempt, {"attempt": attempt, "result": "", "differences": None, "detail": ""})
        if event["event"] == "model_response":
            row["seconds"] = round(event.get("elapsed_seconds") or 0, 1)
        elif event["event"] == "validation":
            row["result"], row["detail"] = "invalid plan", event.get("validation_error", "")[:300]
        elif event["event"] == "comparison":
            row["differences"] = len(event["differences"])
            row["result"] = "pass" if event["status"] == "pass" else row["result"] or "differs from target"
            if event["differences"]:
                first = event["differences"][0]
                row["detail"] = row["detail"] or f"e.g. {first.get('field')}: expected {first.get('expected')}, got {first.get('actual')}"
    return list(rows.values())


def show_run_result(result, key):
    status_banner(result)
    run_dir = Path(result.get("run_dir") or ".")
    if result.get("binding_changes"):
        st.info("Source layout changed since build (resolved by meaning):\n\n- " + "\n- ".join(result["binding_changes"]))
    if result.get("reason"):
        st.write(result["reason"])
    for failure in result.get("check_failures", []):
        st.write("- " + failure)
    if result.get("questions"):
        answer_questions(result["questions"], key)
    if result["status"] == "success" and result.get("output"):
        st.caption(f"Rows written: {result['rows']} to `{result['output']}`")
        outputs = json.loads((run_dir / "outputs.json").read_text())
        for table, rows in outputs.items():
            with st.expander(f"{table} ({len(rows)} rows)", expanded=len(rows) == max(map(len, outputs.values()))):
                st.dataframe(rows, width="stretch")
        download("Download output.xlsx", result["output"], key + "-dl")
    if result.get("run_dir"):
        st.caption(f"Run folder: `{run_dir}`")


def run_panel(key, input_hint):
    workflow = pick("Saved workflow", ".json", "reconciliation", key + "-wf")
    workbook = pick("Input workbook (.xlsx)", ".xlsx", input_hint, key + "-in")
    bank = pick("Bank statement (.pdf)", ".pdf", "", key + "-bank")
    if st.button("Run workflow", type="primary", key=key + "-go", disabled=not (workflow and workbook and bank)):
        try:
            st.session_state[key] = run_workflow(workflow, workbook, bank, new_run(key), answers=st.session_state.answers)
        except Exception as error:
            st.error(f"Run failed: {error}")
    if key in st.session_state:
        show_run_result(st.session_state[key], key)


st.set_page_config(page_title="Ylookup workflow builder", layout="wide")
st.title("Ylookup workflow builder")
st.session_state.setdefault("answers", {})

with st.sidebar:
    st.subheader("Model")
    try:
        _, model = settings()
        st.success(f"`{model}`\n\nOpenRouter key found")
    except ValueError as error:
        model = None
        st.error(str(error))
    st.subheader("Add files")
    for upload in st.file_uploader("Saved to data/uploads", accept_multiple_files=True,
                                   type=["xlsx", "pdf", "txt", "md", "json"]) or []:
        UPLOADS.mkdir(parents=True, exist_ok=True)
        (UPLOADS / upload.name).write_bytes(upload.getvalue())
    st.caption("Files are read from data/, examples/, workflows/ and runs/.")

build_tab, run_tab, adapt_tab, runs_tab = st.tabs(["1 · Build", "2 · Run", "3 · Adapt (stress test)", "Run history"])

with build_tab:
    st.caption("Learns a workflow from input + completed target + context. The model proposes a plan; "
               "code executes it, compares with the target and feeds differences back; fields that already match are locked.")
    left, right = st.columns(2)
    with left:
        workbook = pick("Input workbook", ".xlsx", "input", "b-in")
        target = pick("Completed target workbook", ".xlsx", "output", "b-target")
    with right:
        bank = pick("Bank statement", ".pdf", "", "b-bank")
        context = pick("Customer context (plain text)", ".txt", "context", "b-ctx")
    attempts = st.slider("Model call budget", 1, MAX_ATTEMPTS, MAX_ATTEMPTS)
    with st.expander(f"Accepted answers ({len(st.session_state.answers)})"):
        st.json(st.session_state.answers)
        if st.button("Clear answers"):
            st.session_state.answers = {}
            st.rerun()
    if st.button("Build workflow", type="primary", disabled=not (model and workbook and target and bank and context)):
        key, model = settings()
        run_dir = new_run("build")
        with st.spinner("Building: model proposes, code executes and compares…"):
            try:
                result = build_workflow(workbook, target, bank, context, st.session_state.answers,
                                        OpenRouter(key, model, PROMPT_VERSION, 8000), Events(run_dir, (key,)), attempts)
            except Exception as error:
                result = {"status": "error", "reason": str(error)}
        st.session_state.build = {**result, "run_dir": str(run_dir)}
    if "build" in st.session_state:
        result = st.session_state.build
        status_banner(result)
        if result.get("reason"):
            st.write(result["reason"])
        table = attempts_table(result["run_dir"])
        if table:
            st.subheader("Iterations")
            st.dataframe(table, width="stretch")
        answer_questions(result.get("questions", []), "b")
        if result["status"] == "success":
            st.caption(f"Workflow saved to `{result['workflow']}` after {result['attempts']} model calls")
            st.json(json.loads(Path(result["workflow"]).read_text()), expanded=False)
            download("Download workflow.json", result["workflow"], "b-dl-wf")
            download("Download output.xlsx", result["output"], "b-dl-out")
        st.caption(f"Run folder: `{result['run_dir']}`")

with run_tab:
    st.caption("Runs a saved workflow on an input with no target, no build transcript and no model call.")
    run_panel("run", "input")

def show_adapt_result(result):
    if result.get("adapted"):
        st.success(f"Workflow updated to version {result['workflow_version']}")
    elif result.get("note"):
        st.info(result["note"])
    if result.get("mapping"):
        st.dataframe([{"column": field, **change} for field, change in result["mapping"].items()],
                     width="stretch")
    if result.get("workflow") and result.get("adapted"):
        download("Download updated workflow.json", result["workflow"], "a-dl-wf")
    show_run_result(result, "adapt")


with adapt_tab:
    st.caption("Gives the existing workflow a changed input. Renamed tabs and moved columns are re-found by meaning. "
               "If a known column is gone, the model picks its replacement from the unrecognised headers; code "
               "checks the pick and re-runs every check before saving workflow v+1. Otherwise it asks.")
    left, right = st.columns(2)
    with left:
        a_workflow = pick("Existing workflow", ".json", "reconciliation", "a-wf")
        a_input = pick("Changed input workbook", ".xlsx", "stress", "a-in")
    with right:
        a_bank = pick("Bank statement", ".pdf", "", "a-bank")
        a_context = pick("Customer context", ".txt", "context", "a-ctx")

    def adapt():
        key, model = settings()
        try:
            return adapt_workflow(a_workflow, a_input, a_bank, a_context, OpenRouter(key, model, ADAPT_PROMPT_VERSION, 1500),
                                  new_run("adapt"), st.session_state.answers)
        except Exception as error:
            return {"status": "error", "reason": str(error), "run_dir": ""}

    if st.button("Adapt workflow", type="primary", disabled=not (model and a_workflow and a_input and a_bank)):
        with st.spinner("Running the existing workflow, then adapting if needed…"):
            st.session_state.adapt = adapt()
    if "adapt" in st.session_state:
        show_adapt_result(st.session_state.adapt)

with runs_tab:
    runs = sorted((p for p in RUNS.glob("*") if (p / "events.jsonl").exists()), key=lambda p: p.stat().st_mtime, reverse=True)
    if runs:
        chosen = st.selectbox("Run", runs, format_func=lambda p: p.name)
        events = read_events(chosen)
        st.dataframe([{"#": e["sequence"], "event": e["event"], "attempt": e.get("attempt"), "status": e.get("status"),
                       "summary": str(e.get("reason") or e.get("question") or e.get("operation") or e.get("validation_error") or "")[:200]}
                      for e in events], width="stretch")
        with st.expander("Raw events"):
            st.json(events, expanded=False)
    else:
        st.info("No runs yet.")
