"""Local chat UI: upload files, talk to the builder, watch its trace. Same functions as the CLI.

Nothing assumes file names or roles. The builder works out from the content which upload is the
input, the completed example, a source or a saved workflow; the conversation is the customer's
description; and every question the builder asks appears in the chat for you to answer.
"""

from collections import Counter
import json
from pathlib import Path
import queue
import re
import threading
import time

import streamlit as st

import ylookup
from ylookup.adapt import PROMPT_VERSION as ADAPT_PROMPT_VERSION, adapt_workflow
from ylookup.builder import MAX_ATTEMPTS, PROMPT_VERSION, build_workflow, guess_roles
from ylookup.events import Events, run_name
from ylookup.model import OpenRouter, settings

RUNS = Path("runs")
UPLOADS = Path("data/uploads")
TYPES = ["xlsx", "xlsm", "pdf", "txt", "md", "json"]
WELCOME = ("**1 · Attach** the input workbook, a completed example of the same job, and any source files (PDF or workbook).  \n"
           "**2 · Describe** the job in a sentence or two, then say **build**.  \n"
           "**3 · Run** the saved workflow on a new input; if its layout changed, I adapt it.")
# The sample job of this repository, offered as one click when its files are present (the brief's names or
# the names they were supplied under). UI convenience only; the builder never sees these names.
SAMPLE = next((files for files in (
    [Path("data/capital-call-input.xlsx"), Path("data/capital-call-output.xlsx"), Path("data/bank-statement.pdf")],
    [Path("data/import-capital-call-notice-diu-input.xlsx"), Path("data/import-capital-call-notice-diu-output.xlsx"),
     Path("data/BS-Test-4.pdf")]) if all(p.exists() for p in files)), []) + [Path("examples/reconciliation-context.txt")]
OFF_TOPIC_REPLY = "I can only learn and run spreadsheet workflows, so I left that out of the job description."
LOGO = str(Path(ylookup.__file__).with_name("logo.svg"))
ASSISTANT, USER = LOGO, ":material/person:"
ICONS = {"input": ":material/input:", "completed example": ":material/task_alt:", "source": ":material/description:",
         "workflow": ":material/account_tree:", "description": ":material/notes:", "new input": ":material/fiber_new:"}
# ponytail: a word list, not a classifier; it catches requests for things the builder never makes.
# "draw"/"drawing" are left out on purpose (capital drawdowns). Ask the model if real chats slip past it.
OFF_TOPIC = {"image", "images", "illustration", "illustrations", "picture", "pictures", "photo", "photos", "video",
             "videos", "poem", "song", "logo", "cartoon", "painting", "sketch"}


def intent(text):
    """What a chat message is: "build", "run", "off_topic", "description", or None when it has no words.

    A command is a short message naming it ("build", "run it", "please adapt"); longer text is the job's
    description even when it uses those words ("Finally, I build the upload file").
    """
    words = re.findall(r"[a-z]+", text.lower())
    if not words:
        return None
    if OFF_TOPIC & set(words):
        return "off_topic"
    command = next((w for w in words if w in ("build", "run", "adapt")), None)
    if command and len(words) <= 4:
        return "build" if command == "build" else "run"
    return "description"


# --- state ----------------------------------------------------------------------------------------

def state():
    s = st.session_state
    if "folder" not in s:
        s.folder = UPLOADS / time.strftime("chat-%Y%m%d-%H%M%S")
        s.files, s.description, s.answers, s.job, s.workflow, s.built_from = [], [], {}, None, None, set()
        s.last_build = None
        s.messages = [{"role": "assistant", "text": WELCOME}]
    return s


def say(text, **extra):
    state().messages.append({"role": "assistant", "text": text, **extra})


def roles():
    """File roles from content. After a build, the files it learned from keep their roles and any other
    workbook is a new input to run the workflow on."""
    s = state()
    workbook = lambda p: p.suffix.lower() in (".xlsx", ".xlsm")
    found = guess_roles([p for p in s.files if not s.built_from or p in s.built_from or not workbook(p)])
    found["new_inputs"] = [p for p in s.files if s.built_from and workbook(p) and p not in s.built_from]
    if s.workflow:
        found["workflows"] = [s.workflow] + [w for w in found["workflows"] if w != s.workflow]
    return found


def run_input():
    """The workbook to run a workflow on: the newest upload that was not part of its build."""
    workbooks = [p for p in state().files if p.suffix.lower() in (".xlsx", ".xlsm") and p not in state().built_from]
    return workbooks[-1] if workbooks else None


def description():
    s = state()
    texts = [p.read_text() for p in s.files if p.suffix.lower() in (".txt", ".md")]
    return "\n\n".join(texts + s.description).strip()


def unpaired(r):
    """A lone workbook with nothing to compare it to: it may be the input or the completed example."""
    return r["input"] and not r["target"] and not r["workflows"]


def role_of(path, r):
    if path == r["input"] and unpaired(r):
        return "workbook: the input or the completed example (attach the other to tell)"
    return ("new input" if path in r["new_inputs"] else "input" if path == r["input"]
            else "completed example" if path == r["target"]
            else "source" if path in r["sources"] else "workflow" if path in r["workflows"]
            else "description" if path in r["context"] else "not used")


def file_lines(paths, r, bullet="- "):
    """One line per file: an icon for its role, its name, its role."""
    lines = []
    for path in paths:
        role = role_of(path, r)
        icon = ICONS.get(role, ":material/table:" if role.startswith("workbook") else ":material/block:")
        lines.append(f"{bullet}{icon} `{path.name}` · {role}")
    return ("\n" if bullet else "  \n").join(lines)


def add(paths):
    """Add files to the conversation and say what each one is."""
    s = state()
    added = [p for p in paths if p not in s.files]
    s.files += added
    s.messages.append({"role": "user", "text": "📎 " + ", ".join(f"`{p.name}`" for p in added)})
    say(f"Read {len(added)} file{'s' if len(added) != 1 else ''}:\n" + file_lines(added, roles()) + "\n\n" + next_step())


def next_step():
    r, s = roles(), state()
    if r["workflows"] and run_input() and (not s.built_from or run_input() not in s.built_from):
        return f"Say **run** to run the workflow on `{run_input().name}`; if its layout changed I adapt the workflow."
    if r["input"] and r["target"] and (r["input"], r["target"]) != s.last_build:
        if len(description().split()) < 15:
            return "Now tell me how the job is done, in your own words (or attach a description), and I will build the workflow."
        return "I have what I need. Say **build** to learn the workflow, or add more to the description first."
    if r["workflows"] and run_input():
        return f"Say **run** to run the workflow on `{run_input().name}`; if its layout changed I adapt the workflow."
    if r["workflows"]:
        return "Attach an input workbook to run the workflow on."
    if unpaired(r):
        return ("Attach the other workbook of the job (the input, or the same job completed) so I can compare them, "
                "or a saved workflow to run.")
    if sum(p.suffix.lower() in (".xlsx", ".xlsm") for p in r["sources"]) >= 2:
        return ("None of these workbooks adds rows or fills columns compared with another, so I cannot tell which is "
                "the completed example. Attach the same job completed.")
    return "Attach the input workbook and a completed example of the same job."


# --- background jobs ----------------------------------------------------------------------------------

def start(kind, run_dir, work):
    """Run work(ask) in a thread; ask(questions) waits for the answers given in the chat."""
    job = {"kind": kind, "run_dir": Path(run_dir), "pending": [], "collected": {}, "inbox": queue.Queue(),
           "result": None, "error": None}

    def ask(items):
        job["collected"], job["pending"] = {}, list(items)
        answers = job["inbox"].get()
        job["pending"] = []
        return answers

    def target():
        try:
            job["result"] = work(ask)
        except Exception as error:  # shown in the chat
            job["error"] = str(error)

    job["thread"] = threading.Thread(target=target, daemon=True)
    job["thread"].start()
    state().job = job


def answer(job, item, text):
    job["collected"][item["id"]] = {"question": item["question"], "answer": text}
    state().answers[item["id"]] = job["collected"][item["id"]]
    state().messages += [{"role": "assistant", "text": item["question"]}, {"role": "user", "text": text}]
    if all(q["id"] in job["collected"] for q in job["pending"]):
        job["inbox"].put(dict(job["collected"]))


def client(prompt_version, tokens):
    key, model = settings()
    return OpenRouter(key, model, prompt_version, tokens), key


def build():
    s, r = state(), roles()
    model_client, key = client(PROMPT_VERSION, 8000)
    run_dir = RUNS / run_name("chat-build")
    events = Events(run_dir, (key,))
    (run_dir / "context.txt").write_text(description())
    s.last_build = (r["input"], r["target"])
    files = (r["input"], r["target"], *r["sources"])
    say(f"Learning from `{r['input'].name}` → `{r['target'].name}`"
        + (f" with {', '.join(f'`{p.name}`' for p in r['sources'])}" if r["sources"] else "") + ".")
    answers = dict(s.answers)  # threads cannot read session state
    start("build", run_dir, lambda ask: {**build_workflow(r["input"], r["target"], r["sources"], run_dir / "context.txt",
                                                          answers, model_client, events, MAX_ATTEMPTS, ask),
                                         "files": [str(p) for p in files]})


def run():
    s, r = state(), roles()
    workbook = run_input()
    sources = [p for p in r["sources"] if p != workbook]
    run_dir = RUNS / run_name("chat-run")
    s.folder.mkdir(parents=True, exist_ok=True)
    context = s.folder / f"context-{time.strftime('%H%M%S')}.txt"
    context.write_text(description())
    say(f"Running `{Path(r['workflows'][0]).name}` on `{workbook.name}`"
        + (f" with {', '.join(f'`{p.name}`' for p in sources)}" if sources else "") + ".")
    try:
        model_client, _ = client(ADAPT_PROMPT_VERSION, 1500)
    except ValueError:
        model_client = None  # a run needs no model; only a renamed column does
    answers, workflow = dict(s.answers), r["workflows"][0]  # threads cannot read session state
    start("run", run_dir, lambda ask: adapt_workflow(workflow, workbook, sources, context, model_client, run_dir, answers, ask))


# --- trace ----------------------------------------------------------------------------------------------

def events(run_dir):
    found = []
    for path in sorted(Path(run_dir).rglob("events.jsonl")):
        found += [json.loads(line) for line in path.read_text().splitlines()]
    return found


DOT, OK, BAD, ASK = ":gray[⏺]", ":green[✓]", ":red[✗]", ":orange[?]"


def describe(event):
    line = describe_text(event)
    if not line:
        return None
    name, status = event["event"], event.get("status")
    mark = (OK if (name == "comparison" and status == "pass") or (name == "run_end" and status == "success")
            else BAD if name in ("validation", "model_error") or (name == "comparison") or (name == "run_end" and status != "needs_input")
            else ASK if name == "inquiry" or (name == "run_end" and status == "needs_input") else DOT)
    return f"{mark} {line}"


def describe_text(event):
    name = event["event"]
    if name == "run_start":
        return {"build": "Reading every file and comparing the input with the completed example",
                "run": "Running the saved workflow on this input (no model)"}.get(event.get("phase"))
    if name == "learned":
        outputs = "; ".join(f"`{k}` {'gets new rows' if v['op'] == 'append_rows' else 'gets ' + str(len(v['columns'])) + ' columns filled'}"
                            for k, v in event["outputs"].items())
        return f"Learned what the job produces: {outputs}. {event['checks']} checks to run every time"
    if name == "model_request":
        return f"Model call {event['attempt']}: composing the plan from that evidence"
    if name == "candidate":
        return "Plan: " + " → ".join(f"{s['op']} `{s['table']}`" for s in event["plan"]["steps"])
    if name == "validation":
        return f"Code rejected the plan: {event.get('validation_error', '')[:200]}"
    if name == "comparison":
        if event["status"] == "pass":
            return "Output matches the completed example"
        wrong = Counter(d.get("column") or d.get("check") for d in event["differences"])
        return f"{len(event['differences'])} cells differ from the example: " + ", ".join(f"{c} ({n})" for c, n in wrong.most_common(5))
    if name == "locked_fields":
        return f"Kept {len(event['restored'])} columns that already matched"
    if name == "adopted_candidates":
        return "Code used the only verified expression for " + ", ".join(event["columns"])
    if name == "inquiry" and event.get("attempt"):
        item = event.get("question")
        return "The model asked: " + (item["question"] if isinstance(item, dict) else str(item))[:200]
    if name == "linked":
        return f"Linked {event['linked']} of {event['rows']} rows"
    if name == "operation_start":
        return f"{event['operation']} `{event['table']}`"
    if name == "model_error":
        return "A model call failed; " + ("retrying" if event.get("retrying") else "giving up")
    if name == "run_end":
        status = event.get("status")
        return {"success": "Done", "needs_input": "Needs an answer", "budget_exhausted": "Out of model calls",
                "stopped": "Stopped: " + str(event.get("reason", ""))[:200]}.get(status, status)
    return None


def show_trace(run_dir):
    lines = [line for line in map(describe, events(run_dir)) if line]
    st.caption("  \n".join(lines) or "Starting…")


def show_build(result):
    saved = json.loads(Path(result["workflow"]).read_text())
    st.dataframe([{"#": n, "tool": s["op"], "table": s["table"],
                   "iterates": f"{s['for_each']}{' (linked only)' if s['only_linked'] else ''} × {s['lines']} lines"
                   if s["op"] == "append_rows" else "", "intent": s["intent"]}
                  for n, s in enumerate(saved["plan"]["steps"], 1)], width="stretch", hide_index=True)
    with st.expander("Expressions and learned checks"):
        st.dataframe([{"table": s["table"], "column": c, "expression": json.dumps(e)}
                      for s in saved["plan"]["steps"] for c, e in s.get("columns", {}).items()], width="stretch", hide_index=True)
        kinds = Counter(c["kind"] for c in saved["checks"])
        st.caption(f"{len(saved['checks'])} checks learned from the example: " + ", ".join(f"{k} {n}" for k, n in kinds.most_common()))
        if saved.get("answers"):
            st.caption("Recorded answers: " + "; ".join(f"{a['question']} → {a['answer']}" for a in saved["answers"]))
    evidence = Path(result["run_dir"]) / "evidence.json"
    if evidence.exists():
        with st.expander("What code found in the files (what the model saw, private values masked)"):
            st.json(json.loads(evidence.read_text()), expanded=False)
    download("workflow.json", result["workflow"])
    download("output.xlsx", result["output"])


def show_run(result):
    if result.get("mapping"):
        st.dataframe([{"column": k, "was": v["from"], "now": v["to"], "decided by": v["decided_by"]}
                      for k, v in result["mapping"].items()], width="stretch", hide_index=True)
    for change in result.get("binding_changes", []):
        st.caption("Found by its headers: " + change)
    outputs = Path(result.get("run_dir", "")) / "outputs.json"
    if outputs.exists():
        for table, rows in json.loads(outputs.read_text()).items():
            with st.expander(f"{table} ({len(rows)} rows)"):
                st.dataframe(rows, width="stretch")
    download("output.xlsx", result.get("output", ""))
    if result.get("adapted"):
        download("workflow.json (updated)", result["workflow"])


def download(label, path):
    path = Path(path)
    if path.is_file():
        st.download_button(f"Download {label}", path.read_bytes(), file_name=path.name, key=f"dl-{path}-{label}")


def follow_up_text(result):
    """What a person may want to chase after a run, from the run's own computed differences."""
    lines = []
    for item in result.get("follow_up", []):
        if "left_unlinked" in item:
            count = len(item["left_unlinked"])
            lines.append(f"\n`{item['table']}`: {count} row{'s' if count != 1 else ''} left unlinked to `{item['to']}`, still to place:")
            lines += [f"- {row['row']} · \"{row['text']}\"" for row in item["left_unlinked"]]
            continue
        lines.append(f"\n`{item['column']}` in `{item['table']}`:")
        if item["differs"]:
            lines.append("- not zero after their linked rows: " + ", ".join(f"{d['record']} ({d['value']})" for d in item["differs"]))
        if item["nothing_linked"]:
            records = f"{item['nothing_linked']} record{'s' if item['nothing_linked'] != 1 else ''}"
            lines.append(f"- no matching row found: {records}, so `{item['column']}` is still their full amount")
    return "\n\n**To follow up**\n\n" + "\n".join(lines) if lines else ""


def finish(job):
    s, result = state(), job["result"] or {}
    if job["error"]:
        if job["error"].startswith("Stopped by a recorded answer"):
            # A stop is a decision for this attempt, not a policy: forget it so the next build can ask again.
            s.answers = {k: v for k, v in s.answers.items() if not str(v.get("answer", "")).startswith("Stop")}
            s.last_build = None
            say("Stopped as you chose. When you are ready, attach any missing file and say **build**.", run_dir=str(job["run_dir"]))
        else:
            say(f"That failed: {job['error']}", run_dir=str(job["run_dir"]))
        return
    status = result.get("status")
    if job["kind"] == "build" and status == "success":
        s.workflow = Path(result["workflow"])
        s.built_from = {Path(p) for p in result["files"]}
        text = (f":green[✓] Learned the workflow in {result['attempts']} model call{'s' if result['attempts'] > 1 else ''} and "
                "checked it against the example. To use it, attach a new or changed input and say **run**.")
    elif job["kind"] == "run" and status == "success":
        if result.get("adapted"):
            s.workflow = Path(result["workflow"])
            text = f":green[✓] The input's layout changed, so I adapted the workflow to version {result['workflow_version']} and re-ran every check."
        else:
            text = ":green[✓] Ran the workflow: " + ", ".join(f"`{k}` {v} rows" for k, v in result.get("rows", {}).items()) + "."
        text += follow_up_text(result)
    elif status == "needs_input":
        text = ("I need these answered before I can finish:\n" + "\n".join(f"- {q['question']}" for q in result.get("questions", []))
                + f"\n\nSay **{job['kind']}** again and I will ask them one by one.")
    elif status == "budget_exhausted":
        text = "I could not match the example within the model-call budget. The trace shows what still differed."
    else:
        text = "Stopped: " + str(result.get("reason") or result.get("note") or status)
    say(text, run_dir=str(job["run_dir"]), kind=job["kind"], result=result)


# --- page --------------------------------------------------------------------------------------------

st.set_page_config(page_title="Ylookup workflow builder", page_icon=LOGO, layout="centered")
st.logo(LOGO, size="large")
st.title("Ylookup workflow builder")
st.caption("Learns a spreadsheet job from one worked example, then repeats it on new files.")
s = state()

with st.sidebar:
    try:
        st.caption(f":material/smart_toy: `{settings()[1]}`")
    except ValueError as error:
        st.error(str(error))
    r = roles()
    done = lambda ok, label: f"{':green[✓]' if ok else ':gray[○]'} {label}"
    st.markdown("  \n".join([done(bool(r["input"] and r["target"]) or bool(s.workflow and run_input()), "Files"),
                             done(len(description().split()) >= 15 or bool(s.workflow), "Description"),
                             done(bool(s.workflow), "Workflow")]))
    local = sorted(p for folder in ("data", "examples", "workflows") for p in Path(folder).glob("**/*")
                   if p.suffix.lower().lstrip(".") in TYPES and not p.name.startswith("~$") and UPLOADS not in p.parents)
    with st.form("local-files"):
        picked = st.multiselect("Files on this machine", local, format_func=str, key=f"local-{len(s.files)}")
        add_files = st.form_submit_button("Add to the conversation")
    if picked and add_files:
        add(picked)
        st.rerun()
    if s.files:
        st.caption("In this conversation")
        st.markdown(file_lines(s.files, roles(), bullet=""))
    if s.workflow:
        st.caption(f"Current workflow: `{s.workflow}`")
    if st.button("New conversation"):
        st.session_state.clear()
        st.rerun()

for message in s.messages:
    with st.chat_message(message["role"], avatar=ASSISTANT if message["role"] == "assistant" else USER):
        st.markdown(message["text"])
        if message["text"] == WELCOME and not s.files and len(SAMPLE) == 4 and all(p.exists() for p in SAMPLE):
            if st.button("Try the sample job", icon=":material/play_arrow:"):
                add(SAMPLE)
                st.rerun()
        if message.get("run_dir"):
            with st.expander("Trace"):
                show_trace(message["run_dir"])
        if message.get("result", {}).get("status") == "success":
            (show_build if message.get("kind") == "build" else show_run)(message["result"])


@st.fragment(run_every=1.0)
def live():
    job = state().job
    if job is None:
        return
    with st.chat_message("assistant", avatar=ASSISTANT):
        busy = {"build": "Building…", "run": "Running…"}.get(job["kind"], "Working…")
        with st.status("Waiting for your answer" if job["pending"] else busy, expanded=True):
            show_trace(job["run_dir"])
        current = next((q for q in job["pending"] if q["id"] not in job["collected"]), None)
        if current:
            st.markdown(f":orange[**Question**]  \n{current['question']}")
            for option in current.get("options") or []:
                if st.button(option, key=f"opt-{current['id']}-{option}"):
                    answer(job, current, option)
                    st.rerun(scope="app")
            st.caption(("Pick an option or type" if current.get("options") else "Type") + " your answer below.")
    if not job["thread"].is_alive():
        state().job = None
        finish(job)
        st.rerun(scope="app")


live()

submitted = st.chat_input("Message, or attach files", accept_file="multiple", file_type=TYPES)
if submitted:
    text = (submitted.text or "").strip()
    new = []
    upload_folder = s.folder / str(time.time_ns())
    for upload in submitted.files or []:
        upload_folder.mkdir(parents=True, exist_ok=True)
        path = upload_folder / Path(upload.name).name
        path.write_bytes(upload.getvalue())
        if path not in s.files:
            s.files.append(path)
        new.append(path)
    job = s.job
    pending = next((q for q in job["pending"] if q["id"] not in job["collected"]), None) if job else None
    if pending and text:
        answer(job, pending, text)
    else:
        attached = "".join(f"\n\n📎 `{p.name}`" for p in new)
        if text or new:
            s.messages.append({"role": "user", "text": (text or "") + attached})
        if job:
            say("I am still working on the last request; the trace above shows where I am.")
        else:
            current = roles()
            if new:
                say(f"Read {len(new)} file{'s' if len(new) != 1 else ''}:\n" + file_lines(new, current)
                    + "".join(f"\n- I could not tell: {q['question']}" for q in current["questions"]))
            asked = intent(text)
            if asked == "description":
                s.description.append(text)
            can_build = current["input"] and current["target"]
            if asked == "off_topic":
                say(OFF_TOPIC_REPLY + "\n\n" + next_step())
            elif asked == "build" and can_build:
                build()
            elif asked == "run" and current["workflows"] and run_input():
                run()
            elif asked == "description" and can_build and len(description().split()) >= 15 and (current["input"], current["target"]) != s.last_build:
                build()  # the description just became complete: start learning
            else:
                say(next_step())
    st.rerun()
