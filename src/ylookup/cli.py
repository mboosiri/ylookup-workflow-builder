"""Command line: inspect, build (phase 1), run (phase 2), adapt (phase 3) and the local UI."""

import argparse
import json
import os
from pathlib import Path
import sys
import time

from openpyxl import load_workbook

from .adapt import PROMPT_VERSION as ADAPT_PROMPT_VERSION, adapt_workflow
from .events import Events
from .model import OpenRouter, settings
from .builder import MAX_ATTEMPTS, build_workbook
from .planner import PROMPT_VERSION as FULL_PROMPT_VERSION, run_workflow


def inspect_workbook(path: Path) -> dict:
    workbook = load_workbook(path, read_only=True, data_only=False)
    try:
        sheets = []
        for sheet in workbook:
            preview = []
            populated_rows = 0
            for row in sheet.iter_rows():
                cells = []
                for cell in row:
                    if cell.value is None:
                        continue
                    value = cell.value
                    if cell.data_type == "f":
                        value = getattr(value, "text", value)
                    cells.append({"cell": cell.coordinate, "value": value})
                if cells:
                    populated_rows += 1
                    if len(preview) < 8:
                        preview.append(cells)
            sheets.append({
                "name": sheet.title,
                "reported_rows": sheet.max_row,
                "reported_columns": sheet.max_column,
                "populated_rows": populated_rows,
                "preview": preview,
            })
        return {"file": str(path), "sheets": sheets}
    finally:
        workbook.close()


def main() -> None:
    parser = argparse.ArgumentParser(description="Build, run and adapt a spreadsheet reconciliation workflow.")
    commands = parser.add_subparsers(dest="command", required=True)
    inspect = commands.add_parser("inspect", help="Read XLSX structure and a small preview.")
    inspect.add_argument("workbook", type=Path)
    full = commands.add_parser("build", help="Phase 1: infer, execute and verify a workflow against a target.")
    full.add_argument("--input", type=Path, required=True)
    full.add_argument("--target", type=Path, required=True)
    full.add_argument("--bank", type=Path, required=True)
    full.add_argument("--context", type=Path, required=True)
    full.add_argument("--answers", type=Path)
    full.add_argument("--run-dir", type=Path)
    full.add_argument("--max-attempts", type=int, choices=range(1, MAX_ATTEMPTS + 1), default=MAX_ATTEMPTS)
    run = commands.add_parser("run", help="Phase 2/3: run a saved workflow on an input; no target or model.")
    run.add_argument("--workflow", type=Path, required=True)
    run.add_argument("--input", type=Path, required=True)
    run.add_argument("--bank", type=Path, required=True)
    run.add_argument("--run-dir", type=Path)
    adapt = commands.add_parser("adapt", help="Phase 3: update a saved workflow for a changed input; saves v+1.")
    adapt.add_argument("--workflow", type=Path, required=True)
    adapt.add_argument("--input", type=Path, required=True)
    adapt.add_argument("--bank", type=Path, required=True)
    adapt.add_argument("--context", type=Path)
    adapt.add_argument("--answers", type=Path, help='JSON {"table.field": "header"} answering column_mapping questions.')
    adapt.add_argument("--run-dir", type=Path)
    commands.add_parser("ui", help="Open the local web UI.")
    args = parser.parse_args()
    if args.command == "ui":
        script = Path(__file__).with_name("ui.py")
        os.execvp(sys.executable, [sys.executable, "-m", "streamlit", "run", str(script)])
    if args.command != "inspect" and args.run_dir is None:
        args.run_dir = Path("runs") / f"{args.command}-{time.strftime('%Y%m%d-%H%M%S')}"
    try:
        if args.command == "inspect":
            result = inspect_workbook(args.workbook)
        elif args.command == "build":
            key, model = settings()
            events = Events(args.run_dir, (key,))
            answers = json.loads(args.answers.read_text()) if args.answers else {}
            result = build_workbook(args.input, args.target, args.bank, args.context, answers,
                                    OpenRouter(key, model, FULL_PROMPT_VERSION, 3500),
                                    events, args.max_attempts)
        elif args.command == "adapt":
            key, model = settings()
            result = adapt_workflow(args.workflow, args.input, args.bank, args.context,
                                    OpenRouter(key, model, ADAPT_PROMPT_VERSION, 1500), args.run_dir,
                                    json.loads(args.answers.read_text()) if args.answers else None)
        else:
            result = run_workflow(args.workflow, args.input, args.bank, args.run_dir)
    except Exception as error:
        parser.exit(1, f"Command failed: {error}\n")
    print(json.dumps(result, indent=2, default=str))
    if result.get("status") in {"needs_input", "budget_exhausted", "stopped"}:
        parser.exit(2)


if __name__ == "__main__":
    main()
