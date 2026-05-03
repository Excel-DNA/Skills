#!/usr/bin/env python3
"""Score one captured agent run against a scenario.

Expected run directory shape:
    answer.txt
    trace.json              optional
    workspace/              optional generated files

The scenario JSON is one line from evals/scenarios.jsonl.
This scorer implements cheap deterministic checks. Rubric grading is separate.
"""
from __future__ import annotations
import argparse
import json
import os
import re
from pathlib import Path

def read_all_text(run_dir: Path) -> str:
    parts = []
    answer = run_dir / "answer.txt"
    if answer.exists():
        parts.append(answer.read_text(encoding="utf-8", errors="ignore"))
    workspace = run_dir / "workspace"
    if workspace.exists():
        for path in workspace.rglob("*"):
            if path.is_file() and path.suffix.lower() in {".cs", ".vb", ".fs", ".csproj", ".vbproj", ".fsproj", ".xml", ".dna", ".md", ".txt", ".ps1", ".wxs"}:
                parts.append(f"\n--- {path.relative_to(workspace)} ---\n")
                parts.append(path.read_text(encoding="utf-8", errors="ignore"))
    return "\n".join(parts)


def workspace_files(run_dir: Path, suffixes: set[str] | None = None) -> list[Path]:
    workspace = run_dir / "workspace"
    if not workspace.exists():
        return []
    files = [path for path in workspace.rglob("*") if path.is_file()]
    if suffixes is not None:
        suffixes = {suffix.lower() for suffix in suffixes}
        files = [path for path in files if path.suffix.lower() in suffixes]
    return files


def file_text_contains(paths: list[Path], needle: str) -> bool:
    needle = needle.lower()
    for path in paths:
        try:
            if needle in path.read_text(encoding="utf-8", errors="ignore").lower():
                return True
        except OSError:
            continue
    return False


def check_expr(expr: str, text: str, trace: dict) -> bool:
    expr = expr.strip()
    if not expr:
        return True
    if expr == "skill_triggered==true":
        return bool(trace.get("skill_triggered"))
    if expr == "skill_triggered==false":
        return not bool(trace.get("skill_triggered"))
    if expr.startswith("no "):
        needle = expr[3:].strip().strip("'\"")
        return needle.lower() not in text.lower()
    if expr.startswith("answer mentions "):
        needle = expr[len("answer mentions "):].strip().strip("'\"")
        return needle.lower() in text.lower()
    if expr.startswith("contains "):
        needle = expr[len("contains "):].strip().strip("'\"")
        return needle.lower() in text.lower()
    if expr.startswith("does not contain "):
        needle = expr[len("does not contain "):].strip().strip("'\"")
        return needle.lower() not in text.lower()
    if expr == "project contains .csproj":
        return bool(workspace_files(Path(trace.get("run_dir", ".")), {".csproj"})) or ".csproj" in text.lower()
    if expr.startswith("csproj contains "):
        needle = expr[len("csproj contains "):].strip().strip("'\"")
        run_dir = Path(trace.get("run_dir", "."))
        return file_text_contains(workspace_files(run_dir, {".csproj"}), needle) or needle.lower() in text.lower()
    if expr.startswith("source contains "):
        needle = expr[len("source contains "):].strip().strip("'\"")
        run_dir = Path(trace.get("run_dir", "."))
        return file_text_contains(workspace_files(run_dir, {".cs", ".fs", ".vb"}), needle) or needle.lower() in text.lower()
    if expr == "dotnet build passes when SDK/runtime available":
        commands = trace.get("commands", [])
        return any(
            "dotnet build" in str(command.get("command", "")).lower() and command.get("returncode") == 0
            for command in commands
            if isinstance(command, dict)
        ) or "build succeeded" in text.lower()
    # Fallback: all words must appear somewhere.
    tokens = re.findall(r"[A-Za-z0-9_.#<>-]+", expr)
    return all(tok.lower() in text.lower() for tok in tokens)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--scenario", required=True, help="Path to a JSON scenario file or '-' for stdin")
    ap.add_argument("--run-dir", required=True, type=Path)
    args = ap.parse_args()

    if args.scenario == "-":
        scenario = json.loads(input())
    else:
        scenario = json.loads(Path(args.scenario).read_text(encoding="utf-8"))

    trace_path = args.run_dir / "trace.json"
    trace = json.loads(trace_path.read_text(encoding="utf-8")) if trace_path.exists() else {}
    trace["run_dir"] = str(args.run_dir)
    text = read_all_text(args.run_dir)

    results = []
    for check in scenario.get("deterministic_checks", []):
        ok = check_expr(check, text, trace)
        results.append({"check": check, "pass": ok})

    passed = sum(1 for r in results if r["pass"])
    total = len(results)
    score = passed / total if total else 1.0
    out = {
        "scenario_id": scenario.get("id"),
        "priority": scenario.get("priority"),
        "deterministic_score": score,
        "passed": passed,
        "total": total,
        "results": results,
    }
    print(json.dumps(out, indent=2))
    return 0 if score == 1.0 else 1

if __name__ == "__main__":
    raise SystemExit(main())
