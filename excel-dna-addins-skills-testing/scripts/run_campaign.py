#!/usr/bin/env python3
"""Skeleton campaign runner.

This script intentionally does not assume a specific agent CLI. Provide a runner command
that accepts a prompt on stdin and writes the final answer to stdout. Example:

    python scripts/run_campaign.py \
      --scenarios evals/scenarios.jsonl \
      --out runs/candidate-001 \
      --runner "codex exec --full-auto"

The runner integration should be adapted to your chosen framework so traces include
skill activation, references read, commands run, token usage, and artifacts.
"""
from __future__ import annotations
import argparse
import json
import os
import subprocess
from pathlib import Path

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--scenarios", required=True, type=Path)
    ap.add_argument("--out", required=True, type=Path)
    ap.add_argument("--runner", required=True, help="Shell command for agent runner")
    ap.add_argument("--limit", type=int, default=0)
    args = ap.parse_args()

    args.out.mkdir(parents=True, exist_ok=True)
    scenarios = [json.loads(line) for line in args.scenarios.read_text(encoding="utf-8").splitlines() if line.strip()]
    if args.limit:
        scenarios = scenarios[:args.limit]

    for sc in scenarios:
        run_dir = args.out / sc["id"]
        run_dir.mkdir(parents=True, exist_ok=True)
        prompt = sc["prompt"]
        (run_dir / "scenario.json").write_text(json.dumps(sc, indent=2), encoding="utf-8")
        proc = subprocess.run(args.runner, input=prompt, text=True, shell=True, cwd=run_dir, capture_output=True)
        (run_dir / "answer.txt").write_text(proc.stdout, encoding="utf-8")
        (run_dir / "stderr.txt").write_text(proc.stderr, encoding="utf-8")
        (run_dir / "trace.json").write_text(json.dumps({
            "scenario_id": sc["id"],
            "returncode": proc.returncode,
            "skill_triggered": None,
            "references_read": [],
            "commands": []
        }, indent=2), encoding="utf-8")
        print(f"{sc['id']}: returncode={proc.returncode}")

if __name__ == "__main__":
    raise SystemExit(main())
