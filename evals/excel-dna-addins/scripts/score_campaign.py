#!/usr/bin/env python3
"""Aggregate deterministic scores for a captured Excel-DNA skill campaign."""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from score_agent_run import check_expr, read_all_text  # noqa: E402


def load_scenarios(path: Path) -> list[dict]:
    return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]


def score_one(scenario: dict, run_dir: Path) -> dict:
    trace_path = run_dir / "trace.json"
    trace = json.loads(trace_path.read_text(encoding="utf-8")) if trace_path.exists() else {}
    trace["run_dir"] = str(run_dir)
    text = read_all_text(run_dir)
    results = []
    for check in scenario.get("deterministic_checks", []):
        results.append({"check": check, "pass": check_expr(check, text, trace)})
    passed = sum(1 for item in results if item["pass"])
    total = len(results)
    return {
        "scenario_id": scenario["id"],
        "priority": scenario.get("priority"),
        "deterministic_score": passed / total if total else 1.0,
        "passed": passed,
        "total": total,
        "results": results,
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--scenarios", required=True, type=Path)
    ap.add_argument("--run-root", required=True, type=Path)
    ap.add_argument("--out", required=True, type=Path)
    ap.add_argument("--priority")
    args = ap.parse_args()

    scenarios = load_scenarios(args.scenarios)
    if args.priority:
        scenarios = [sc for sc in scenarios if sc.get("priority") == args.priority]

    scored = []
    missing = []
    for scenario in scenarios:
        run_dir = args.run_root / scenario["id"]
        if not run_dir.exists():
            missing.append(scenario["id"])
            continue
        scored.append(score_one(scenario, run_dir))

    passed = sum(1 for item in scored if item["deterministic_score"] == 1.0)
    aggregate = {
        "run_root": str(args.run_root),
        "total_selected": len(scenarios),
        "completed": len(scored),
        "missing": missing,
        "passed": passed,
        "pass_rate": passed / len(scored) if scored else 0.0,
        "scenarios": scored,
    }
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(aggregate, indent=2), encoding="utf-8")
    print(json.dumps(aggregate, indent=2))
    return 0 if not missing and passed == len(scored) else 1


if __name__ == "__main__":
    raise SystemExit(main())
