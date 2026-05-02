#!/usr/bin/env python3
"""Run a filtered Excel-DNA skill campaign with the local Codex CLI.

This script matches the Tier 1/Tier 2 commands documented in AGENTS.md. It keeps
each scenario in its own generated run directory and records stdout, stderr, and
a small trace file. Override the runner with --runner or CODEX_CAMPAIGN_RUNNER.
"""
from __future__ import annotations

import argparse
import json
import os
import subprocess
from pathlib import Path

DEFAULT_RUNNER = "codex exec --skip-git-repo-check --sandbox workspace-write -"


def load_scenarios(path: Path) -> list[dict]:
    return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]


def scenario_prompt(scenario: dict, skill_dir: Path, activation: str) -> str:
    user_prompt = scenario["prompt"]
    if activation == "explicit":
        return f"""You are running an Excel-DNA skill evaluation scenario.

Explicitly use the local skill package at:
{skill_dir.resolve()}

Start by reading that package's SKILL.md, then load only the relevant references.
Work in the current directory. Put generated project files under workspace/ when files are needed.

Scenario id: {scenario["id"]}
Priority: {scenario.get("priority", "")}
Topic: {scenario.get("topic", "")}

User prompt:
{user_prompt}
"""

    return user_prompt


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--skill-dir", required=True, type=Path)
    ap.add_argument("--scenarios", required=True, type=Path)
    ap.add_argument("--out", required=True, type=Path)
    ap.add_argument("--priority")
    ap.add_argument("--activation", choices=["explicit", "implicit"], required=True)
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--runner", default=os.environ.get("CODEX_CAMPAIGN_RUNNER", DEFAULT_RUNNER))
    ap.add_argument("--dry-run", action="store_true", help="Validate inputs and print the selected scenario ids without invoking the runner")
    args = ap.parse_args()

    if not (args.skill_dir / "SKILL.md").exists():
        raise SystemExit(f"Skill directory is missing SKILL.md: {args.skill_dir}")
    if not args.scenarios.exists():
        raise SystemExit(f"Scenario file does not exist: {args.scenarios}")

    scenarios = load_scenarios(args.scenarios)
    if args.priority:
        scenarios = [sc for sc in scenarios if sc.get("priority") == args.priority]
    if args.limit:
        scenarios = scenarios[: args.limit]
    if not scenarios:
        raise SystemExit("No scenarios selected")

    print(f"Selected {len(scenarios)} scenario(s): {', '.join(sc['id'] for sc in scenarios)}")
    print(f"Runner: {args.runner}")
    if args.dry_run:
        return 0

    args.out.mkdir(parents=True, exist_ok=True)
    for scenario in scenarios:
        run_dir = args.out / scenario["id"]
        workspace = run_dir / "workspace"
        workspace.mkdir(parents=True, exist_ok=True)
        prompt = scenario_prompt(scenario, args.skill_dir, args.activation)

        (run_dir / "scenario.json").write_text(json.dumps(scenario, indent=2), encoding="utf-8")
        (run_dir / "prompt.txt").write_text(prompt, encoding="utf-8")

        proc = subprocess.run(
            args.runner,
            input=prompt,
            text=True,
            shell=True,
            cwd=run_dir,
            capture_output=True,
        )
        (run_dir / "answer.txt").write_text(proc.stdout, encoding="utf-8")
        (run_dir / "stderr.txt").write_text(proc.stderr, encoding="utf-8")
        (run_dir / "trace.json").write_text(
            json.dumps(
                {
                    "scenario_id": scenario["id"],
                    "returncode": proc.returncode,
                    "activation": args.activation,
                    "skill_dir": str(args.skill_dir.resolve()),
                    "skill_triggered": True if args.activation == "explicit" else None,
                    "runner": args.runner,
                    "references_read": [],
                    "commands": [],
                },
                indent=2,
            ),
            encoding="utf-8",
        )
        print(f"{scenario['id']}: returncode={proc.returncode}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
