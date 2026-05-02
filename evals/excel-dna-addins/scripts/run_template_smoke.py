#!/usr/bin/env python3
"""Scaffold, validate, and build an Excel-DNA template target."""
from __future__ import annotations

import argparse
import shutil
import subprocess
from pathlib import Path


def run(command: list[str], cwd: Path, log: Path) -> int:
    proc = subprocess.run(command, cwd=cwd, text=True, capture_output=True)
    with log.open("a", encoding="utf-8") as writer:
        writer.write(f"> {' '.join(command)}\n")
        writer.write(proc.stdout)
        writer.write(proc.stderr)
        writer.write(f"\nexit_code={proc.returncode}\n\n")
    return proc.returncode


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--target", choices=["net10", "net48", "multi", "nativeaot"], default="net10")
    ap.add_argument("--out", required=True, type=Path)
    args = ap.parse_args()

    repo = Path.cwd()
    scaffold = repo / "skills" / "excel-dna-addins" / "scripts" / "scaffold_excel_dna_addin.py"
    validator = repo / "skills" / "excel-dna-addins" / "scripts" / "validate_excel_dna_project.py"
    if not scaffold.exists():
        raise SystemExit(f"Missing scaffold script: {scaffold}")
    if not validator.exists():
        raise SystemExit(f"Missing validator script: {validator}")

    if args.out.exists():
        shutil.rmtree(args.out)
    args.out.mkdir(parents=True)
    log = args.out / "template-smoke.log"

    rc = run(["python", str(scaffold), "--name", "SmokeAddIn", "--target", args.target, "--features", "async,ribbon,testing", "--output", str(args.out)], repo, log)
    if rc:
        return rc

    project = args.out / "SmokeAddIn" / "SmokeAddIn.csproj"
    rc = run(["python", str(validator), str(project)], repo, log)
    if rc:
        return rc

    rc = run(["dotnet", "build", str(project), "-v:minimal"], repo, log)
    if rc:
        return rc

    print(f"PASS: template smoke target={args.target} project={project}")
    print(f"Log: {log}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
