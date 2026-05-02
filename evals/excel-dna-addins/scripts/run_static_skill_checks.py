#!/usr/bin/env python3
"""Static validation for the excel-dna-addins skill package.

Usage:
    python scripts/run_static_skill_checks.py /path/to/excel-dna-addins

This is intentionally conservative. It catches package-shape errors, obvious stale guidance,
broken direct references, and dangerous phrases before any model run is needed.
"""
from __future__ import annotations
import argparse
import os
import re
import sys
from pathlib import Path

NAME_RE = re.compile(r"^[a-z0-9-]{1,64}$")
MD_LINK_RE = re.compile(r"\[[^\]]+\]\(([^)]+)\)")

STALE_OR_DANGEROUS = [
    (re.compile(r"\.NET\s*6(?:\.0)?\s*\(Long[- ]term support\)", re.I), ".NET 6 described as current LTS"),
    (re.compile(r"\.NET\s*6(?:\.0)?\s+is\s+.*current\s+LTS", re.I), ".NET 6 described as current LTS"),
    (re.compile(r"XLLs?\s+(work|run|are supported)\s+on\s+(Mac|macOS|web)", re.I), "XLL claimed to work on Mac/web"),
    (re.compile(r"disable\s+all\s+Excel\s+security\s+warnings", re.I), "blanket Excel security bypass"),
    (re.compile(r"call\s+Excel\s+COM\s+from\s+.*background\s+thread", re.I), "unsafe background-thread COM advice"),
]

def parse_frontmatter(text: str):
    if not text.startswith("---\n"):
        return None, "SKILL.md must start with YAML frontmatter"
    end = text.find("\n---", 4)
    if end == -1:
        return None, "SKILL.md frontmatter is not closed"
    raw = text[4:end].strip().splitlines()
    data = {}
    for line in raw:
        if ":" not in line:
            return None, f"Invalid frontmatter line: {line}"
        k, v = line.split(":", 1)
        data[k.strip()] = v.strip().strip('"')
    return data, None

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("skill_dir", type=Path)
    args = ap.parse_args()
    root = args.skill_dir
    errors = []
    warnings = []

    skill_md = root / "SKILL.md"
    if not skill_md.exists():
        errors.append("Missing SKILL.md")
    else:
        text = skill_md.read_text(encoding="utf-8")
        fm, err = parse_frontmatter(text)
        if err:
            errors.append(err)
        else:
            name = fm.get("name")
            desc = fm.get("description")
            if not name or not NAME_RE.match(name):
                errors.append("Invalid or missing frontmatter name")
            if not desc:
                errors.append("Missing frontmatter description")
            elif len(desc) > 1024:
                errors.append("Description exceeds 1024 characters")
            else:
                trigger_terms = ["Excel-DNA", "Excel", "add-in", "xll", "UDF", ".NET"]
                missing = [t for t in trigger_terms if t.lower() not in desc.lower()]
                if missing:
                    warnings.append(f"Description may miss trigger terms: {missing}")
        body = text.split("\n---", 1)[-1]
        body_lines = [ln for ln in body.splitlines() if ln.strip()]
        if len(body_lines) > 500:
            warnings.append(f"SKILL.md body has {len(body_lines)} nonblank lines; target <= 500")
        if "\\" in text:
            warnings.append("SKILL.md contains backslash path characters; prefer forward slashes")

        for pattern, label in STALE_OR_DANGEROUS:
            if pattern.search(text):
                errors.append(f"Stale/dangerous phrase in SKILL.md: {label}")

        # Direct links should exist for local markdown/resource paths.
        for link in MD_LINK_RE.findall(text):
            if link.startswith(("http://","https://","mailto:","#")):
                continue
            target = (root / link).resolve()
            if not str(target).startswith(str(root.resolve())):
                errors.append(f"Link escapes skill root: {link}")
            elif not target.exists():
                errors.append(f"Broken local link in SKILL.md: {link}")

    # Reference checks
    ref_dir = root / "references"
    if ref_dir.exists():
        for md in ref_dir.glob("*.md"):
            txt = md.read_text(encoding="utf-8")
            nonblank = [ln for ln in txt.splitlines() if ln.strip()]
            if len(nonblank) > 100 and not re.search(r"(?i)table of contents|^##\s+contents", txt[:2500], re.M):
                warnings.append(f"Long reference lacks obvious table of contents: {md.relative_to(root)}")
            if "\\" in txt:
                warnings.append(f"Backslash path characters in {md.relative_to(root)}")
            for pattern, label in STALE_OR_DANGEROUS:
                if pattern.search(txt):
                    errors.append(f"Stale/dangerous phrase in {md.relative_to(root)}: {label}")
            # Avoid nested local reference chains: references should not be the only path to important docs.
            for link in MD_LINK_RE.findall(txt):
                if link.startswith(("http://","https://","mailto:","#")):
                    continue
                if "/" in link.strip("./") and not link.startswith("../"):
                    warnings.append(f"Nested-looking reference link in {md.relative_to(root)}: {link}")

    # Script smoke check: existence and obvious Python syntax.
    scripts = list((root / "scripts").glob("*.py")) if (root / "scripts").exists() else []
    for script in scripts:
        try:
            compile(script.read_text(encoding="utf-8"), str(script), "exec")
        except SyntaxError as e:
            errors.append(f"Python syntax error in {script.relative_to(root)}: {e}")

    print("Static skill validation")
    print(f"Root: {root}")
    for w in warnings:
        print(f"WARNING: {w}")
    for e in errors:
        print(f"ERROR: {e}")
    if errors:
        print(f"FAIL: {len(errors)} error(s), {len(warnings)} warning(s)")
        return 1
    print(f"PASS: 0 errors, {len(warnings)} warning(s)")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
