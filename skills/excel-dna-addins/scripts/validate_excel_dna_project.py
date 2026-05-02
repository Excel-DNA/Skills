#!/usr/bin/env python3
"""Lightweight Excel-DNA project validator.

This is intentionally conservative. It parses SDK-style and old-style project XML enough to flag common modernization issues, including NativeAOT preview pitfalls.
"""
from __future__ import annotations

import argparse
import re
import sys
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class PackageRef:
    include: str
    version: str
    condition: str


def strip_ns(tag: str) -> str:
    return tag.split('}', 1)[-1]


def read_xml(path: Path) -> ET.Element:
    text = path.read_text(encoding="utf-8")
    return ET.fromstring(text)


def elements(root: ET.Element, name: str):
    for el in root.iter():
        if strip_ns(el.tag) == name:
            yield el


def get_texts(root: ET.Element, name: str) -> list[str]:
    return [(el.text or "").strip() for el in elements(root, name) if (el.text or "").strip()]


def package_refs(root: ET.Element) -> list[PackageRef]:
    refs: list[PackageRef] = []
    for el in elements(root, "PackageReference"):
        include = el.attrib.get("Include") or el.attrib.get("Update")
        if include:
            version = el.attrib.get("Version", "")
            for child in el:
                if strip_ns(child.tag) == "Version" and (child.text or "").strip():
                    version = (child.text or "").strip()
            refs.append(PackageRef(include=include, version=version, condition=el.attrib.get("Condition", "")))
    return refs


def prop_bool(root: ET.Element, name: str) -> bool | None:
    values = get_texts(root, name)
    if not values:
        return None
    value = values[-1].strip().lower()
    if value in {"true", "1", "yes"}:
        return True
    if value in {"false", "0", "no"}:
        return False
    return None


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("project", help="Path to .csproj/.fsproj/.vbproj")
    args = parser.parse_args()
    path = Path(args.project)
    root = read_xml(path)

    warnings: list[str] = []
    refs = package_refs(root)
    refs_lower = {r.include.lower() for r in refs}
    ordinary_refs = [r for r in refs if r.include.lower() == "exceldna.addin"]
    native_aot_refs = [r for r in refs if r.include.lower() == "exceldna.addin.nativeaot"]
    has_native_aot = bool(native_aot_refs)
    has_ordinary = bool(ordinary_refs)

    if "excel-dna" in refs_lower:
        warnings.append("Deprecated/old package reference 'Excel-DNA' found; prefer 'ExcelDna.AddIn' for managed add-ins or 'ExcelDna.AddIn.NativeAOT' for NativeAOT preview projects.")

    if not has_ordinary and not has_native_aot:
        warnings.append("No PackageReference to 'ExcelDna.AddIn' or 'ExcelDna.AddIn.NativeAOT' found.")

    if has_native_aot and has_ordinary:
        unconditional = [r for r in native_aot_refs + ordinary_refs if not r.condition]
        if unconditional:
            warnings.append("NativeAOT preview project appears to mix 'ExcelDna.AddIn' and 'ExcelDna.AddIn.NativeAOT' unconditionally; prefer a NativeAOT-only project or condition/version-align package references.")
        else:
            warnings.append("Project references both managed and NativeAOT Excel-DNA packages conditionally; verify all ExcelDna.* versions are aligned and build outputs are separated.")

    if has_native_aot and "exceldna.intellisense" in refs_lower:
        warnings.append("NativeAOT preview baseline should not reference 'ExcelDna.IntelliSense'; the overlay UI is not yet a safe NativeAOT feature.")

    tfms = []
    for value in get_texts(root, "TargetFramework"):
        tfms.append(value)
    for value in get_texts(root, "TargetFrameworks"):
        tfms.extend([x.strip() for x in value.split(';') if x.strip()])

    if not tfms:
        warnings.append("No TargetFramework/TargetFrameworks found.")

    for tfm in tfms:
        if re.match(r"net[0-9]+\.[0-9]+$", tfm) and not tfm.endswith("-windows"):
            warnings.append(f"Modern .NET target '{tfm}' should usually use '-windows' for Excel-DNA/Windows desktop integration.")
        if tfm.startswith("net6.0") or tfm.startswith("net7.0"):
            warnings.append(f"Target '{tfm}' is out of support; consider net48 for stability or a currently supported modern .NET target.")

    explicit = prop_bool(root, "ExcelAddInExplicitExports")
    if explicit is not True:
        warnings.append("Consider <ExcelAddInExplicitExports>true</ExcelAddInExplicitExports> for production projects.")

    if has_native_aot:
        publish_aot = prop_bool(root, "PublishAot")
        if publish_aot is not True:
            warnings.append("NativeAOT preview project should set <PublishAot>true</PublishAot>.")

        rids = get_texts(root, "RuntimeIdentifier") + [rid for value in get_texts(root, "RuntimeIdentifiers") for rid in value.split(';') if rid.strip()]
        rids = [rid.strip() for rid in rids if rid.strip()]
        if not rids:
            warnings.append("NativeAOT preview project should set <RuntimeIdentifier>win-x64</RuntimeIdentifier>.")
        elif "win-x64" not in {rid.lower() for rid in rids}:
            warnings.append(f"NativeAOT preview currently targets 64-bit Excel; expected RuntimeIdentifier 'win-x64', found {', '.join(rids)}.")

        if not any(t.startswith("net10.0-windows") for t in tfms):
            warnings.append("Current Excel-DNA NativeAOT preview examples use net10.0-windows; verify package support carefully for other target frameworks.")

        if prop_bool(root, "RunExcelDnaPack") is True:
            warnings.append("NativeAOT preview projects normally publish and load the published '*-AddIn64.xll'; do not rely on ordinary ExcelDnaPack output.")
    else:
        if any(t.startswith("net8.0") or t.startswith("net9.0") or t.startswith("net10.0") for t in tfms):
            roll = get_texts(root, "RollForward")
            if not roll:
                warnings.append("Modern .NET target has no RollForward policy; choose one deliberately and document runtime prerequisites.")

    if warnings:
        print(f"Excel-DNA project review for {path}:")
        for warning in warnings:
            print(f"- {warning}")
        sys.exit(1)
    else:
        print(f"No common Excel-DNA project issues found in {path}.")


if __name__ == "__main__":
    main()
