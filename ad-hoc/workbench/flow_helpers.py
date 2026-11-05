"""Helpers for running the agent's workflow on a REAL workbook copy, shared by test_workspace_flows.py and agent_acceptance.py.

Nothing here reads or prints client text beyond what the caller asks for: they pick an account that suits a flow, find the
cell a figure is printed in, and change a saved value in a COPY of the workbook (never the original).
"""
from __future__ import annotations

import os
import re
import zipfile
from xml.etree import ElementTree as ET


def wan(value: float) -> str:
    return f"{abs(value) / 1e4:,.1f}万元"


def date_text(period: str) -> str:
    y, m, d = (int(x) for x in period.split("-"))
    return f"{y}年{m}月{d}日"


def sheet_part(zf: zipfile.ZipFile, sheet: str) -> str:
    """The zip entry that holds a sheet's cells."""
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    book = ET.fromstring(zf.read("xl/workbook.xml"))
    rid = next(s.get("{%s}id" % ns["r"]) for s in book.find("m:sheets", ns) if s.get("name") == sheet)
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    target = next(r.get("Target") for r in rels if r.get("Id") == rid)
    return "xl/" + target.lstrip("/").replace("xl/", "")


def bump_cached_values(path: str, sheet: str, row: int, now_in_sheet_units: float, delta: float) -> list:
    """Add `delta` to the cached value of every numeric cell in an Excel row that currently holds `now_in_sheet_units`.
    Edits the saved value only (formulas are left as they are), which is what a reader of the file sees. Returns the cells."""
    changed = []
    with zipfile.ZipFile(path) as src:
        part = sheet_part(src, sheet)
        xml = src.read(part).decode("utf-8")

        def edit(match):
            if int(re.search(r"\d+", match.group("ref")).group()) != row:
                return match.group(0)
            body = match.group("body")
            v = re.search(r"<v>([^<]*)</v>", body)
            if v is None or ' t="s"' in match.group("attrs") or ' t="str"' in match.group("attrs"):
                return match.group(0)
            try:
                value = float(v.group(1))
            except ValueError:
                return match.group(0)
            if abs(value - now_in_sheet_units) > 1e-6 * max(1.0, abs(value)):
                return match.group(0)
            changed.append(match.group("ref"))
            return match.group(0).replace(v.group(0), f"<v>{value + delta!r}</v>")
        xml = re.sub(r'<c r="(?P<ref>[A-Z]+\d+)"(?P<attrs>[^>]*)>(?P<body>.*?)</c>', edit, xml, flags=re.S)
        out = path + ".tmp"
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
            for item in src.infolist():
                dst.writestr(item, xml.encode("utf-8") if item.filename == part else src.read(item.filename))
    os.replace(out, path)
    return changed


def find_figure(path: str, sheet: str, wanted: float):
    """[(row, column letter, value as printed, unit factor)] of numeric cells whose printed value, at 1x, 1000x or 10000x, is `wanted`."""
    from openpyxl import load_workbook
    out = []
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        for r, cells in enumerate(wb[sheet].iter_rows(max_col=80, values_only=True), 1):
            for c, v in enumerate(cells, 1):
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    for k in (1.0, 1e3, 1e4):
                        if abs(abs(v) * k - abs(wanted)) < 1.0:
                            out.append((r, c, v, k))
                            break
    finally:
        wb.close()
    return out


def pick_account(source, copy_path: str, require_base: bool = True):
    """An account the base run commented on, with two periods, a computed movement, an unreplaced latest total and a findable
    cell. Prefers one whose latest total has no Financials figure to follow (a change to the tab then moves the figure the
    draft cites); otherwise one that ties (the tab moves, the deck's figure does not)."""
    from openpyxl import load_workbook
    wb = load_workbook(copy_path, read_only=True, data_only=True)
    found = {"free": None, "tied": None}
    try:
        for key in list((source.base_results or {}) if require_base else source.contracts):
            contract = source.contracts.get(key)
            if not contract or len(contract.periods) < 2 or str(key).startswith("__"):
                continue
            last, prev = contract.periods[-1], contract.periods[-2]
            perm = contract.permissions.get(last) or {}
            raw = next((f for f in contract.facts if f["period"] == last and f["role"] == "total" and f["basis"] == "raw_schedule"), None)
            check = contract.check(last, "total_vs_financials") or {}
            if not (raw and perm.get("compute_movement") and perm.get("describe_total") and raw.get("sheet") and raw.get("row") is not None):
                continue
            if check.get("overridden") or abs(raw["value"]) < 5e5 or raw["sheet"] not in wb.sheetnames:
                continue
            row = int(raw["row"]) + 1
            cells = next(wb[raw["sheet"]].iter_rows(min_row=row, max_row=row, values_only=True))
            unit = next((k for k in (1.0, 1e3, 1e4) if any(isinstance(v, (int, float)) and abs(v * k - raw["value"]) < 1.0 for v in cells)), None)
            if not unit:
                continue
            slot = "free" if check.get("status") in ("missing", "not_comparable") else ("tied" if check.get("status") == "pass" else None)
            if slot and found[slot] is None:
                found[slot] = (key, last, prev, raw, row, unit)
    finally:
        wb.close()
    return found["free"] or found["tied"]


