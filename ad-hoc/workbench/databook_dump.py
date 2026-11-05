#!/usr/bin/env python3
"""A workbook as text that keeps what the extractor reads, and a way to rebuild a workbook from it. Free; no model.

    python databook_dump.py dump WORKBOOK.xlsx --out dump.txt
    python databook_dump.py restore dump.txt --out rebuilt.xlsx

Use this only when the workbook's own bytes cannot travel (text_transfer.py carries them exactly, and then nothing here is needed).

What the dump holds, per sheet, at the exact cell positions: the sheet's order, name and state (visible / hidden / veryHidden); every cell
with a value, typed (number, text, date, true/false, error); for each such cell its indentation level, whether it is bold and its number
format; each row's outline level and whether it is hidden; hidden columns; merged ranges. That is every property of a workbook that
fdd_utils/workbook reads (the extractor reads cached values, never formulas; cell styles are read for indent, bold and number format;
rows for outline level; the sheet for merged ranges and state). It reads the XML inside the file directly, so it does not load pivot
caches, charts or anything else it does not need.

What it does not hold: formulas, fonts and fills other than bold, borders, column widths, comments, charts, pivot tables, defined names,
images. A comment or a defined name that the extractor starts to read later would be missing from a rebuild.

The dump ends with a digest of everything it holds. `restore` writes the workbook, reads it back with the same code, and compares: equal
means the rebuild reproduces every captured property of the original. It does NOT make the rebuild the original. Its bytes differ, its
description says it was rebuilt, a `.rebuilt.json` sits beside it, and baseline_pack.py never calls a run's baseline complete on it.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import time
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

VERSION = 1
GENERAL = "General"
NAMES_PREFIX = "REBUILT FROM A TEXT DUMP"


def local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def col_letters(index_1based: int) -> str:
    out = ""
    n = index_1based
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def col_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def canon(obj: Any) -> str:
    return json.dumps(obj, sort_keys=True, ensure_ascii=False, separators=(",", ":"))


# ---------------------------------------------------------------------------------------------------------------------------
# reading the workbook's XML
# ---------------------------------------------------------------------------------------------------------------------------

def read_shared_strings(z: zipfile.ZipFile) -> List[str]:
    try:
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    out = []
    for si in root:
        parts = []
        for child in si:
            kind = local(child.tag)
            if kind == "t":
                parts.append(child.text or "")
            elif kind == "r":
                parts.extend(t.text or "" for t in child if local(t.tag) == "t")
        out.append("".join(parts))
    return out


def read_styles(z: zipfile.ZipFile) -> List[Dict[str, Any]]:
    """Per cellXfs index: {indent, bold, numfmt (code)}."""
    from openpyxl.styles.numbers import BUILTIN_FORMATS
    try:
        root = ET.fromstring(z.read("xl/styles.xml"))
    except KeyError:
        return []
    children = {local(el.tag): el for el in root}
    custom = {}
    if "numFmts" in children:
        for nf in children["numFmts"]:
            custom[int(nf.get("numFmtId", "0"))] = nf.get("formatCode") or GENERAL
    bold_fonts = []
    if "fonts" in children:
        for font in children["fonts"]:
            b = next((el for el in font if local(el.tag) == "b"), None)
            bold_fonts.append(b is not None and (b.get("val") or "1").lower() not in ("0", "false"))
    out = []
    if "cellXfs" in children:
        for xf in children["cellXfs"]:
            alignment = next((el for el in xf if local(el.tag) == "alignment"), None)
            try:
                indent = int(float((alignment.get("indent") if alignment is not None else 0) or 0))
            except ValueError:
                indent = 0
            font_id = int(xf.get("fontId", "0") or 0)
            fmt_id = int(xf.get("numFmtId", "0") or 0)
            out.append({"indent": indent, "bold": bold_fonts[font_id] if font_id < len(bold_fonts) else False,
                        "numfmt": custom.get(fmt_id) or BUILTIN_FORMATS.get(fmt_id) or GENERAL})
    return out


def read_sheet_list(z: zipfile.ZipFile) -> Tuple[List[Dict[str, str]], bool]:
    root = ET.fromstring(z.read("xl/workbook.xml"))
    date1904 = False
    sheets = []
    rels = {}
    rel_root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    for rel in rel_root:
        rels[rel.get("Id")] = rel.get("Target")
    for el in root.iter():
        kind = local(el.tag)
        if kind == "workbookPr":
            date1904 = (el.get("date1904") or "0").lower() in ("1", "true")
        elif kind == "sheet":
            rid = next((v for k, v in el.attrib.items() if local(k) == "id"), None)
            target = rels.get(rid, "")
            path = target.lstrip("/") if target.startswith("/") else "xl/" + target
            sheets.append({"name": el.get("name") or "", "state": el.get("state") or "visible", "path": path})
    return sheets, date1904


def number_value(text: str) -> Any:
    try:
        f = float(text)
    except ValueError:
        return text
    if f.is_integer() and abs(f) < 1e15:
        return int(f)
    return f


def read_sheet(z: zipfile.ZipFile, path: str, shared: List[str], styles: List[Dict[str, Any]], date1904: bool, numfmts: List[str],
               stats: Dict[str, int]) -> Dict[str, Any]:
    from openpyxl.styles.numbers import is_date_format
    from openpyxl.utils.datetime import MAC_EPOCH, WINDOWS_EPOCH, from_excel
    info: Dict[str, Any] = {"merged": [], "hidden_rows": [], "hidden_cols": [], "outline": {}}
    rows: List[Dict[str, Any]] = []
    current: Optional[int] = None
    cells: List[List[Any]] = []
    with z.open(path) as fh:
        for event, el in ET.iterparse(fh, events=("start", "end")):
            kind = local(el.tag)
            if event == "start":
                if kind == "row":
                    current = int(el.get("r") or 0)
                    cells = []
                    if (el.get("hidden") or "0").lower() in ("1", "true"):
                        info["hidden_rows"].append(current)
                    level = int(el.get("outlineLevel") or 0)
                    if level > 0:
                        info["outline"][str(current)] = level
                continue
            if kind == "c" and current is not None:
                ref = el.get("r") or ""
                m = re.match(r"([A-Z]+)(\d+)", ref)
                if not m:
                    el.clear()
                    continue
                style = styles[int(el.get("s") or 0)] if 0 <= int(el.get("s") or 0) < len(styles) else {"indent": 0, "bold": False, "numfmt": GENERAL}
                ctype = el.get("t") or "n"
                v = next((c for c in el if local(c.tag) == "v"), None)
                value: Any = None
                tag = "s"
                if ctype == "s" and v is not None and v.text is not None:
                    value = shared[int(v.text)] if int(v.text) < len(shared) else None
                elif ctype == "inlineStr":
                    inline = next((c for c in el if local(c.tag) == "is"), None)
                    value = "".join(t.text or "" for t in inline.iter() if local(t.tag) == "t") if inline is not None else None
                elif ctype == "str" and v is not None:
                    value = v.text or ""
                elif ctype == "b" and v is not None:
                    value, tag = (v.text or "0") in ("1", "true"), "b"
                elif ctype == "e" and v is not None:
                    value, tag = v.text or "", "e"
                elif ctype == "d" and v is not None and v.text:
                    value, tag = v.text.replace("Z", "")[:19], "d"
                elif ctype == "n" and v is not None and v.text not in (None, ""):
                    number = number_value(v.text)
                    if isinstance(number, (int, float)) and is_date_format(style["numfmt"]):
                        try:
                            when = from_excel(float(number), epoch=MAC_EPOCH if date1904 else WINDOWS_EPOCH)
                            if isinstance(when, datetime):
                                value, tag = when.isoformat(timespec="seconds"), "d"
                            else:                               # a time of day: kept as the number it is, with its format
                                value, tag = number, "n"
                        except Exception:                       # noqa: BLE001
                            value, tag = number, "n"
                    else:
                        value, tag = number, "n"
                if value is not None and value != "":
                    flags = ""
                    if style["indent"] > 0:
                        flags += f"i{style['indent']}"
                        stats["indented"] += 1
                    if style["bold"]:
                        flags += "b"
                        stats["bold"] += 1
                    if style["numfmt"] != GENERAL and tag in ("n", "d"):
                        if style["numfmt"] not in numfmts:
                            numfmts.append(style["numfmt"])
                        flags += f"f{numfmts.index(style['numfmt'])}"
                        stats["formatted"] += 1
                    cells.append([m.group(1), tag, value] + ([flags] if flags else []))
                    stats["cells"] += 1
                el.clear()
            elif kind == "row" and current is not None:
                if cells:
                    cells.sort(key=lambda c: col_index(c[0]))
                    rows.append({"k": "row", "r": current, "c": cells})
                current, cells = None, []
                el.clear()
            elif kind == "mergeCell":
                info["merged"].append(el.get("ref") or "")
            elif kind == "col" and (el.get("hidden") or "0").lower() in ("1", "true"):
                lo, hi = int(el.get("min") or 0), min(int(el.get("max") or 0), int(el.get("min") or 0) + 300)
                info["hidden_cols"].extend(col_letters(c) for c in range(lo, hi + 1))
    info["rows"] = rows
    return info


def dump_workbook(path: str) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """(the records, a summary with counts only). The records are what is written; the digest covers all but the first."""
    stats = {"cells": 0, "indented": 0, "bold": 0, "formatted": 0}
    numfmts: List[str] = [GENERAL]
    records: List[Dict[str, Any]] = []
    with zipfile.ZipFile(path) as z:
        sheets, date1904 = read_sheet_list(z)
        shared, styles = read_shared_strings(z), read_styles(z)
        body: List[Dict[str, Any]] = []
        for index, sheet in enumerate(sheets):
            info = read_sheet(z, sheet["path"], shared, styles, date1904, numfmts, stats)
            body.append({"k": "sheet", "i": index, "name": sheet["name"], "state": sheet["state"], "merged": info["merged"],
                         "hidden_rows": info["hidden_rows"], "hidden_cols": info["hidden_cols"], "outline": info["outline"]})
            for row in info["rows"]:
                body.append({**row, "sheet": index})
    digest = hashlib.sha256(("\n".join([canon({"numfmts": numfmts})] + [canon(r) for r in body])).encode("utf-8")).hexdigest()
    sha = hashlib.sha256(Path(path).read_bytes()).hexdigest()
    meta = {"k": "meta", "version": VERSION, "source_sha256": sha, "source_bytes": Path(path).stat().st_size, "sheets": len(sheets),
            "hidden_sheets": sum(1 for s in sheets if s["state"] != "visible"), "numfmts": numfmts, "dumped_at": time.strftime("%Y-%m-%dT%H:%M:%S"), **stats,
            "outline_rows": sum(len(r.get("outline", {})) for r in body if r["k"] == "sheet"),
            "merged_ranges": sum(len(r.get("merged", [])) for r in body if r["k"] == "sheet")}
    records = [meta, *body, {"k": "end", "digest": digest, "cells": stats["cells"]}]
    return records, {**{k: meta[k] for k in ("sheets", "hidden_sheets", "cells", "indented", "bold", "formatted", "outline_rows", "merged_ranges")},
                     "numfmts": len(numfmts) - 1, "digest": digest}


def write_dump(records: List[Dict[str, Any]], out: Path) -> None:
    out.write_text("FDDDUMP 1\n" + "\n".join(json.dumps(r, ensure_ascii=False, separators=(",", ":")) for r in records) + "\n", encoding="utf-8")


# ---------------------------------------------------------------------------------------------------------------------------
# restoring
# ---------------------------------------------------------------------------------------------------------------------------

def parse_dump(path: Path) -> List[Dict[str, Any]]:
    """The records from a dump file, which may sit inside a .md with fences and prose around it: only lines that are a JSON object with a
    known kind are taken."""
    records = []
    for line in path.read_text(encoding="utf-8", errors="ignore").splitlines():
        line = line.strip()
        if not line.startswith("{"):
            continue
        try:
            obj = json.loads(line)
        except json.JSONDecodeError:
            continue
        if isinstance(obj, dict) and obj.get("k") in ("meta", "sheet", "row", "end"):
            records.append(obj)
    return records


def check_records(records: List[Dict[str, Any]]) -> List[str]:
    problems = []
    if not records or records[0].get("k") != "meta":
        return ["the dump has no meta line (the start was cut off)"]
    if records[-1].get("k") != "end":
        return ["the dump has no end line (the text was cut off)"]
    if records[0].get("version") != VERSION:
        problems.append(f"written by version {records[0].get('version')} of this tool")
    body = records[1:-1]
    digest = hashlib.sha256(("\n".join([canon({"numfmts": records[0].get("numfmts")})] + [canon(r) for r in body])).encode("utf-8")).hexdigest()
    if digest != records[-1].get("digest"):
        problems.append("the lines that arrived do not match the digest at the end: a line was lost or changed in copying")
    cells = sum(len(r["c"]) for r in body if r["k"] == "row")
    if cells != records[-1].get("cells"):
        problems.append(f"the dump says {records[-1].get('cells')} cells and {cells} arrived")
    return problems


def restore_workbook(records: List[Dict[str, Any]], out: Path) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font
    numfmts = records[0]["numfmts"]
    wb = Workbook()
    wb.remove(wb.active)
    sheets: Dict[int, Any] = {}
    for rec in records[1:-1]:
        if rec["k"] == "sheet":
            ws = wb.create_sheet(rec["name"])
            sheets[rec["i"]] = ws
            for r, level in (rec.get("outline") or {}).items():
                ws.row_dimensions[int(r)].outlineLevel = int(level)
            for r in rec.get("hidden_rows") or []:
                ws.row_dimensions[int(r)].hidden = True
            for letters in rec.get("hidden_cols") or []:
                ws.column_dimensions[letters].hidden = True
            ws._dump_state = rec["state"]
            ws._dump_merged = list(rec.get("merged") or [])
        elif rec["k"] == "row":
            ws = sheets[rec["sheet"]]
            for cell in rec["c"]:
                letters, tag, value = cell[0], cell[1], cell[2]
                flags = cell[3] if len(cell) > 3 else ""
                target = ws[f"{letters}{rec['r']}"]
                if tag == "d":
                    target.value = datetime.fromisoformat(value)
                elif tag == "s":
                    target.value = value
                    target.data_type = "s"                      # a text that starts with "=" stays text
                else:
                    target.value = value
                m = re.search(r"i(\d+)", flags)
                if m:
                    target.alignment = Alignment(horizontal="left", indent=int(m.group(1)))
                if "b" in re.sub(r"i\d+|f\d+", "", flags):
                    target.font = Font(bold=True)
                m = re.search(r"f(\d+)", flags)
                if m:
                    target.number_format = numfmts[int(m.group(1))]
    for ws in sheets.values():
        for rng in ws._dump_merged:
            ws.merge_cells(rng)
        ws.sheet_state = ws._dump_state
    meta = records[0]
    wb.properties.creator = "databook_dump restore"
    wb.properties.description = f"{NAMES_PREFIX} (source sha256 {meta['source_sha256']}, digest {records[-1]['digest'][:16]}). Not the original file."
    if all(ws.sheet_state != "visible" for ws in wb.worksheets):
        wb.worksheets[0].sheet_state = "visible"                # a workbook with no visible sheet cannot be saved
    wb.save(out)
    Path(str(out) + ".rebuilt.json").write_text(json.dumps({
        "rebuilt_from_a_text_dump": True, "source_sha256": meta["source_sha256"], "source_bytes": meta["source_bytes"], "digest": records[-1]["digest"],
        "restored_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
        "note": "this file is not the original: its bytes differ. It reproduces every property the dump captured; see databook_dump.py"}, indent=1), encoding="utf-8")


def is_rebuilt(path: str) -> Optional[Dict[str, Any]]:
    """The marker written beside (and inside) a rebuilt workbook, or None for a file nobody marked."""
    side = Path(str(path) + ".rebuilt.json")
    if side.exists():
        try:
            return json.loads(side.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return {"rebuilt_from_a_text_dump": True, "note": "the marker file could not be read"}
    try:
        with zipfile.ZipFile(path) as z:
            core = z.read("docProps/core.xml").decode("utf-8", errors="ignore")
        if NAMES_PREFIX in core:
            return {"rebuilt_from_a_text_dump": True, "note": "marked inside the workbook's properties"}
    except (OSError, KeyError, zipfile.BadZipFile):
        pass
    return None


def main(argv: Optional[List[str]] = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except (ValueError, OSError):
                pass
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    d = sub.add_parser("dump")
    d.add_argument("workbook")
    d.add_argument("--out", required=True)
    r = sub.add_parser("restore")
    r.add_argument("dump")
    r.add_argument("--out", required=True)
    args = ap.parse_args(argv)
    if args.cmd == "dump":
        records, summary = dump_workbook(args.workbook)
        write_dump(records, Path(args.out))
        size = Path(args.out).stat().st_size
        print(f"dumped {summary['sheets']} sheet(s) ({summary['hidden_sheets']} hidden), {summary['cells']:,} cells; {size:,} bytes of text")
        print(f"  indented cells {summary['indented']:,}   bold cells {summary['bold']:,}   cells with a number format {summary['formatted']:,} "
              f"({summary['numfmts']} distinct)   outline rows {summary['outline_rows']:,}   merged ranges {summary['merged_ranges']:,}")
        print(f"  digest {summary['digest'][:16]}  (the same digest is printed again when the rebuild is checked)")
        return 0
    records = parse_dump(Path(args.dump))
    problems = check_records(records)
    if problems:
        print("NOT RESTORED:")
        for p in problems:
            print("  -", p)
        return 2
    out = Path(args.out)
    if out.exists():
        print(f"NOT RESTORED: {out} exists")
        return 2
    restore_workbook(records, out)
    again, summary = dump_workbook(str(out))
    same = summary["digest"] == records[-1]["digest"]
    print(f"restored {summary['sheets']} sheet(s), {summary['cells']:,} cells; indented {summary['indented']:,}, bold {summary['bold']:,}, "
          f"formatted {summary['formatted']:,}, outline rows {summary['outline_rows']:,}, merged {summary['merged_ranges']:,}")
    print(f"read back with the same code: digest {'EQUAL' if same else 'DIFFERENT'} ({summary['digest'][:16]} vs {records[-1]['digest'][:16]})")
    if not same:
        orig = {canon(x): x for x in records[1:-1]}
        back = {canon(x) for x in again[1:-1]}
        lost = [x for k, x in orig.items() if k not in back]
        print(f"  {len(lost)} record(s) of the dump are not reproduced; first: " + (f"{lost[0]['k']} {lost[0].get('r', lost[0].get('i'))} (sheet {lost[0].get('sheet', lost[0].get('i'))})" if lost else "-"))
        return 2
    print("This file is a rebuild, not the original: its bytes differ and a baseline never counts it as one.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
