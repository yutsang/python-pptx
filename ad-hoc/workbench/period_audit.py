"""Do the period labels the writer is shown agree with the dates written in the workbook? Free; no model call.

    python ad-hoc/workbench/period_audit.py --session <trial_<arm>_<stamp>.agent-session.json>
    python ad-hoc/workbench/period_audit.py --databook <xlsx> [--entity N] [--sheet S] [--financials-from <rollup> --financials-sheet S]

Why it exists (2026-11-11): a reviewer found 13 sentences of one acceptance book whose dates disagreed with the source
table headers while the figures they cite carried the same dates. The verifier checks a sentence against the evidence, so a
wrong label inside the evidence passes it. This reads the workbook's own date cells with a parser of its own (not the
pipeline's, so it is a second opinion) and prints
  1. per source sheet: the dates the sheet writes (rows with two or more dates, then rows with one) against the labels the
     accounts of that sheet use; labels no date cell supports, and dates written often that no account uses;
  2. the Financials sheet the same way;
  3. supporting-table columns whose period was defaulted to the latest while the header names another time
     ('期初余额', '2024年9月', 'Q3').
It prints sheet names, column headers and dates, never an amount or a row label.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ROOT))

# '1-9月' after a year: the range ends in month 9
_RANGE = re.compile(r"((?:19|20)\d{2})\s*年\s*\d{1,2}\s*[-~至到—–]\s*(\d{1,2})\s*月")
_YM = re.compile(r"(?<!\d)((?:19|20)\d{2})\s*[-/.年]\s*(\d{1,2})(?!\d)")                       # 2025-09-30, 2025/9, 2025年9月
_NM = re.compile(r"(?<![A-Za-z0-9])(\d{1,2})\s*M\s*((?:19|20)?\d{2})(?!\d)", re.I)               # 5M24, 9M2025
_MON = re.compile(r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?[\s\-/']*((?:19|20)?\d{2})(?!\d)", re.I)
_QH = re.compile(r"(?<![A-Za-z])([QH])([1-4])[\s\-/']*((?:19|20)?\d{2})(?!\d)|((?:19|20)\d{2})[\s\-/']*([QH])([1-4])", re.I)
_FY = re.compile(r"^FY\s*((?:19|20)?\d{2})\s*[AEFB]?$", re.I)                                        # FY23, FY2023E
_YEAR = re.compile(r"^((?:19|20)\d{2})\s*(?:年度?|[AEFB])?$", re.I)                                    # 2023, 2023年, 2023A
_MONTHS = {m: i for i, m in enumerate(("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"), 1)}
# a header that names a time other than 'the latest period'; 期末 / 年末 are left out because the latest is what they mean
_OTHER_TIME = re.compile(r"(?:19|20)\d{2}|\d{1,2}\s*月|(?<![A-Za-z])[QH][1-4](?!\d)|期初|年初|上期|上年|去年|同期|季|半年|"
                         r"opening|beginning|prior|previous|\bYTD\b|\bFY\s*\d", re.I)
Token = Tuple[int, Optional[int]]                   # (year, month); month None when the cell names only a year


def _year(text: str) -> int:
    value = int(text)
    return value if value > 99 else 2000 + value


def tokens(value: Any) -> List[Token]:
    """What time a cell names: (year, month), or (year, None) for 'FY23' / '2023年'. Own parser, not the pipeline's."""
    if isinstance(value, (datetime, date)):
        return [(value.year, value.month)]
    if not isinstance(value, str):
        return []
    text = _RANGE.sub(lambda m: f"{m.group(1)}年{m.group(2)}月", value.strip())
    out: List[Token] = []
    for m in _YM.finditer(text):
        if 1 <= int(m.group(2)) <= 12:
            out.append((int(m.group(1)), int(m.group(2))))
    rest = _YM.sub(" ", text)
    for m in _NM.finditer(rest):
        if 1 <= int(m.group(1)) <= 12:
            out.append((_year(m.group(2)), int(m.group(1))))
    rest = _NM.sub(" ", rest)
    for m in _MON.finditer(rest):
        out.append((_year(m.group(2)), _MONTHS[m.group(1).lower()]))
    rest = _MON.sub(" ", rest)
    for m in _QH.finditer(rest):
        letter, n, year = (m.group(1), m.group(2), m.group(3)) if m.group(1) else (m.group(5), m.group(6), m.group(4))
        out.append((_year(year), int(n) * (3 if letter.upper() == "Q" else 6)))
    rest = _QH.sub(" ", rest).strip()
    for pattern in (_FY, _YEAR):
        m = pattern.match(rest)
        if m:
            out.append((_year(m.group(1)), None))
            break
    return out


def label_token(label: str) -> Token:
    return int(label[:4]), int(label[5:7])


def show(token: Token) -> str:
    return f"{token[0]}-{token[1]:02d}" if token[1] else f"{token[0]}"


def sheet_dates(path: str, sheet: str, max_rows: int = 4000, max_cols: int = 80) -> Dict[str, Counter]:
    """{'header': rows with two or more time cells, 'single': rows with exactly one} -> Counter of (year, month)."""
    import openpyxl
    book = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet not in book.sheetnames:
            return {"header": Counter(), "single": Counter(), "missing": Counter({sheet: 1})}
        header, single = Counter(), Counter()
        for n, row in enumerate(book[sheet].iter_rows(values_only=True, max_col=max_cols), 1):
            if n > max_rows:
                break
            found = [t for cell in row for t in tokens(cell)]
            if len(found) >= 2:
                header.update(found)
            elif found:
                single.update(found)
        return {"header": header, "single": single}
    finally:
        book.close()


def find_by_sha(folder: str, sha: Optional[str]) -> Optional[str]:
    """The workbook in folder whose sha256 is sha. A trial reads a temporary copy of the workbook and deletes it, so the path a
    finished session holds no longer exists; its bytes are in the folder the trial was run from."""
    from fdd_utils.run_identity import file_sha256
    if not sha:
        return None
    for path in sorted(Path(folder).glob("*.xlsx")):
        if file_sha256(str(path)) == sha:
            return str(path)
    return None


def load_session(args: argparse.Namespace):
    from fdd_utils.agent.session import SourceSession
    if args.session:
        saved = json.loads(Path(args.session).read_text(encoding="utf-8"))
        source = saved.get("source") or {}
        params = dict(source.get("load_params") or {})
        if not params.get("workbook"):
            raise SystemExit(f"{args.session} holds no source.load_params; pass --databook instead")
        params.pop("base_run", None)
        binding = source.get("binding") or {}
        if args.databook:
            params["workbook"] = args.databook
        elif not Path(params["workbook"]).exists():
            found = find_by_sha(args.dir, binding.get("workbook_sha256"))
            if not found:
                raise SystemExit(f"the workbook this session was written on is not in {args.dir} (sha256 {str(binding.get('workbook_sha256'))[:12]}); "
                                 "pass --dir <folder> or --databook <xlsx>")
            params["workbook"] = found
        if params.get("financials_from") and not Path(params["financials_from"]).exists():
            params["financials_from"] = find_by_sha(args.dir, binding.get("financials_sha256")) or params["financials_from"]
        print(f"workbook {Path(params['workbook']).name}; sha256 {str(binding.get('workbook_sha256'))[:12]}", flush=True)
    else:
        params = {"workbook": args.databook, "entity": args.entity or "", "sheet": args.sheet or None,
                  "financials_from": args.financials_from or None, "financials_sheet": args.financials_sheet or None}
    return SourceSession.load(**params), params


def top(counter: Counter, n: int) -> str:
    return ", ".join(f"{show(t)} x{c}" for t, c in counter.most_common(n)) or "none"


def supported(label: str, seen: Iterable[Token]) -> bool:
    """A label (2025-09-30) is supported by a cell naming that month, or naming only that year ('FY25')."""
    year, month = label_token(label)
    return (year, month) in seen or (year, None) in seen


def audit(session: Any, params: Dict[str, Any], width: int = 8, verbose: bool = False) -> List[str]:
    from fdd_utils.ai import evidence_contract as ec
    out = [f"PERIOD AUDIT  workbook {Path(params['workbook']).name}  accounts {len(session.contracts)}"
           + (f"  Financials {Path(params['financials_from']).name}!{params.get('financials_sheet')}" if params.get("financials_from") else "")]
    by_sheet: Dict[str, Dict[str, Any]] = {}
    for key, contract in session.contracts.items():
        sheet = str((session.dfs[key].attrs or {}).get("source_sheet_name") or "")
        slot = by_sheet.setdefault(sheet, {"accounts": [], "labels": Counter()})
        slot["accounts"].append(key)
        slot["labels"].update(contract.periods)
    flagged, clean = 0, Counter()
    for sheet, slot in sorted(by_sheet.items(), key=lambda kv: kv[0]):
        if not sheet:
            out.append(f"(no source sheet recorded for {', '.join(slot['accounts'][:4])})")
            continue
        found = sheet_dates(params["workbook"], sheet)
        if "missing" in found:
            out.append(f"SHEET {sheet}: not found in the workbook (hidden or renamed)")
            continue
        seen = set(found["header"]) | set(found["single"])
        labels = sorted(slot["labels"])
        used = {label_token(lab) for lab in labels}
        unsupported = [lab for lab in labels if not supported(lab, seen)]
        unused = [t for t, c in found["header"].most_common() if c >= 2 and t[1] and t not in used][:width]    # month-level only
        if not (unsupported or unused):
            clean[tuple(labels)] += 1
            if not verbose:
                continue
        flagged += bool(unsupported or unused)
        out.append(f"SHEET {sheet} ({len(slot['accounts'])} acct): labels {labels}" + ("  <-- CHECK" if unsupported or unused else ""))
        out.append(f"    time cells in header rows: {top(found['header'], width)} | in single-time rows: {top(found['single'], 4)}")
        if unsupported:
            out.append(f"    labels no cell supports: {unsupported}")
        if unused:
            out.append(f"    months written 2+ times in header rows that no account uses: {[show(t) for t in unused]}")
    for labels, n in clean.most_common():
        out.append(f"{n} sheet(s) agree with their labels {list(labels)}" + ("" if not verbose else " (listed above)"))
    out.append(f"sheets to check (a label no cell supports, or a month written often that no account uses): {flagged} of {len(by_sheet)}")
    if params.get("financials_from"):
        fin = sheet_dates(params["financials_from"], str(params.get("financials_sheet") or ""))
        every = {lab for contract in session.contracts.values() for lab in contract.periods}
        seen = set(fin["header"]) | set(fin["single"])
        out.append(f"FINANCIALS {params.get('financials_sheet')}: time cells in header rows {top(fin['header'], width)} | single {top(fin['single'], 4)}")
        out.append(f"    labels no Financials cell supports: {sorted(lab for lab in every if not supported(lab, seen)) or 'none'}")
        out.append(f"    Financials months written 2+ times that no account label uses: "
                   f"{[show(t) for t, c in fin['header'].most_common() if c >= 2 and t[1] and t not in {label_token(x) for x in every}] or 'none'}")

    columns: Dict[Any, int] = {}
    for key, contract in session.contracts.items():
        for fact in contract.facts:
            if fact.get("role") == ec.ROLE_SUPPORT:
                k = (key, str(fact.get("table")), str(fact.get("column") or ""), str(fact.get("period")), bool(fact.get("period_assumed")))
                columns[k] = columns.get(k, 0) + 1
    assumed = [k for k in columns if k[4]]
    named = [k for k in assumed if _OTHER_TIME.search(k[2])]
    out.append(f"SUPPORTING columns {len(columns)}; period defaulted to the latest {len(assumed)}; of those the header names another time {len(named)}")
    for key, table, header, period, _a in named[:30]:
        out.append(f"    {key} {table} header {header!r} -> {period}")
    if len(named) > 30:
        out.append(f"    ... {len(named) - 30} more")
    return out


def main(argv: Optional[Iterable[str]] = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--session", default="", help="trial_<arm>_<stamp>.agent-session.json: the workbook and Financials it was written on")
    ap.add_argument("--dir", default="databooks", help="with --session: where to find the workbook (by sha256) when its trial copy is gone")
    ap.add_argument("--databook", default="")
    ap.add_argument("--entity", default="")
    ap.add_argument("--sheet", default="")
    ap.add_argument("--financials-from", default="")
    ap.add_argument("--financials-sheet", default="")
    ap.add_argument("--verbose", action="store_true", help="also list the sheets that agree")
    args = ap.parse_args(list(argv) if argv is not None else None)
    if not args.session and not args.databook:
        ap.error("pass --session or --databook")
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    session, params = load_session(args)
    print("\n".join(audit(session, params, verbose=args.verbose)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
