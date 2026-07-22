#!/usr/bin/env python3
"""Structural consistency check across every 'AB-' prefixed raw-data tab.

Before automating a bridge/waterfall chart FROM the raw AB-* tabs (instead
of reading a pre-built '<entity>-量价桥图' helper tab, which per the
project team likely only exists for one entity so far), we need evidence
of whether every AB-* tab actually follows the same structural convention
confirmed on AB-CD:
  - a small number of "type tag" rows near the top, each containing the
    SAME short text value repeated across many consecutive columns (this
    is what a SUMIFS criteria-range formula matches against elsewhere)
  - a 'Year' row and a 'Days in period' row (formula-detectable: YEAR(...)
    and end-start+1 patterns)
  - metric rows (area/revenue/etc.) that repeat once per type at a
    consistent row offset

This is exploratory -- it does NOT assume AB-CD's exact row numbers (24,
33, 42, etc.) apply anywhere else; it re-derives candidate rows per tab
from generic signal, then reports a compact summary so we can see how much
the format actually varies before committing to one detection strategy.

Usage:
    python inspect_ab_tabs_structure.py "databooks/xx.xlsx"
    python inspect_ab_tabs_structure.py "databooks/xx.xlsx" --tab AB-KS1 --verbose
"""
import argparse
import re
import sys
from collections import Counter
from typing import Dict, List, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


# Excludes period/date-range labels like "LTM Jun 26" / "2024年" -- these
# repeat across many columns exactly like a real type tag does (confirmed:
# 13/17 real tabs flagged row 1 as a "type tag" purely because of this),
# but a genuine asset-type name doesn't contain digits or a month/year word.
_PERIOD_LABEL_RE = re.compile(
    r"\d|年|月|日|LTM|Q[1-4]\b|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec", re.IGNORECASE
)


def _is_short_text(v) -> bool:
    if not (isinstance(v, str) and 1 <= len(v.strip()) <= 8):
        return False
    text = v.strip()
    if text.replace(".", "").isdigit():
        return False
    if _PERIOD_LABEL_RE.search(text):
        return False
    return True


def find_tag_rows(ws, min_repeat: int = 4, max_scan_row: int = 20) -> List[Tuple[int, Dict[str, int]]]:
    """A tag row: mostly-empty except for a small number of DISTINCT short
    text values, each repeated >= min_repeat times across the row (i.e. one
    value tags a contiguous run of month-columns). Returns (row, {label:
    count}) for every row in the first max_scan_row rows that qualifies."""
    hits = []
    for row in ws.iter_rows(min_row=1, max_row=min(max_scan_row, ws.max_row)):
        counts = Counter()
        for cell in row:
            if _is_short_text(cell.value):
                counts[cell.value.strip()] += 1
        qualifying = {label: n for label, n in counts.items() if n >= min_repeat}
        if qualifying:
            hits.append((row[0].row, qualifying))
    return hits


def find_year_days_rows(ws_values, ws_formulas, max_scan_row: int = 20) -> Dict[str, Optional[int]]:
    year_row = days_row = None
    for r in range(1, min(max_scan_row, ws_formulas.max_row) + 1):
        for c in range(1, ws_formulas.max_column + 1):
            f = ws_formulas.cell(row=r, column=c).value
            if not isinstance(f, str):
                continue
            if year_row is None and re.search(r"YEAR\s*\(", f, re.IGNORECASE):
                year_row = r
            if days_row is None and re.search(r"=\s*\+?\w+\d+\s*-\s*\w+\d+\s*\+\s*1", f):
                days_row = r
        if year_row and days_row:
            break
    return {"year_row": year_row, "days_row": days_row}


def find_metric_row_candidates(ws_values, tag_rows: List[int], max_scan_row: int = 60) -> List[int]:
    """Rows below the tag rows that are mostly numeric (not text-heavy) and
    have enough non-empty cells to be real data, not stray notes -- a loose
    filter, meant for eyeballing stride patterns, not a final answer."""
    if not tag_rows:
        return []
    start = max(tag_rows) + 1
    candidates = []
    for r in range(start, min(max_scan_row, ws_values.max_row) + 1):
        numeric_count = 0
        total_count = 0
        for c in range(1, ws_values.max_column + 1):
            v = ws_values.cell(row=r, column=c).value
            if v is None:
                continue
            total_count += 1
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                numeric_count += 1
        if total_count >= 5 and numeric_count / total_count >= 0.7:
            candidates.append(r)
    return candidates


def _stride_summary(rows: List[int]) -> str:
    if len(rows) < 2:
        return "(not enough rows to detect a stride)"
    diffs = [b - a for a, b in zip(rows, rows[1:])]
    return f"diffs={diffs}"


def dump_rows(ws_values, ws_formulas, rows: List[int]):
    """Full raw content (values, formulas where present) for specific rows
    -- for verifying what a flagged row ACTUALLY contains, e.g. confirming
    a candidate tag row is a real type label vs. a period marker."""
    for r in rows:
        print(f"\n--- row {r} ---")
        cells = []
        for c in range(1, ws_values.max_column + 1):
            v = ws_values.cell(row=r, column=c).value
            f = ws_formulas.cell(row=r, column=c).value
            if v is None and f is None:
                continue
            addr = get_column_letter(c)
            if isinstance(f, str) and f.startswith("="):
                cells.append(f"{addr}={f!r} -> {v!r}")
            else:
                cells.append(f"{addr}={v!r}")
        print("  " + " | ".join(cells) if cells else "  (empty)")


def _parse_row_spec(spec: str) -> List[int]:
    rows = []
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            lo, hi = part.split("-")
            rows.extend(range(int(lo), int(hi) + 1))
        else:
            rows.append(int(part))
    return rows


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("path", help="path to the databook .xlsx")
    ap.add_argument("--tab", default=None, help="only inspect this one AB- tab (default: all of them)")
    ap.add_argument("--verbose", "-v", action="store_true", help="print full tag-row/metric-row detail per tab")
    ap.add_argument("--dump-rows", default=None, metavar="SPEC",
                     help="skip the structural scan and just dump full raw content for these rows of "
                          "--tab, e.g. --tab AB-KS1 --dump-rows 1-20,25")
    args = ap.parse_args()

    print(f"Loading {args.path!r}...")
    wb_values = load_workbook(args.path, data_only=True)
    wb_formulas = load_workbook(args.path, data_only=False)

    if args.dump_rows:
        if not args.tab:
            print("❌ --dump-rows requires --tab")
            return 1
        if args.tab not in wb_values.sheetnames:
            print(f"❌ tab {args.tab!r} not found.")
            return 1
        dump_rows(wb_values[args.tab], wb_formulas[args.tab], _parse_row_spec(args.dump_rows))
        return 0

    all_sheets = wb_values.sheetnames
    ab_tabs = [s for s in all_sheets if s.startswith("AB-") or s.startswith("AB")]
    if args.tab:
        ab_tabs = [s for s in ab_tabs if s == args.tab]
        if not ab_tabs:
            print(f"❌ tab {args.tab!r} not found among AB- sheets.")
            return 1

    bridge_like = [s for s in all_sheets if ("桥图" in s or "桥" in s) and s not in ab_tabs]
    print(f"\nPre-built bridge-style tab(s) already in this workbook ({len(bridge_like)}): {bridge_like}")
    print(f"Inspecting {len(ab_tabs)} 'AB-' tab(s) for structural consistency with AB-CD...\n")

    print(f"{'Tab':16s} {'Dims':14s} {'TagRows':10s} {'#Types':7s} {'YearRow':8s} {'DaysRow':8s} {'MetricCandidates(top20)'}")
    print("-" * 110)
    for tab in ab_tabs:
        ws_v = wb_values[tab]
        ws_f = wb_formulas[tab]
        tag_hits = find_tag_rows(ws_v)
        tag_rows = [r for r, _ in tag_hits]
        n_types = len({label for _, labels in tag_hits for label in labels})
        yd = find_year_days_rows(ws_v, ws_f)
        metric_candidates = find_metric_row_candidates(ws_v, tag_rows)[:20]
        print(f"{tab:16s} {ws_v.dimensions:14s} {str(tag_rows):10s} {n_types:<7d} "
              f"{str(yd['year_row']):8s} {str(yd['days_row']):8s} {metric_candidates}")
        if args.verbose:
            for r, labels in tag_hits:
                print(f"    row {r}: {labels}")
            print(f"    metric candidate stride: {_stride_summary(metric_candidates)}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
