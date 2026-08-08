#!/usr/bin/env python3
"""Shows exactly how a databook's Financials sheet is being carved into a
balance sheet and an income statement, and why.

The reconciliation page is built from those two blocks. When it comes back
listing operating KPIs (EBITDA%, 出租率, 单位租金) as if they were accounts,
or missing the real P&L lines entirely, the cause is in this carve-up --
but nothing printed it, so it could only be guessed at.

This prints, in order:
  1. every row that matched a balance-sheet or income-statement section
     marker, and which one was chosen as the start (the FIRST match wins,
     so a later heading that happens to contain '利润' can hijack it);
  2. the resolved boundaries for each block;
  3. the actual rows falling inside each block, flagged where a row looks
     like a ratio/per-unit metric rather than an account;
  4. what the extractor finally produced.

Read-only.

Usage:
    python inspect_financials_sheet.py "for_test/x.xlsx"
    python inspect_financials_sheet.py "for_test/x.xlsx" --sheet Financials --rows 60
"""

# moved into ad-hoc/ -- put the repo root back on sys.path so
# `import fdd_utils...` still resolves when run from anywhere.
import sys as _sys
from pathlib import Path as _Path
_sys.path.insert(0, str(_Path(__file__).resolve().parents[2]))
import argparse
import sys
import warnings

warnings.filterwarnings("ignore")
sys.path.insert(0, ".")

import pandas as pd

from fdd_utils.workbook import (
    _find_section_end_row,
    _POST_IS_SECTION_MARKERS,
    _RATIO_ROW_MARKERS,
    extract_balance_sheet_and_income_statement,
)

_BS_KEYWORDS = [
    "示意性调整后资产负债表", "示意性調整後資產負債表",
    "Indicative adjusted balance sheet", "Balance sheet",
]
_IS_KEYWORDS = [
    "示意性调整后利润表", "示意性調整後利潤表",
    "Indicative adjusted income statement", "Income statement",
    "profit and loss", "statement of comprehensive income",
]
# workbook.py falls back to these looser terms when no primary keyword hits.
# The first version of this tool only tested the primary list, so a real file
# whose headings are plain 资产负债表 / 利润表 was reported as "NO MATCH" even
# though the extractor located both blocks perfectly -- misleading, since it
# suggested the blocks could not be found when only the strict pass had failed.
_BS_KEYWORDS_RELAXED = ["资产负债表", "資產負債表", "balance sheet"]
_IS_KEYWORDS_RELAXED = ["利润表", "利潤表", "income statement", "profit and loss"]


def _row_text(df, idx) -> str:
    return " ".join(str(v) for v in df.iloc[idx].values if str(v) not in ("nan", ""))


def _looks_ratio(text: str) -> bool:
    low = text.lower()
    return any(m in low for m in _RATIO_ROW_MARKERS)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("path", help="path to the databook .xlsx")
    ap.add_argument("--sheet", default=None,
                     help="sheet holding the statements (default: auto-detect a 'Financial'-named sheet)")
    ap.add_argument("--rows", type=int, default=40, help="max rows to print per block (default 40)")
    ap.add_argument("--entity", default=None,
                     help="entity name for the pipeline comparison in section [5]")
    args = ap.parse_args()

    xl = pd.ExcelFile(args.path)
    sheet = args.sheet
    if not sheet:
        for name in xl.sheet_names:
            if "financial" in name.lower() or "财务" in name or "報表" in name or "报表" in name:
                sheet = name
                break
    if not sheet:
        sheet = xl.sheet_names[0]
    print(f"Workbook : {args.path!r}")
    print(f"Sheets   : {xl.sheet_names}")
    print(f"Using    : {sheet!r}\n")

    df = pd.read_excel(args.path, sheet_name=sheet, header=None)
    print(f"Sheet shape: {df.shape}\n")

    print("=" * 78)
    print("[1] SECTION MARKER MATCHES  (the FIRST match becomes the block start)")
    print("=" * 78)
    def _scan(keywords):
        hits = []
        for idx in range(len(df)):
            row_str = " ".join(df.iloc[idx].astype(str).values).lower()
            for kw in keywords:
                if kw.lower() in row_str:
                    hits.append((idx, kw))
                    break
        return hits

    bs_hits, is_hits = _scan(_BS_KEYWORDS), _scan(_IS_KEYWORDS)
    bs_relaxed = _scan(_BS_KEYWORDS_RELAXED) if not bs_hits else []
    is_relaxed = _scan(_IS_KEYWORDS_RELAXED) if not is_hits else []
    if bs_relaxed:
        bs_hits = bs_relaxed
    if is_relaxed:
        is_hits = is_relaxed

    def _report(hits, label, used_relaxed=False):
        if not hits:
            print(f"  {label}: NO MATCH on either the strict or the relaxed keywords "
                  f"-- this block genuinely cannot be located")
            return None
        if used_relaxed:
            print(f"  {label}: no strict-keyword match; located via the RELAXED fallback "
                  f"(this is normal, not a fault)")
        print(f"  {label}: {len(hits)} matching row(s)")
        for i, (idx, kw) in enumerate(hits):
            mark = "  <-- CHOSEN (first match)" if i == 0 else ""
            print(f"    row {idx:>4}  matched {kw!r}{mark}")
            print(f"              text: {_row_text(df, idx)[:110]}")
        if len(hits) > 1:
            print(f"    ⚠️ more than one row matches; only the first is used, so if the real")
            print(f"       statement starts at a LATER one the block will be carved from the")
            print(f"       wrong place.")
        return hits[0][0]

    bs_start = _report(bs_hits, "BALANCE SHEET", bool(bs_relaxed))
    print()
    is_start = _report(is_hits, "INCOME STATEMENT", bool(is_relaxed))

    print("\n" + "=" * 78)
    print("[2] RESOLVED BOUNDARIES")
    print("=" * 78)
    bs_end = is_start if is_start is not None else len(df)
    print(f"  balance sheet : rows {bs_start} .. {bs_end}"
          if bs_start is not None else "  balance sheet : not located")
    if is_start is not None:
        is_end = _find_section_end_row(df, is_start)
        print(f"  income stmt   : rows {is_start} .. {is_end}")
        if is_end < len(df):
            print(f"    bounded at row {is_end}: {_row_text(df, is_end)[:90]!r}")
        else:
            print(f"    runs to end of sheet (no KPI/ratio section detected below it)")
    else:
        is_end = None
        print("  income stmt   : not located")

    if is_start is not None:
        print("\n" + "=" * 78)
        print("[3] ROWS INSIDE THE INCOME-STATEMENT BLOCK")
        print("=" * 78)
        ratio_count = account_count = 0
        for idx in range(is_start, min(is_end, is_start + args.rows)):
            text = _row_text(df, idx)
            if not text.strip():
                continue
            if _looks_ratio(text):
                ratio_count += 1
                flag = "  <-- looks like a RATIO/per-unit metric, not an account"
            else:
                account_count += 1
                flag = ""
            print(f"  row {idx:>4}: {text[:100]}{flag}")
        print(f"\n  {account_count} account-like row(s), {ratio_count} ratio-like row(s) in this block")
        if ratio_count and ratio_count >= account_count:
            print("  ⚠️ this block is mostly ratios -- the start marker is probably pointing")
            print("     BELOW the real P&L (see [1]: check whether a later matching row is")
            print("     the real statement heading), or the real P&L uses a heading none of")
            print("     the keywords cover.")

        print("\n  Rows immediately BELOW the block (excluded):")
        for idx in range(is_end, min(len(df), is_end + 8)):
            text = _row_text(df, idx)
            if text.strip():
                print(f"    row {idx:>4}: {text[:100]}")

    print("\n" + "=" * 78)
    print("[4] WHAT THE EXTRACTOR ACTUALLY PRODUCED")
    print("=" * 78)
    try:
        res = extract_balance_sheet_and_income_statement(args.path, sheet)
    except Exception as exc:
        print(f"  ❌ extraction raised {type(exc).__name__}: {exc}")
        return 1
    for key in ("balance_sheet", "income_statement"):
        d = (res or {}).get(key)
        if d is None or getattr(d, "empty", True):
            print(f"  {key}: None/empty")
            continue
        print(f"  {key}: {d.shape[0]} rows x {d.shape[1]} cols")
        for v in d.iloc[:, 0].tolist()[:25]:
            t = str(v)
            print(f"    {t[:70]}" + ("   <-- ratio-like" if _looks_ratio(t) else ""))

    # The reconciliation page is built from process_workbook_data's own
    # bs_is_results, which selects its Financials sheet through the full
    # pipeline rather than this tool's auto-detection. If the two disagree,
    # the direct call above is not what the app is actually reconciling --
    # so compare them explicitly instead of assuming they match.
    print("\n" + "=" * 78)
    print("[5] WHAT THE PIPELINE ITSELF USES (this is what reconciliation reads)")
    print("=" * 78)
    try:
        from fdd_utils.workbook import process_workbook_data
        pr = process_workbook_data(temp_path=args.path, entity_name=args.entity or "x",
                                    selected_sheet=None)
    except Exception as exc:
        print(f"  ❌ process_workbook_data raised {type(exc).__name__}: {str(exc)[:160]}")
        return 1
    pipe = pr.get("bs_is_results") or {}
    for key in ("balance_sheet", "income_statement"):
        d = pipe.get(key)
        if d is None or getattr(d, "empty", True):
            print(f"  {key}: None/empty")
            continue
        rows = [str(v) for v in d.iloc[:, 0].tolist()]
        ratio_rows = [t for t in rows if _looks_ratio(t)]
        print(f"  {key}: {d.shape[0]} rows x {d.shape[1]} cols"
              f"   ({len(ratio_rows)} ratio-like)")
        for t in rows[:30]:
            print(f"    {t[:70]}" + ("   <-- RATIO, should not be here" if _looks_ratio(t) else ""))
        if ratio_rows:
            print(f"  ⚠️ these ratio rows are what surface on the reconciliation page as")
            print(f"     accounts with rate values read as amounts.")

    direct_is = (res or {}).get("income_statement")
    pipe_is = pipe.get("income_statement")
    if direct_is is not None and pipe_is is not None:
        same = list(direct_is.iloc[:, 0]) == list(pipe_is.iloc[:, 0])
        print(f"\n  direct call vs pipeline income statement: "
              f"{'IDENTICAL' if same else 'DIFFERENT -- the pipeline is reading something else'}")
        if not same:
            print(f"    direct  : {[str(v)[:18] for v in direct_is.iloc[:, 0]][:12]}")
            print(f"    pipeline: {[str(v)[:18] for v in pipe_is.iloc[:, 0]][:12]}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
