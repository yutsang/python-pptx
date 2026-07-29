#!/usr/bin/env python3
"""Measures the gap between what an IS account's DATA warrants and what
the current prompt rules ALLOW the commentary to say.

Context: the current pipeline applies a FIXED sentence/word cap per
account type (prompts.yml, 2_Auditor) regardless of whether that account
was flat and boring or swung several hundred percent -- and separately
tells the model to DELETE any driver not explicitly stated in the data or
remarks. Together those two rules mean a large, genuinely interesting
movement gets the same terse composition-plus-figures treatment as a
stable one. Nothing currently measures how often that actually bites.

For every IS account this prints:
  * period-over-period movement on the account's own total row, with
    partial (annualised) periods marked, since a 1-month tail period
    otherwise looks like a catastrophic decline against a full year;
  * whether remarks/notes exist that COULD explain the movement;
  * the sentence/word cap that account type will hit;
  * a verdict flagging accounts where a large movement has explanatory
    material available but the cap leaves no room to use it.

It also computes the revenue-vs-expense growth asymmetry across the
whole statement -- an observation no single account's commentary can
currently make, because each account is generated in isolation from its
own DataFrame with no visibility of any other account.

Read-only: loads and analyses, never writes to the databook.

Usage:
    python inspect_is_variance.py "for_test/xxx.xlsx" --entity "Name"
    python inspect_is_variance.py "for_test/xxx.xlsx" --entity "Name" --account "G&A expenses"
"""
import argparse
import sys
import warnings

warnings.filterwarnings("ignore")
sys.path.insert(0, ".")

from fdd_utils.workbook import process_workbook_data, INTERNAL_ROW_KEY


# Mirrors the caps in prompts.yml 2_Auditor (Eng + Chi). Kept as data here
# purely so this tool can REPORT which cap an account will hit -- it does
# not drive any generation.
_CAP_TIERS = [
    (("cash", "ar", "receivable", "prepayment", "oci", "reserve", "dta", "ncl"),
     "1-3 sentences / 25-80 words (Chi 2-3 句 / 40-90 字)"),
    (("investment propert", "op", "other payable"),
     "4-7 sentences / 100-200 words (Chi 4-7 句 / 150-280 字)"),
    (("operating income", "revenue", "cogs", "operating cost"),
     "2-3 sentences / 60-100 words (Chi 3-5 句 / 100-180 字)"),
    (("financial expense", "tax", "g&a", "general and admin", "s&d", "selling",
      "income tax", "non-operating"),
     "1-3 sentences / 30-80 words (Chi 2-3 句 / 60-130 字)  <-- tightest tier"),
]
_TIGHTEST = _CAP_TIERS[-1][1]


def _cap_for(account: str) -> str:
    low = account.lower()
    for needles, cap in _CAP_TIERS:
        if any(n in low for n in needles):
            return cap
    return "(no explicit tier -- falls back to general rules)"


def _total_row_values(df):
    """(period_label, value) pairs for the account's total row -- the same
    row build_trend_summary focuses on. Excludes INTERNAL_ROW_KEY, which
    is a bookkeeping column (the sheet row index), not a period."""
    period_cols = [c for c in df.columns[1:]
                   if str(c) != INTERNAL_ROW_KEY and not str(c).endswith("_formatted")]
    row_types = df.attrs.get("row_types_by_description") or {}
    desc_col = df.columns[0]
    total_idx = None
    for idx, row in df.iterrows():
        if str(row_types.get(str(row[desc_col]), "")).lower() in ("total", "subtotal"):
            total_idx = idx
    if total_idx is None:
        # No labelled total -- sum the detail rows instead.
        return [(str(c), float(df[c].fillna(0).sum())) for c in period_cols]
    return [(str(c), float(df.loc[total_idx, c] or 0)) for c in period_cols]


def _pct(prev, curr):
    if prev == 0:
        return None
    return (curr - prev) / abs(prev) * 100


def _fmt_pct(p):
    return "n/a (from nil)" if p is None else f"{p:+,.1f}%"


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("path", help="path to the databook .xlsx")
    ap.add_argument("--entity", default=None, help="entity name, as you'd type it in the app")
    ap.add_argument("--sheet", default=None, help="specific sheet, if the app asks you to pick one")
    ap.add_argument("--account", default=None, help="only report this one account")
    ap.add_argument("--threshold", type=float, default=50.0,
                     help="movement %% above which an account is called high-variance (default 50)")
    ap.add_argument("--list-sheets", action="store_true",
                     help="just list this workbook's sheets and exit -- use this first when you "
                          "don't know what to pass for --entity / --financials-sheet")
    ap.add_argument("--financials-from", default=None,
                     help="path to a separate roll-up/master workbook holding the Financials "
                          "sheet, when this entity's own file doesn't carry one")
    ap.add_argument("--financials-sheet", default=None,
                     help="name of the Financials sheet to source from (needed when it lives in "
                          "a master/roll-up sheet rather than a per-entity tab)")
    args = ap.parse_args()

    if args.list_sheets:
        from openpyxl import load_workbook
        wb = load_workbook(args.path, read_only=True)
        print(f"{len(wb.sheetnames)} sheet(s) in {args.path!r}:")
        for name in wb.sheetnames:
            print(f"  {name}")
        return 0

    if not args.entity:
        print("❌ --entity is required (or use --list-sheets to see what's in this file first).")
        return 1

    print(f"Loading {args.path!r} (entity={args.entity!r})...")
    result = process_workbook_data(temp_path=args.path, entity_name=args.entity,
                                    selected_sheet=args.sheet,
                                    financials_from=args.financials_from,
                                    financials_sheet=args.financials_sheet)
    dfs = result["dfs"]
    print(f"{len(dfs)} account(s) processed. Language detected: {result.get('language')}\n")

    is_accounts = {}
    for key, df in dfs.items():
        integrity = df.attrs.get("integrity") or {}
        if str(integrity.get("statement_type", "")).upper() == "IS":
            is_accounts[key] = df
    if not is_accounts:
        print("❌ No IS accounts found -- is this a BS-only databook, or did the entity name not match?")
        return 1
    print(f"{len(is_accounts)} income-statement account(s): {', '.join(sorted(is_accounts))}\n")

    flagged, revenue_growth, expense_growth = [], None, {}

    for key in sorted(is_accounts):
        if args.account and key != args.account:
            continue
        df = is_accounts[key]
        integrity = df.attrs.get("integrity") or {}
        months = integrity.get("annualization_months")
        series = _total_row_values(df)
        if len(series) < 2:
            continue

        notes = df.attrs.get("supporting_notes") or []
        rhs = df.attrs.get("adjacent_detail_rows") or []
        linked = df.attrs.get("table_linked_remarks") or []
        has_expl = bool(notes or rhs or linked)

        print("=" * 78)
        print(f"{key}")
        print("=" * 78)
        print(f"  periods ({len(series)}):")
        for i, (p, v) in enumerate(series):
            tail = ""
            if i == len(series) - 1 and months and 0 < months < 12:
                tail = f"   <-- PARTIAL PERIOD ({months} month(s)); not comparable to a full year as-is"
            print(f"    {p:14s} {v:>18,.2f}{tail}")

        # Compare the last two FULL periods -- comparing a 1-month tail
        # against a full year would report a ~-92% "collapse" that is
        # purely a period-length artefact, not a real movement.
        full = series[:-1] if (months and 0 < months < 12) else series
        print(f"\n  period-over-period movement (full periods only):")
        biggest = 0.0          # largest real measurable % move
        from_nil = False       # a nil -> non-nil start, which has no meaningful %
        for (p0, v0), (p1, v1) in zip(full, full[1:]):
            p = _pct(v0, v1)
            mark = ""
            if p is not None and abs(p) >= args.threshold:
                mark = "   <-- HIGH VARIANCE"
                biggest = max(biggest, abs(p))
            elif p is None and v1 != 0:
                # Commonly just the entity commencing operations (a
                # pre-operational zero year), NOT a swing needing
                # explanation -- tracked separately so it never gets
                # reported as if it were a measured percentage move.
                mark = "   <-- from nil (new activity, no meaningful %)"
                from_nil = True
            print(f"    {p0} -> {p1}: {v0:>16,.2f} -> {v1:>16,.2f}  {_fmt_pct(p)}{mark}")

        low = key.lower()
        if any(n in low for n in ("operating income", "revenue")) and "non-operating" not in low:
            if len(full) >= 2:
                revenue_growth = _pct(full[-2][1], full[-1][1])
        elif any(n in low for n in ("expense", "cost", "cogs", "tax")):
            if len(full) >= 2:
                expense_growth[key] = _pct(full[-2][1], full[-1][1])

        cap = _cap_for(key)
        print(f"\n  explanatory material available:")
        print(f"    supporting notes    : {len(notes)}")
        print(f"    RHS remark rows     : {len(rhs)}")
        print(f"    table-linked remarks: {len(linked)}")
        print(f"  prompt length cap for this account type:\n    {cap}")

        if biggest >= args.threshold and cap == _TIGHTEST and has_expl:
            print(f"\n  ⚠️  FLAGGED: moved {biggest:,.0f}% and HAS explanatory material,")
            print(f"      but sits in the tightest cap tier -- current rules leave no room")
            print(f"      to use that material, and 'delete any driver not explicitly stated'")
            print(f"      removes what little analysis survives.")
            flagged.append((key, biggest, cap))
        elif biggest >= args.threshold and not has_expl:
            print(f"\n  ℹ️  moved {biggest:,.0f}% but NO remarks/notes exist to explain it --")
            print(f"      deeper analysis here would have to be invented, so the current")
            print(f"      'facts only' behaviour is arguably correct for this account.")
        elif from_nil:
            print(f"\n  ·  starts from nil (entity likely pre-operational in the first period);")
            print(f"     largest measurable move afterwards is {biggest:,.1f}% -- below the "
                  f"{args.threshold:,.0f}% threshold.")
        print()

    print("=" * 78)
    print("CROSS-ACCOUNT: revenue vs expense growth asymmetry")
    print("=" * 78)
    print("(no single account's commentary can currently observe this -- each is")
    print(" generated in isolation from its own DataFrame)\n")
    if revenue_growth is None:
        print("  Could not determine revenue growth (no operating-income account matched).")
    else:
        print(f"  revenue growth (latest full period): {_fmt_pct(revenue_growth)}\n")
        if abs(revenue_growth) >= 200:
            print("  ⚠️  Revenue moved >200%, which usually means the entity was still")
            print("      ramping up rather than trading at a steady state. Against a")
            print("      baseline that large, essentially EVERY expense line reads as")
            print("      'asymmetric' -- treat the gaps below as not meaningful here, and")
            print("      judge this check on a stabilised entity instead.\n")
        for key, g in sorted(expense_growth.items(), key=lambda kv: -(abs(kv[1]) if kv[1] else 0)):
            if g is None:
                print(f"    {key:34s} n/a (from nil)")
                continue
            gap = g - revenue_growth
            flag = "   <-- ASYMMETRIC" if abs(gap) >= args.threshold else ""
            print(f"    {key:34s} {g:+9,.1f}%   vs revenue: {gap:+9,.1f} pts{flag}")

    print("\n" + "=" * 78)
    print(f"SUMMARY: {len(flagged)} account(s) where the data warrants analysis the current")
    print("         rules structurally prevent")
    print("=" * 78)
    for key, mv, _cap in flagged:
        print(f"  - {key} (moved {mv:,.0f}%)")
    if not flagged:
        print("  (none -- on this databook the current caps aren't the binding constraint)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
