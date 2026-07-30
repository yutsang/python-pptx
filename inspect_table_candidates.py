#!/usr/bin/env python3
"""Which databook tables are CANDIDATES for the report, by tie-out.

The rule this implements is general, not a hardcoded account list: any block
found in a sheet whose figures agree with that account's own indicative-
adjusted totals is a candidate, because agreeing with the reported figures is
what makes a breakdown reportable. A fee-rate workpaper, a rollforward or a
bank-account listing does not tie, so it never becomes a candidate however
report-like it looks.

Per account this shows whether a candidate exists, which periods it ties on,
the difference where it does not, and therefore whether it should be
considered. It stops at reporting -- deciding to include a table is an
editorial call, and a real reference deck narrates one account (营业收入)
whose table ties perfectly, so tie-out identifies candidates rather than
settling the question.

Read-only.

Usage:
    python inspect_table_candidates.py "for_test/x.xlsx"
    python inspect_table_candidates.py "for_test/x.xlsx" --account 管理费用 --show-values
"""
import argparse
import sys
import warnings

warnings.filterwarnings("ignore")
sys.path.insert(0, ".")

from fdd_utils.workbook import process_workbook_data


def _norm_key(label):
    """Column labels carry a stage or an 'annualised' suffix; the bare date is
    what a candidate block's periods can be compared against."""
    import re
    m = re.search(r"(\d{4}-\d{2}-\d{2})", str(label))
    return m.group(1) if m else str(label).strip()


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("path")
    ap.add_argument("--entity", default="x", help="entity name (soft filter for most files)")
    ap.add_argument("--sheet", default=None)
    ap.add_argument("--account", default=None, help="only this account")
    ap.add_argument("--show-values", action="store_true",
                     help="print every component's figures, not just the labels")
    ap.add_argument("--tolerance", type=float, default=0.02,
                     help="relative difference treated as a tie (default 2%%)")
    args = ap.parse_args()

    print(f"Loading {args.path!r}...")
    result = process_workbook_data(temp_path=args.path, entity_name=args.entity,
                                    selected_sheet=args.sheet)
    dfs = result["dfs"]
    print(f"{len(dfs)} account(s). Language: {result.get('language')}\n")

    candidates, no_table, not_tied = [], [], []

    for key in sorted(dfs):
        if args.account and key != args.account:
            continue
        df = dfs[key]
        attrs = df.attrs or {}
        table = attrs.get("presentation_detail_table")
        integrity = attrs.get("integrity") or {}
        stmt = str(integrity.get("statement_type") or "?")
        # projection_totals_by_date only holds the projection column (plus its
        # annualised twin), so tying against it alone checks ONE period. The
        # account's own total row covers every period, and since the parsed df
        # is already the PREFERRED_STAGE ("Indicative adjusted") view, those
        # ARE the indicative-adjusted totals -- which is what a reportable
        # breakdown has to agree with.
        norm_totals = {}
        for label, value in (attrs.get("projection_totals_by_date") or {}).items():
            k = _norm_key(label)
            if k and isinstance(value, (int, float)):
                norm_totals.setdefault(k, value)
        row_types = attrs.get("row_types_by_description") or {}
        desc_col = df.columns[0]
        total_idx = None
        for idx, row in df.iterrows():
            if str(row_types.get(str(row[desc_col]), "")).lower() in ("total", "subtotal"):
                total_idx = idx
        period_cols = [c for c in list(df.columns)[1:]
                       if str(c) != "__source_row_idx" and not str(c).endswith("_formatted")]
        for c in period_cols:
            k = _norm_key(c)
            if k in norm_totals:
                continue
            try:
                value = (float(df.loc[total_idx, c]) if total_idx is not None
                         else float(df[c].fillna(0).sum()))
            except Exception:
                continue
            norm_totals[k] = value

        if not table:
            no_table.append((key, stmt))
            continue

        rows = table["rows"]
        total_row = table.get("total_row")
        periods = table.get("periods") or []

        print("=" * 78)
        print(f"{key}   [{stmt}]")
        print("=" * 78)
        if table.get("title"):
            print(f"  block title : {table['title']}")
        print(f"  located at  : sheet row {table['header_row'] + 1}, "
              f"label column index {table['label_col']}")
        print(f"  components  : {len(rows)}  {[r['label'] for r in rows][:8]}")

        print(f"\n  TIE-OUT against this account's indicative-adjusted totals:")
        tied, differed = [], []
        for period in periods:
            account_total = norm_totals.get(period)
            if account_total is None:
                print(f"    {period}  (no indicative-adjusted total for this period)")
                continue
            if total_row and period in (total_row.get("values") or {}):
                block_total = total_row["values"][period]
                src = "block total row"
            else:
                block_total = sum(r["values"].get(period, 0.0) for r in rows)
                src = "sum of components"
            diff = block_total - account_total
            if abs(account_total) < 1e-6:
                # Both nil is agreement, not a failure. Treating a zero
                # denominator as "cannot tie" wrongly downgraded a candidate
                # whose account genuinely had no balance that period.
                rel = None
                ok = abs(block_total) < max(1.0, abs(account_total))
            else:
                rel = abs(diff) / abs(account_total)
                ok = rel <= args.tolerance
            (tied if ok else differed).append(period)
            rel_s = f"{rel * 100:6.2f}%" if rel is not None else ("both nil" if ok else "   n/a")
            print(f"    {period}  block {block_total:>14,.1f} ({src})  "
                  f"account {account_total:>14,.1f}  diff {diff:>+12,.1f}  {rel_s}  "
                  f"{'✅' if ok else '❌'}")

        verdict_tied = bool(tied)
        print(f"\n  VERDICT: ", end="")
        if verdict_tied and not differed:
            print("✅ CANDIDATE -- ties on every period with a comparable total")
            candidates.append((key, stmt, len(rows), "all periods"))
        elif verdict_tied:
            print(f"✅ CANDIDATE -- ties on {len(tied)}, differs on {len(differed)} "
                  f"({', '.join(differed)})")
            print(f"           A period that differs is usually the annualised column, which a")
            print(f"           breakdown table does not carry -- check before dismissing it.")
            candidates.append((key, stmt, len(rows), f"{len(tied)}/{len(tied) + len(differed)} periods"))
        else:
            print("❌ NOT a candidate -- ties on no period, so this block is not this")
            print("           account's reported breakdown")
            not_tied.append((key, stmt))

        if args.show_values:
            print(f"\n  COMPONENT FIGURES:")
            for r in rows:
                vals = "  ".join(f"{p}={r['values'].get(p, 0):,.1f}" for p in periods)
                print(f"    {r['label'][:22]:24s} {vals}")
            if total_row:
                vals = "  ".join(f"{p}={total_row['values'].get(p, 0):,.1f}" for p in periods)
                print(f"    {'合计 / total':24s} {vals}")
        print()

    print("=" * 78)
    print("SUMMARY")
    print("=" * 78)
    print(f"  candidates (tie to indicative-adjusted) : {len(candidates)}")
    for key, stmt, n, how in candidates:
        print(f"    ✅ {key:22s} [{stmt}]  {n} components, ties on {how}")
    if not_tied:
        print(f"\n  blocks found but NOT tying               : {len(not_tied)}")
        for key, stmt in not_tied:
            print(f"    ❌ {key:22s} [{stmt}]")
    print(f"\n  accounts with no candidate block         : {len(no_table)}")
    if no_table and not args.account:
        print(f"    {', '.join(k for k, _s in no_table[:14])}"
              + (" ..." if len(no_table) > 14 else ""))
    print("\n  Tie-out identifies candidates; it does not decide inclusion. A real")
    print("  reference deck narrates 营业收入 even though its block ties perfectly, so")
    print("  the editorial call stays with whoever writes the report.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
