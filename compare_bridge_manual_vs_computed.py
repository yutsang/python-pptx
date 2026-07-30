#!/usr/bin/env python3
"""Puts a hand-built '<entity>-量价桥图' tab next to what this tool computes
from that entity's raw AB- tab, item by item, and localises any difference.

Needed because "the bridge is off by 60" cannot be answered from either
side alone. The manual tab and the computed decomposition use the same
algebra, so a gap comes from one of three places, and this separates them:

  1. the tab's own internal check doesn't close (the manual work has a
     residual before any comparison with us);
  2. an INPUT differs -- unit rent, area or days for one phase/period, which
     is what a stale cross-sheet reference or a different averaging window
     produces;
  3. only the FACTOR SPLIT differs while the start and end totals agree --
     the same total attributed differently between price, area and days.

Case 3 matters: it means neither side is wrong about the movement, only
about its attribution. That is the case a user sees as "the middle
breakdown is off but the rest is fine".

Read-only.

Usage:
    python compare_bridge_manual_vs_computed.py "databooks/x.xlsx" --manual "SJ-量价桥图" --raw AB-SJ
    python compare_bridge_manual_vs_computed.py "databooks/x.xlsx" --manual "成都-量价桥图" --raw AB-CD
"""
import argparse
import re
import sys

sys.path.insert(0, ".")

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from fdd_utils.bridge_chart_prototype import find_bridge_blocks
from fdd_utils.generate_bridge_waterfall_batch import build_bridges_for_ab_tab

# Longest marker first per factor: the days label is '运营天数变动', so
# splitting on the bare '天数' would leave '运营' stuck on the phase name and
# a manual '干仓运营天数增加' would not line up with a computed one whose
# phase came from a different source.
_FACTOR_WORDS = (
    ("days", ("运营天数", "運營天數", "运营天", "運營天", "天数", "天數")),
    ("price", ("单价", "單價")),
    ("area", ("出租率", "面积", "面積")),
)


def _classify(label: str):
    """(phase, factor) from a bridge item label like '干仓单价增加'."""
    text = str(label or "").strip()
    for factor, words in _FACTOR_WORDS:
        for w in words:
            if w in text:
                phase = text.split(w)[0].strip()
                # Trailing connective left by a partial marker match.
                phase = re.sub(r"(运营|運營)$", "", phase).strip()
                return phase, factor
    return text, None


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("path")
    ap.add_argument("--manual", required=True, help="the hand-built bridge tab, e.g. 'SJ-量价桥图'")
    ap.add_argument("--raw", required=True, help="that entity's raw data tab, e.g. AB-SJ")
    ap.add_argument("--tol", type=float, default=0.5, help="difference below which items are treated as equal")
    args = ap.parse_args()

    wb = load_workbook(args.path, data_only=True)
    for name in (args.manual, args.raw):
        if name not in wb.sheetnames:
            print(f"❌ sheet {name!r} not found. Available: {wb.sheetnames}")
            return 1

    manual_blocks = find_bridge_blocks(wb[args.manual])
    if not manual_blocks:
        print(f"❌ no Base/Change block found on {args.manual!r}.")
        return 1
    _blocks, computed = build_bridges_for_ab_tab(wb[args.raw], args.raw, log=lambda _m: None)
    if not computed:
        print(f"❌ {args.raw!r} produced no computed bridge.")
        return 1

    print(f"manual tab  : {args.manual!r}  ({len(manual_blocks)} block(s))")
    print(f"computed from: {args.raw!r}  ({len(computed)} transition(s))\n")

    for bi, mblock in enumerate(manual_blocks):
        start_label = mblock.items[0].label
        end_label = mblock.items[-1].label
        print("=" * 78)
        print(f"MANUAL BLOCK {bi + 1}: {start_label} -> {end_label}")
        print("=" * 78)

        # 1. Does the manual tab close against itself?
        m_start = mblock.items[0].value
        m_end = mblock.items[-1].value
        m_recon = m_start + sum(it.value for it in mblock.items[1:-1])
        m_resid = m_recon - m_end
        print(f"  [1] MANUAL TAB'S OWN CHECK")
        print(f"      start {m_start:>14,.2f}  + factors = {m_recon:>14,.2f}")
        print(f"      stated end                      {m_end:>14,.2f}")
        print(f"      residual                        {m_resid:>14,.2f}"
              f"   {'✅ closes' if abs(m_resid) <= args.tol else '❌ DOES NOT CLOSE'}")
        if abs(m_resid) > args.tol:
            print(f"      => the hand-built tab already carries this residual BEFORE any")
            print(f"         comparison with us. Fix it there first; comparing a tab that")
            print(f"         does not close against ours cannot tell you which side is right.")

        # 2. Match to the computed transition with the closest end total.
        best = min(computed, key=lambda r: abs(r.bridge.items[-1].value - m_end))
        c_start = best.bridge.items[0].value
        c_end = best.bridge.items[-1].value
        print(f"\n  [2] TOTALS vs computed ({best.year_a} -> "
              f"{'LTM' if best.is_ltm else best.year_b})")
        print(f"      start : manual {m_start:>13,.2f}   computed {c_start:>13,.2f}   "
              f"diff {m_start - c_start:>+12,.2f}")
        print(f"      end   : manual {m_end:>13,.2f}   computed {c_end:>13,.2f}   "
              f"diff {m_end - c_end:>+12,.2f}")
        totals_agree = abs(m_start - c_start) <= args.tol and abs(m_end - c_end) <= args.tol

        # 3. Factor-by-factor.
        m_items = {}
        for it in mblock.items[1:-1]:
            phase, factor = _classify(it.label)
            if factor:
                m_items[(phase, factor)] = (it.label, it.value)
        c_items = {}
        for it in best.bridge.items[1:-1]:
            phase, factor = _classify(it.label)
            if factor:
                c_items[(phase, factor)] = (it.label, it.value)

        print(f"\n  [3] FACTOR-BY-FACTOR")
        print(f"      {'phase':10s} {'factor':6s} {'manual':>14s} {'computed':>14s} {'diff':>12s}")
        print("      " + "-" * 60)
        keys = sorted(set(m_items) | set(c_items))
        total_abs_diff = 0.0
        biggest = None
        for k in keys:
            phase, factor = k
            mv = m_items.get(k, (None, None))[1]
            cv = c_items.get(k, (None, None))[1]
            if mv is None:
                print(f"      {phase:10s} {factor:6s} {'(absent)':>14s} {cv:>14,.2f}"
                      f" {'-':>12s}   <-- only in computed")
                continue
            if cv is None:
                print(f"      {phase:10s} {factor:6s} {mv:>14,.2f} {'(absent)':>14s}"
                      f" {'-':>12s}   <-- only in manual")
                continue
            d = mv - cv
            total_abs_diff += abs(d)
            if biggest is None or abs(d) > abs(biggest[2]):
                biggest = (phase, factor, d)
            mark = "" if abs(d) <= args.tol else "   <-- DIFFERS"
            print(f"      {phase:10s} {factor:6s} {mv:>14,.2f} {cv:>14,.2f} {d:>+12,.2f}{mark}")

        print(f"\n  [4] VERDICT")
        m_factor_sum = sum(v for _l, v in m_items.values())
        c_factor_sum = sum(v for _l, v in c_items.values())
        print(f"      sum of manual factors  : {m_factor_sum:>14,.2f}")
        print(f"      sum of computed factors: {c_factor_sum:>14,.2f}")
        print(f"      difference             : {m_factor_sum - c_factor_sum:>+14,.2f}")
        if totals_agree and abs(m_factor_sum - c_factor_sum) <= args.tol and total_abs_diff > args.tol:
            print(f"      => Totals AND the factor sum agree, but individual factors do not.")
            print(f"         Neither side is wrong about the movement -- only about how it is")
            print(f"         SPLIT between price, area and days. That comes from the order of")
            print(f"         sequential substitution, so check which order the manual tab uses")
            print(f"         (its own formulas) against price->area->days.")
        elif not totals_agree:
            print(f"      => The TOTALS differ, so this is an input problem, not attribution.")
            print(f"         Compare the inputs behind the disagreeing side -- unit rent, area")
            print(f"         and days per phase. A stale cross-sheet reference in the manual")
            print(f"         tab produces exactly this (one was already found on another tab).")
        elif biggest and abs(biggest[2]) > args.tol:
            print(f"      => Largest single gap: {biggest[0]} {biggest[1]} "
                  f"{biggest[2]:+,.2f}. Start there.")
        else:
            print(f"      => Manual and computed agree within ±{args.tol}.")
        print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
