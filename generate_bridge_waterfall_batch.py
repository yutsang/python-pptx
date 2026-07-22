#!/usr/bin/env python3
"""Combines the raw-data bridge extractor (extract_bridge_from_raw.py) with
the waterfall chart builder (bridge_chart_prototype.py): computes the
price/occupancy/days factor decomposition per phase per year-transition
directly from a raw AB-* tab (no pre-built '<entity>-量价桥图' helper tab
needed), then renders a real PPTX waterfall chart -- across EVERY AB- tab
in the workbook in one pass, covering the entities that don't have a
manually-built bridge tab (16 of the 17 real ones, per this session's
17-tab scan -- only 成都/AB-CD has one).

Decomposition formula -- confirmed bit-exact against AB-CD's OWN real
bridge-tab formulas (2024->2025 transition, all 3 phases) this session,
sequential substitution in price -> area -> days order:
    price_effect = (price_B - price_A) * area_A * days_A / 1000
    area_effect  = price_B * (area_B - area_A) * days_A / 1000
    days_effect  = price_B * area_B * (days_B - days_A) / 1000
Each phase contributes 3 factor bars per year-transition; the bridge starts
at year_A's total revenue (summed across all phases) and ends at year_B's.

Usage:
    python generate_bridge_waterfall_batch.py "databooks/xx.xlsx"
        # every AB- tab, every consecutive year transition found, one PPTX
    python generate_bridge_waterfall_batch.py "databooks/xx.xlsx" --tab AB-CD --validate
        # cross-checks the 2024->2025 transition against AB-CD's own real
        # bridge-tab factor values (the "answer key")
"""
import argparse
import sys
from typing import Dict, List, Optional

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from extract_bridge_from_raw import find_phase_blocks, extract_annual_series, PhaseBlock
from inspect_ab_tabs_structure import find_labeled_rows
from bridge_chart_prototype import BridgeItem, BridgeBlock, build_waterfall_chart

from pptx import Presentation
from pptx.util import Inches

_FACTOR_SUFFIX = {"price": "单价变动", "area": "出租率/面积变动", "days": "运营天数变动"}


def decompose_transition(phase_label: str, series_a: Dict[str, float], series_b: Dict[str, float]) -> List[BridgeItem]:
    price_a, price_b = series_a["unit_rent"], series_b["unit_rent"]
    area_a, area_b = series_a["area"], series_b["area"]
    days_a, days_b = series_a["days"], series_b["days"]

    price_effect = (price_b - price_a) * area_a * days_a / 1000
    area_effect = price_b * (area_b - area_a) * days_a / 1000
    days_effect = price_b * area_b * (days_b - days_a) / 1000

    return [
        BridgeItem(label=f"{phase_label}{_FACTOR_SUFFIX['price']}", kind="delta", value=price_effect),
        BridgeItem(label=f"{phase_label}{_FACTOR_SUFFIX['area']}", kind="delta", value=area_effect),
        BridgeItem(label=f"{phase_label}{_FACTOR_SUFFIX['days']}", kind="delta", value=days_effect),
    ]


def find_year_days_rows(ws_values) -> Dict[str, Optional[int]]:
    labeled = find_labeled_rows(ws_values)
    year_row = days_row = None
    for r in sorted(labeled):
        cats = {cat for _, _, cat in labeled[r]}
        if year_row is None and "period_year" in cats:
            year_row = r
        if days_row is None and "period_days" in cats:
            days_row = r
    return {"year_row": year_row, "days_row": days_row}


def build_bridge_for_transition(ws_values, blocks: List[PhaseBlock], year_row: int, days_row: int,
                                 year_a: int, year_b: int, start_label: str, end_label: str) -> Optional[BridgeBlock]:
    all_series = [extract_annual_series(ws_values, b, year_row, days_row) for b in blocks]
    if any(year_a not in s or year_b not in s for s in all_series):
        return None

    total_a = sum(s[year_a]["revenue_k"] for s in all_series)
    total_b = sum(s[year_b]["revenue_k"] for s in all_series)

    items: List[BridgeItem] = [BridgeItem(label=start_label, kind="total", value=total_a)]
    for block, series in zip(blocks, all_series):
        items.extend(decompose_transition(block.label, series[year_a], series[year_b]))
    items.append(BridgeItem(label=end_label, kind="total", value=total_b))

    reconstructed = total_a + sum(it.value for it in items[1:-1])
    # A real client-data residual is expected here (confirmed in AB-CD's own
    # notes: revenue is "all-in" but leased area excludes some pallet-priced
    # space, so price*area*days never reconstructs revenue PERFECTLY) --
    # this is informational, not a reason to withhold the chart.
    check_ok = abs(reconstructed - total_b) < max(5.0, abs(total_b) * 0.02)
    return BridgeBlock(header_row=-1, label_col=-1, base_col=-1, change_col=-1, items=items, check_ok=check_ok)


# Hardcoded from the REAL '成都-量价桥图' bridge tab's own 2024->2025
# factor values (already confirmed this session against the real file) --
# the "answer key" for --validate.
_AB_CD_EXPECTED_FACTORS_2024_2025 = {
    "干仓": {"price": 2059.545458900859, "area": 1463.5320655100995, "days": 3298.960105589041},
    "综合楼": {"price": 66.9393853357857, "area": -55.45101747277201, "days": 29.103442136986306},
    "冷库": {"price": 0.0, "area": 0.0, "days": 3061.83285},
}


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("path", help="path to the databook .xlsx")
    ap.add_argument("--tab", default=None, help="only process this one AB- tab (default: every AB- tab)")
    ap.add_argument("--validate", action="store_true",
                     help="cross-check --tab AB-CD's 2024->2025 transition against its own real "
                          "bridge-tab factor values")
    ap.add_argument("--out", default="bridge_waterfall_batch_output.pptx", help="output pptx path")
    args = ap.parse_args()

    print(f"Loading {args.path!r}...")
    wb_values = load_workbook(args.path, data_only=True)
    all_sheets = wb_values.sheetnames
    ab_tabs = [s for s in all_sheets if s.startswith("AB-") or s.startswith("AB")]
    if args.tab:
        ab_tabs = [s for s in ab_tabs if s == args.tab]
        if not ab_tabs:
            print(f"❌ tab {args.tab!r} not found among AB- sheets.")
            return 1

    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    slides_added = 0
    all_ok = True
    for tab in ab_tabs:
        ws = wb_values[tab]
        yd = find_year_days_rows(ws)
        if not yd["year_row"] or not yd["days_row"]:
            print(f"--- {tab}: ⚠️ no Year/Days row found, skipping ---")
            continue
        blocks = find_phase_blocks(ws)
        if not blocks:
            print(f"--- {tab}: ⚠️ no phase blocks found, skipping ---")
            continue

        years = sorted({y for b in blocks
                         for y in extract_annual_series(ws, b, yd["year_row"], yd["days_row"]).keys()})
        transitions = list(zip(years, years[1:]))
        print(f"--- {tab}: {len(blocks)} phase(s) [{', '.join(b.label for b in blocks)}], "
              f"years {years} ---")

        for year_a, year_b in transitions:
            bridge = build_bridge_for_transition(
                ws, blocks, yd["year_row"], yd["days_row"], year_a, year_b,
                start_label=f"{year_a}年收入", end_label=f"{year_b}年收入",
            )
            if bridge is None:
                continue
            check_msg = "✅" if bridge.check_ok else "⚠️ residual > tolerance"
            print(f"    {year_a}->{year_b}: start={bridge.items[0].value:,.1f}k "
                  f"end={bridge.items[-1].value:,.1f}k  {check_msg}")

            if args.validate and tab == "AB-CD" and (year_a, year_b) == (2024, 2025):
                for block, phase_name in zip(blocks, ["干仓", "综合楼", "冷库"]):
                    exp = _AB_CD_EXPECTED_FACTORS_2024_2025.get(phase_name)
                    if not exp:
                        continue
                    got_items = [it for it in bridge.items if it.label.startswith(phase_name)]
                    got = {"price": got_items[0].value, "area": got_items[1].value, "days": got_items[2].value}
                    for factor in ("price", "area", "days"):
                        tol = max(0.5, abs(exp[factor]) * 0.005)
                        ok = abs(got[factor] - exp[factor]) <= tol
                        all_ok = all_ok and ok
                        status = "✅" if ok else "❌"
                        print(f"      {status} {phase_name} {factor}: got={got[factor]:.4f} expected={exp[factor]:.4f}")

            slide = prs.slides.add_slide(blank_layout)
            title = f"{tab}: {year_a}年→{year_b}年 收入量价桥图"
            build_waterfall_chart(
                slide, bridge, title,
                left=Inches(0.5), top=Inches(0.5), width=Inches(12.3), height=Inches(6.3),
            )
            slides_added += 1

    prs.save(args.out)
    print(f"\nSaved {slides_added} chart(s) to {args.out!r}.")
    if args.validate:
        print("✅ Factor decomposition matches AB-CD's own real values." if all_ok
              else "❌ MISMATCH -- do not trust this decomposition yet.")
        return 0 if all_ok else 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
