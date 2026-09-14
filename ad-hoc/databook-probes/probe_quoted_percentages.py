"""Does every CHANGE percentage a bullet quotes match one the pipeline computed?

prompts.yml:806 tells the model outright not to derive a movement percentage
itself -- 「引用变动幅度时直接使用 percent_change 提供的数值，不要自行由两期金额
相除得出」. If that instruction is followed, every change percentage in the deck
should be findable among the percentages the pipeline handed over, and a
percentage that matches none is one the model invented.

That would make the check a GROUNDING question ("is this figure ours?") rather
than an arithmetic one ("do these two numbers divide to that?"), which matters:
an adversarial sweep found that every sentence-local arithmetic formulation
misfires on house-mandated text -- the prompt's own 首句强制 nil opening prints
the BASE after the decline token, 占比/利率/出租率/残值率 put an unrelated
percentage in the same sentence, and a multi-leg trajectory carries one
percentage for one leg among three pairs.

This measures whether the grounding formulation is viable before anything is
built on it. It does NOT change behaviour.

    python ad-hoc/databook-probes/probe_quoted_percentages.py <databook.xlsx> --run <run folder>

The workbook supplies the computed percentages (free, no AI); the run folder
supplies the commentary that shipped. They must be the same entity.
"""
from __future__ import annotations

import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

#: A percentage that belongs to a MOVEMENT, not a share, a rate or a ratio. The
#: exclusions are the shapes the sweep measured as normal house style: 占比 /
#: 出租率 / 利率 / 残值率 / 计提比例 / 持股.
_CHANGE_PCT = re.compile(
    r"(?:增长|增幅|上升|增加|提升|下降|降幅|减少|下滑|回落)[^。；;，,]{0,12}?"
    r"(?:约|达|至)?\s*([\d,]+(?:\.\d+)?)\s*%"
)
_SHARE_PCT = re.compile(r"(?:占|比重|比例|率为|利率|残值率|出租率|持股|计提)[^。；;]{0,10}?([\d,]+(?:\.\d+)?)\s*%")


def _computed_percentages(df) -> set:
    """Every percentage this account's own frame handed the model."""
    out = set()
    attrs = getattr(df, "attrs", None) or {}
    for movement in (attrs.get("significant_movements") or []):
        if not isinstance(movement, dict):
            continue
        pct = movement.get("percent_change")
        if isinstance(pct, (int, float)):
            out.add(round(abs(float(pct)), 1))
        for key in ("percent_of_total_change", "share_of_total"):
            value = movement.get(key)
            if isinstance(value, (int, float)):
                out.add(round(abs(float(value)), 1))
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook")
    ap.add_argument("--run", required=True, help="run folder whose results.yml holds the shipped text")
    ap.add_argument("--tol", type=float, default=0.5, help="pp tolerance (default 0.5, matching _DIR_PCT_TOL_PP)")
    args = ap.parse_args()

    import yaml
    from fdd_utils.workbook import process_workbook_data

    with open(os.path.join(args.run, "results.yml"), encoding="utf-8") as fh:
        results = yaml.safe_load(fh) or {}
    state = process_workbook_data(temp_path=args.workbook, entity_name="", selected_sheet=None)
    dfs = state["dfs"]

    matched = unmatched = shares = 0
    misses = []
    for key, record in sorted(results.items()):
        if key.startswith("__") or not isinstance(record, dict) or key not in dfs:
            continue
        text = ""
        for field in ("final", "agent_4_validation", "agent_2_content", "agent_1_content"):
            value = record.get(field)
            if isinstance(value, dict):
                value = value.get("content") or value.get("text")
            if isinstance(value, str) and value.strip():
                text = value
                break
        if not text:
            continue
        computed = _computed_percentages(dfs[key])
        share_spans = {m.span(1) for m in _SHARE_PCT.finditer(text)}
        for m in _CHANGE_PCT.finditer(text):
            if m.span(1) in share_spans:       # a share wearing a movement verb
                shares += 1
                continue
            quoted = round(float(m.group(1).replace(",", "")), 1)
            if any(abs(quoted - c) <= args.tol for c in computed):
                matched += 1
            else:
                unmatched += 1
                lo = max(0, m.start() - 40)
                misses.append((key, quoted, text[lo:m.end() + 20].replace("\n", " "),
                               sorted(computed)[:6]))

    total = matched + unmatched
    print("\nCHANGE percentages quoted in the deck: %d" % total)
    if total:
        print("  matched a computed one (+/-%.1fpp): %d  (%.0f%%)" % (args.tol, matched, 100.0 * matched / total))
        print("  matched NOTHING computed          : %d  (%.0f%%)" % (unmatched, 100.0 * unmatched / total))
    print("  excluded as share/rate/ratio        : %d" % shares)
    if misses:
        print("\nQUOTED BUT NOT COMPUTED -- each of these would be flagged by a grounding check:")
        for key, quoted, context, computed in misses[:25]:
            print("\n  [%s] quoted %.1f%%" % (key, quoted))
            print("      ...%s..." % context.strip())
            print("      this account's computed percentages: %s" % (computed or "NONE"))
    print("\nRead every miss before concluding. A high miss rate means the grounding")
    print("formulation is NOT viable and the percentage must be checked some other way.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
