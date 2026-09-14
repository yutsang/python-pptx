"""What the deterministic house-style pass actually removed from ONE run.

Counting nil mentions in a fresh deck and comparing against a previous deck
measures nothing: the model rewrites every bullet each run, so the two counts
describe two different texts. A real comparison came out 8 then 10 and told us
only that the second run's commentary happened to contain more of them.

This reads a finished run's OWN commentary out of results.yml, applies the
payload pass to it, and reports the difference on that single fixed text.

    python ad-hoc/databook-probes/probe_house_style_delta.py                  # newest run
    python ad-hoc/databook-probes/probe_house_style_delta.py --run 20260914_101530
    python ad-hoc/databook-probes/probe_house_style_delta.py --show           # print each rewrite

Free, no model calls, seconds. REMAINING lists the sentences a nil mention survived in. Do NOT assume they are
known limits: the first run of this probe filed three as unpairable when the
sentence named its periods perfectly well and the rewrite was simply blind to
the 「人民币」 in front of the figures. Read them.
"""
from __future__ import annotations

import argparse
import glob
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from fdd_utils.pptx.payloads import (  # noqa: E402
    _normalize_slide_commentary_text, _ANY_FRAME, _BARE_SERIES, _SERIES_AMOUNT,
)

LOG_ROOT = os.path.join("fdd_utils", "logs")

#: A nil amount only counts when the zero IS the amount -- not the last digit
#: of 「194.0万元」. Getting this wrong once made the rounding fix look like it
#: had ADDED nil mentions.
NIL = re.compile(r"(?<![\d.])-?0(?:\.0+)?\s*(?:万元|亿元|元)")
OVERPRECISE = re.compile(r"\d\.\d{2,}\s*万元|\d\.\d{3,}\s*亿元")


def _newest_run() -> str:
    runs = sorted(glob.glob(os.path.join(LOG_ROOT, "run_*")))
    if not runs:
        sys.exit("no runs under %s -- run inspect_databook.py --run-ai first" % LOG_ROOT)
    return runs[-1]


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--run", default=None, help="run id or folder (default: newest)")
    ap.add_argument("--show", action="store_true", help="print every sentence that changed")
    args = ap.parse_args()

    folder = args.run if (args.run and os.path.isdir(args.run)) else (
        os.path.join(LOG_ROOT, "run_%s" % str(args.run).replace("run_", "")) if args.run else _newest_run())
    if not os.path.isdir(folder):
        sys.exit("no such run: %s" % folder)

    import yaml
    path = os.path.join(folder, "results.yml")
    if not os.path.exists(path):
        sys.exit("no results.yml in %s" % folder)
    with open(path, encoding="utf-8") as fh:
        results = yaml.safe_load(fh) or {}

    rows, tot_b, tot_a, tot_pb, tot_pa = [], 0, 0, 0, 0
    for key, record in sorted(results.items()):
        if key.startswith("__") or not isinstance(record, dict):
            continue
        before = ""
        for field in ("final", "agent_4_validation", "agent_2_content", "agent_1_content"):
            value = record.get(field)
            if isinstance(value, dict):
                value = value.get("content") or value.get("text")
            if isinstance(value, str) and value.strip():
                before = value
                break
        if not before.strip():
            continue
        after = _normalize_slide_commentary_text(before)
        nb, na = len(NIL.findall(before)), len(NIL.findall(after))
        pb, pa = len(OVERPRECISE.findall(before)), len(OVERPRECISE.findall(after))
        tot_b, tot_a, tot_pb, tot_pa = tot_b + nb, tot_a + na, tot_pb + pb, tot_pa + pa
        if nb or pb:
            rows.append((key, nb, na, pb, pa, before, after))

    # `rows` holds only the accounts that HAD something to fix -- calling that
    # "accounts with commentary" overstated the corpus every time it printed.
    print("RUN %s   %d account(s) carried something to fix\n" % (os.path.basename(folder), len(rows)))
    print("%-12s %-14s %s" % ("account", "nil in list", "over-precise 万元/亿元"))
    print("-" * 62)
    for key, nb, na, pb, pa, _b, _a in rows:
        print("%-12s %2d -> %-9d %d -> %d%s"
              % (key, nb, na, pb, pa, "" if na == 0 else "   <- see REMAINING"))
    print("-" * 62)
    print("%-12s %2d -> %-9d %d -> %d" % ("TOTAL", tot_b, tot_a, tot_pb, tot_pa))

    remaining = [(k, b, a) for k, nb, na, _pb, _pa, b, a in rows if na]
    if remaining:
        print("\nREMAINING -- read these; a survivor is a miss until shown otherwise.")
        print("Each one prints the pairing decision, so the cause needs no guessing:")
        for key, _b, after in remaining:
            for sentence in re.split(r"[。；;]", after):
                if not NIL.search(sentence):
                    continue
                print("\n  [%s] %s" % (key, sentence.strip()[:160]))
                # WHY it was not paired: the frame in force, and the two counts.
                at = after.index(sentence)
                frames = [m.group(0) for m in _ANY_FRAME.finditer(after[:at + len(sentence)])]
                series = [m for m in _BARE_SERIES.finditer(sentence)]
                amounts = len(_SERIES_AMOUNT.findall(series[0].group(1))) if series else 0
                if not frames:
                    print("        frame in force: NONE found anywhere before it"
                          "  ->  %d amount(s) cannot be paired" % amounts)
                else:
                    periods = [p for p in re.split(r"[、及和]", frames[-1]) if p.strip()]
                    print("        frame in force: %s  (%d period(s))" % (frames[-1], len(periods)))
                    print("        this series:    %d amount(s)  ->  %s"
                          % (amounts, "counts match, so this is a MISS"
                             if amounts == len(periods) else "counts differ, correctly refused"))
    if args.show:
        print("\nREWRITES")
        for key, nb, na, _pb, _pa, b, a in rows:
            if a != b:
                print("\n  [%s]\n    BEFORE %s\n    AFTER  %s" % (key, b[:300], a[:300]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
