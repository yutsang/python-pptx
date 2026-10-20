"""Read the scorecard every --run-ai appends, and say how far from "most files at 90" we are.

    python ad-hoc/workbench/scorecard.py            # latest run of every file
    python ad-hoc/workbench/scorecard.py --history  # every run, oldest first

Free: reads fdd_utils/logs/scorecard.jsonl (gitignored, per machine). Each
line is one databook run: score = accounts the model wrote with no
hallucination flag, over accounts run. Highlighted inferences ("reasoning")
do not lower it; they are counted alongside so a run that went quiet by
saying less is visible.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

PATH = Path("fdd_utils/logs/scorecard.jsonl")


def main() -> int:
    if not PATH.exists():
        print(f"No scorecard yet: {PATH} is written by inspect_databook.py --run-ai.")
        return 1
    rows = [json.loads(line) for line in PATH.read_text(encoding="utf-8").splitlines() if line.strip()]
    if "--history" in sys.argv:
        shown = rows
    else:
        latest = {}
        for row in rows:
            latest[row["file"]] = row
        shown = sorted(latest.values(), key=lambda r: r["score"])
    # miss = high-severity analyst observations no commentary mentions (6c);
    # "-" on rows written before that check existed. Not part of the score.
    print("%-46s %6s %8s %6s %6s %6s %5s %7s %8s" % ("file", "score", "clean", "hallu", "degr", "infer", "miss",
                                                     "commit", "time"))
    for r in shown:
        miss = str(r["uncovered_high"]) if "uncovered_high" in r else "-"
        print("%-46s %6.0f %8s %6d %6d %6d %5s %7s %8s" % (
            r["file"][:46], r["score"], f"{r['clean']}/{r['accounts']}", len(r["hallucination_accounts"]),
            len(r["degraded_accounts"]), r["reasoning_flags"], miss, r.get("commit", ""), r["time"][5:16]))
        for key in r["hallucination_accounts"] + r["degraded_accounts"]:
            print("%-46s %s" % ("", "  review: " + key))
    if "--history" not in sys.argv and shown:
        scores = sorted(r["score"] for r in shown)
        at90 = sum(1 for s in scores if s >= 90)
        print(f"\n{len(shown)} file(s): {at90} at 90+ ({100 * at90 / len(shown):.0f}%), "
              f"median {scores[len(scores) // 2]:.0f}, lowest {scores[0]:.0f}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
