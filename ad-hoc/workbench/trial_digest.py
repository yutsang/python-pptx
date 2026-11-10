"""What a trial's review file says in a screenful: per arm, accounts accepted and cost, the accounts not accepted, and every
sentence the deck marks with why. No model call; free. entity_trial.py prints it at the end of a run; this reads any trial.

    python ad-hoc/workbench/trial_digest.py [--trial <trial_quality_*.md>] [--width 60]

The review file holds every paragraph, citation and form and runs to several hundred KB; pasting it back cost more than it
showed. The digest is what is needed to find a wrong refusal or a needless mark: the sentence (cut to --width characters) and
the reason. It holds client text, like the file, so it stays in the chat it is pasted to.
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Dict, List

ROOT = Path(__file__).resolve().parents[2]
TRIAL_DIR = ROOT / "fdd_utils" / "output" / "agent_trial"
MAX_MARKS = 25


def _read(path: Path) -> str:
    raw = path.read_bytes()
    try:
        return raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        return raw.decode("gbk", errors="replace")


def digest(text: str, width: int = 60) -> List[str]:
    arms = re.search(r"Arms: ([^.\n]+)\.", text)
    names = [a.strip() for a in arms.group(1).split(",")] if arms else []
    stamp = re.search(r"# Trial (\S+)", text)
    code = re.search(r"Code (\w+)\.", text)
    per: Dict[str, Dict] = {a: {"done": 0, "cases": 0, "cost": 0.0, "open": [], "marks": []} for a in names}
    for block in re.split(r"\n## case \d+: ", text)[1:]:
        account = block.split(" (", 1)[0].strip()
        for seg in re.split(r"\n\*\*(?=\w+\*\*: end `)", block)[1:]:
            m = re.match(r"(\w+)\*\*: end `([^`]+)`", seg)
            if not m or m.group(1) not in per:
                continue
            arm, end = m.group(1), m.group(2)
            row = per[arm]
            row["cases"] += 1
            cost = re.search(r", ([\d.]+) yuan, ", seg.split("\n", 1)[0])
            row["cost"] += float(cost.group(1)) if cost else 0.0
            if end == "done":
                row["done"] += 1
            else:
                row["open"].append(f"{account}({end})")
            for n, colour, sentence, why in re.findall(r"\n- \[(\d+)\] (RED|ORANGE): (.*?)\n\s+why: (.*?)(?=\n)", seg):
                row["marks"].append((colour[0], account, n, sentence.strip(), why.strip()))
    out = [f"DIGEST trial {stamp.group(1) if stamp else '?'} code {code.group(1) if code else '?'} "
           "(paste this back; the review file stays local)"]
    for arm in names:
        row = per[arm]
        red = sum(1 for x in row["marks"] if x[0] == "R")
        by_review = sum(1 for x in row["marks"] if x[4].startswith("review_objection"))
        out.append(f"{arm}: {row['done']}/{row['cases']} accepted, {row['cost']:.2f} yuan, marks {red} red {len(row['marks']) - red} orange"
                   + (f" ({by_review} by the reviewer)" if by_review else "")
                   + (f" | not accepted: {', '.join(row['open'])}" if row["open"] else ""))
        for colour, account, n, sentence, why in row["marks"][:MAX_MARKS]:
            cut = sentence if len(sentence) <= width else sentence[:width] + "…"
            out.append(f"  {colour} {account}[{n}] {why[:110]}")
            out.append(f"      {cut}")
        if len(row["marks"]) > MAX_MARKS:
            out.append(f"  ... {len(row['marks']) - MAX_MARKS} more marks in the review file")
    return out


def newest() -> Path:
    found = sorted(TRIAL_DIR.glob("trial_quality_*.md"), key=lambda f: f.stat().st_mtime)
    if not found:
        raise SystemExit("no trial_quality_*.md in fdd_utils/output/agent_trial; pass --trial")
    return found[-1]


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--trial", default="", help="trial_quality_<stamp>.md (default: the newest)")
    ap.add_argument("--width", type=int, default=60, help="characters of each marked sentence to show")
    args = ap.parse_args()
    path = Path(args.trial) if args.trial else newest()
    print("\n".join(digest(_read(path), args.width)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
