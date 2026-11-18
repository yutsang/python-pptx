"""Free; no model call, no workbook. The sentences the verifier refused in earlier paid runs, read again by the verifier as it is now.

    python ad-hoc/workbench/issue_cases.py replay [--dir fdd_utils/output/agent_issues] [--only open] [--width 160]
    python ad-hoc/workbench/issue_cases.py list   [--dir ...]

entity_trial.py writes one case per account that ended a run with a refused sentence (fdd_utils/agent/issues.py). `replay` is the
fix queue: which cases a change fixed, which are still refused and with which codes. It holds client sentences: it stays in the chat
it is pasted to.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from fdd_utils.agent import issues  # noqa: E402

DEFAULT_DIR = ROOT / "fdd_utils" / "output" / "agent_issues"


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("action", choices=["replay", "list"])
    ap.add_argument("--dir", default=str(DEFAULT_DIR))
    ap.add_argument("--only", choices=["open", "partly", "fixed"], default="", help="show only cases with this status")
    ap.add_argument("--width", type=int, default=160)
    args = ap.parse_args(argv)
    paths = issues.cases(Path(args.dir))
    if not paths:
        print(f"no cases under {args.dir}")
        return 0
    if args.action == "list":
        for p in paths:
            print(p.relative_to(args.dir))
        print(f"{len(paths)} case(s)")
        return 0
    results = [issues.replay(p) for p in paths]
    for r in results:
        if args.only and r["status"] != args.only:
            continue
        print(f"{r['status'].upper():8s} {r['entity']} / {r['account']}  (captured {r.get('captured')}, code {r.get('code')}): "
              f"{len(r['was'])} refused then, {len(r['now'])} now")
        for item in r["now"]:
            print(f"    [{','.join(item['codes'])}] {item['text'][:args.width]}")
            for m in item["messages"][:2]:
                print(f"        {m[:args.width]}")
    s = issues.summary(results)
    print(f"\nCASES {s['cases']}: " + ", ".join(f"{k} {v}" for k, v in sorted(s["by_status"].items())))
    if s["open_by_code"]:
        print("still refused, by code: " + ", ".join(f"{k} {v}" for k, v in s["open_by_code"].items()))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
