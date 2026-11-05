"""Print what each stage wrote for an account in the newest run, and why it was flagged.

A deck shows the text that survived; this shows the Generator's, the Auditor's
and the final text side by side, with every unsupported clause and its reason,
so a sentence that went wrong can be traced to the stage that wrote it.
Free: it only reads a run folder under fdd_utils/logs.

    python ad-hoc/workbench/show_run_text.py 营业收入 营业成本
    python ad-hoc/workbench/show_run_text.py --run 20261012_153000 应付账款

With no account named it lists the accounts in the run.
Output names accounts and quotes figures; do not paste it anywhere public.
"""
from __future__ import annotations

import glob
import os
import sys

import yaml


def main() -> int:
    args = [a for a in sys.argv[1:]]
    run = None
    if "--run" in args:
        i = args.index("--run")
        run = args[i + 1]
        del args[i:i + 2]
    root = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
                        "fdd_utils", "logs")
    folders = sorted(glob.glob(os.path.join(root, "run_*")))
    if run:
        folders = [f for f in folders if run in os.path.basename(f)]
    folders = [f for f in folders if os.path.exists(os.path.join(f, "results.yml"))]
    if not folders:
        print(f"no run with a results.yml under {root}")
        return 1
    folder = folders[-1]
    with open(os.path.join(folder, "results.yml"), encoding="utf-8") as handle:
        results = yaml.safe_load(handle) or {}
    print(f"run: {os.path.basename(folder)}")
    accounts = [k for k in results if not str(k).startswith("__") and isinstance(results[k], dict)]
    if not args:
        print("accounts: " + ", ".join(str(a) for a in accounts))
        return 0
    for name in args:
        record = results.get(name)
        if not isinstance(record, dict):
            print(f"\n=== {name}: not in this run")
            continue
        print(f"\n{'=' * 78}\n  {name}\n{'=' * 78}")
        for stage, label in (("subagent_1", "Generator"), ("subagent_2", "Auditor"),
                             ("subagent_4", "Validator"), ("final", "FINAL (what the deck gets)")):
            text = record.get(stage)
            if isinstance(text, dict):
                text = text.get("content") or text.get("final_content") or str(text)
            text = str(text or "").strip()
            print(f"\n--- {label}  ({len(text)} chars) ---\n{text if text else '(empty)'}")
        validation = record.get("agent_4_validation") or {}
        flagged = [r for r in (validation.get("clause_reviews") or [])
                   if isinstance(r, dict) and not r.get("supported")]
        print(f"\n--- flagged clauses: {len(flagged)} ---")
        for review in flagged:
            print(f"  [{review.get('category')}] {review.get('code') or ''} {str(review.get('clause'))[:120]}")
            print(f"      {str(review.get('reason'))[:260]}")
        for key in ("feedback_retries", "feedback_arbiter", "feedback_loop"):
            if record.get(key):
                print(f"\n--- {key} ---\n{str(record.get(key))[:900]}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
