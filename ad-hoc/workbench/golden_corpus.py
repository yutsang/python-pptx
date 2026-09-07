#!/usr/bin/env python3
"""Golden-corpus regression gate for the deterministic verifier (M8.3 scaffolding).

The repo has no tests, no fixtures and no CI. Every gate in the plan today is an
import check, a self-referential replay, or an eyeball of a paid run — nothing
regression-tests a verifier change against expected output. This is the missing
piece: a small set of (databook, archived run) cases whose verdicts are recorded
once and re-scored for free thereafter.

WHAT IS COMMITTED, AND WHY IT IS HASHED
Every real databook at the repo root is gitignored client data, and so is the
commentary text inside a run's results.yml. A golden file holding clause text
could not be committed, which would defeat the whole point. So the golden file
stores, per clause, only:

    sha1(account + "\\x1f" + normalized clause)[:16]  ->  "<category>/<supported>"

That fingerprint is stable across machines, carries no recoverable client text
(the clauses are free prose, not a guessable key space), and still detects every
verdict change: a flipped verdict changes the value, an edited or re-segmented
clause changes the key. What it cannot show is WHICH clause moved — for that,
run replay_verification.py --diff on two snapshots, which keeps the text locally.

THE CASES ARE NOT COMMITTABLE EITHER
A case names a databook path and a run id, both machine-local. So the manifest
lives next to the golden file and is written by --record on whoever's machine has
the files. Re-recording on a different machine with different databooks produces
a different corpus; the golden file records which manifest it came from and
--check refuses to score a mismatched pair.

STILL OPEN (M8.3's own text, not doable from this file's scope): the tracked,
client-free pair — make_demo_databook.py and fdd_utils/demo_results.json — is
the corpus that COULD be committed whole. It is unusable today: the demo loader
reads fdd_utils/ui/demo_results.json while the tracked file is
fdd_utils/demo_results.json (ui/ai_panel.py:197, failure swallowed), and the
demo results' account keys ("Other CA", "IP", "Advances") match no generated
databook's dfs keys. Fixing that path and aligning those keys turns this harness
into a corpus that needs no local client file at all.

Usage:
    # once, on a machine that has the databooks:
    python ad-hoc/workbench/golden_corpus.py --record \\
        --case "<databook>.xlsx::20260715_010741" \\
        --case "<databook>.xlsx::20260718_095717"

    # thereafter, free, before and after any verifier change:
    python ad-hoc/workbench/golden_corpus.py --check
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import sys
import time
from collections import Counter
from pathlib import Path
from typing import Any, Dict, List, Tuple

sys.path.insert(0, str(Path(__file__).resolve().parent))
from replay_verification import (  # noqa: E402
    _clause_key, build_dfs, check_pairing, load_results, replay_run, resolve_run_dir,
)

HERE = Path(__file__).resolve().parent
GOLDEN_PATH = HERE / "golden_verdicts.json"


def fingerprint(account: str, clause: str) -> str:
    raw = f"{account}\x1f{_clause_key(clause)}".encode("utf-8")
    return hashlib.sha1(raw).hexdigest()[:16]


def parse_case(spec: str) -> Tuple[str, str]:
    if "::" not in spec:
        sys.exit(f"❌ --case must be '<databook path>::<run id>', got {spec!r}")
    databook, run = spec.split("::", 1)
    return databook.strip(), run.strip()


def score_case(databook: str, run: str, entity: str) -> Dict[str, str]:
    run_dir = resolve_run_dir(run)
    results = load_results(run_dir)
    dfs = build_dfs(databook, entity, None)
    check_pairing(results, dfs, databook, run_dir)
    records = replay_run(results, dfs)
    verdicts: Dict[str, str] = {}
    for r in records:
        verdicts[fingerprint(r["account"], r["clause"])] = f"{r['category']}/{int(r['supported'])}"
    return verdicts


def run_cases(cases: List[str], entity: str) -> Dict[str, Any]:
    corpus: Dict[str, Any] = {"cases": {}, "recorded": time.strftime("%Y-%m-%d %H:%M:%S")}
    for spec in cases:
        databook, run = parse_case(spec)
        if not os.path.exists(databook):
            sys.exit(f"❌ No such databook: {databook}")
        print(f"scoring {os.path.basename(databook)} × {run} ...")
        verdicts = score_case(databook, run, entity)
        corpus["cases"][spec] = {
            "run": run,
            "databook_basename": os.path.basename(databook),
            "verdicts": verdicts,
        }
        flagged = sum(1 for v in verdicts.values() if v.endswith("/0"))
        print(f"  {len(verdicts)} clauses, {flagged} unsupported")
    return corpus


def cmd_record(args: argparse.Namespace) -> None:
    if not args.case:
        sys.exit("❌ --record needs at least one --case '<databook>::<run id>'")
    corpus = run_cases(args.case, args.entity)
    GOLDEN_PATH.write_text(json.dumps(corpus, indent=1, sort_keys=True), encoding="utf-8")
    total = sum(len(c["verdicts"]) for c in corpus["cases"].values())
    print(f"\n💾 recorded {len(corpus['cases'])} case(s), {total} clause fingerprints -> {GOLDEN_PATH}")
    print("   Commit this file. It holds no clause text — see the module docstring.")


def cmd_check(args: argparse.Namespace) -> None:
    if not GOLDEN_PATH.exists():
        sys.exit(f"❌ No golden file at {GOLDEN_PATH}. Run --record first.")
    golden = json.loads(GOLDEN_PATH.read_text(encoding="utf-8"))
    cases = args.case or sorted(golden["cases"])
    unknown = [c for c in cases if c not in golden["cases"]]
    if unknown:
        sys.exit(f"❌ Not in the golden file: {unknown}\n   Known cases: {sorted(golden['cases'])}")

    total_drift = 0
    for spec in cases:
        databook, run = parse_case(spec)
        expected = golden["cases"][spec]["verdicts"]
        if not os.path.exists(databook):
            print(f"⏭  SKIP {spec} — databook not on this machine")
            continue
        actual = score_case(databook, run, args.entity)
        changed = {k: (expected[k], actual[k]) for k in expected.keys() & actual.keys()
                   if expected[k] != actual[k]}
        gone = sorted(expected.keys() - actual.keys())
        new = sorted(actual.keys() - expected.keys())
        drift = len(changed) + len(gone) + len(new)
        total_drift += drift
        status = "PASS" if drift == 0 else "DRIFT"
        print(f"\n[{status}] {spec}")
        print(f"  clauses: {len(expected)} golden, {len(actual)} now")
        if changed:
            print(f"  {len(changed)} verdict change(s):")
            for k, (was, now) in sorted(changed.items()):
                print(f"    {k}  {was} -> {now}")
            moves = Counter(f"{was} -> {now}" for was, now in changed.values())
            for move, count in moves.most_common():
                print(f"      {count:>4}  {move}")
        if gone:
            print(f"  {len(gone)} clause(s) in golden but not produced now (re-segmented or dropped)")
        if new:
            print(f"  {len(new)} clause(s) produced now but not in golden")

    print(f"\n=== {'PASS — verdict-identical' if total_drift == 0 else f'DRIFT — {total_drift} change(s)'} ===")
    if total_drift:
        print("A change here is not automatically a regression: M0 and M2a both intend to move")
        print("verdicts. Read the direction with replay_verification.py --diff, decide, then")
        print("re-record. Silent drift is the thing this gate exists to prevent.")
    sys.exit(1 if total_drift else 0)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--record", action="store_true", help="score the cases and write the golden file")
    ap.add_argument("--check", action="store_true", help="re-score and report drift (exit 1 on drift)")
    ap.add_argument("--case", action="append",
                    help="'<databook path>::<run id>'; repeatable. With --check, defaults to every recorded case.")
    ap.add_argument("--entity", default="", help="entity name filter passed to process_workbook_data")
    args = ap.parse_args()

    if args.record == args.check:
        ap.error("pass exactly one of --record / --check")
    (cmd_record if args.record else cmd_check)(args)


if __name__ == "__main__":
    main()
