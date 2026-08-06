#!/usr/bin/env python3
"""Measure what each AI subagent stage actually CHANGES, per account.

Motivating question: the active pipeline is Generator -> Auditor -> Validator
(3 LLM calls per account). If most of the time a stage returns its input
essentially unchanged, that stage is being paid for without earning it, and
the pipeline can be restructured. This script answers that from a real run's
own log instead of by eyeballing a few accounts.

For every account in a run's results.yml it reports:
  * whether each stage was a NO-OP (output identical to its input)
  * how much it changed (character-level similarity + net length delta)
  * WHAT it changed, classified into the categories the prompts actually
    care about: dropped an advisory recommendation, changed a number,
    stripped bullet markers, etc.
  * how many clauses the Validator flagged unsupported -- i.e. whether the
    existing (currently disabled) feedback loop would ever have fired

Usage:
    # newest run under fdd_utils/logs/
    python inspect_stage_value.py

    # a specific run
    python inspect_stage_value.py fdd_utils/logs/run_20260806_101500/results.yml

    # show the actual diff text for stages that DID change something
    python inspect_stage_value.py --show-diffs

Read-only. Touches nothing but the log file.
"""
from __future__ import annotations

import argparse
import difflib
import glob
import os
import re
import sys
from typing import Any, Dict, List, Optional, Tuple

import yaml

# The active sequence (fdd_utils/ai.py SUBAGENT_SEQUENCE). subagent_3/Refiner
# is dormant and is skipped here if absent rather than reported as missing.
STAGES = [("subagent_1", "Generator"), ("subagent_2", "Auditor"), ("subagent_4", "Validator")]

BULLET_MARKERS = ("➢", "■", "•", "- ")
# Phrases the Auditor prompt explicitly exists to strip (point 9: "REMOVE
# generic consultant advisory"). Used to tell a substantive edit from a
# cosmetic one.
ADVISORY_RE = re.compile(r"建议|應注意|应关注|建議|You should|we recommend|recommend that", re.I)
NUM_RE = re.compile(r"-?[\d,]+\.?\d*")


def _find_latest_results() -> Optional[str]:
    candidates = sorted(glob.glob("fdd_utils/logs/run_*/results.yml"), key=os.path.getmtime)
    return candidates[-1] if candidates else None


def _norm(text: Any) -> str:
    """Whitespace-normalised text, so a pure re-wrap doesn't read as a change."""
    if not isinstance(text, str):
        return ""
    return re.sub(r"\s+", "", text)


def _similarity(a: str, b: str) -> float:
    if not a and not b:
        return 1.0
    return difflib.SequenceMatcher(None, a, b).ratio()


def _numbers(text: str) -> List[str]:
    return NUM_RE.findall(text or "")


def _classify_change(before: str, after: str) -> List[str]:
    """What KIND of edit was this? Multiple labels can apply."""
    tags: List[str] = []
    if _numbers(before) != _numbers(after):
        tags.append("NUMBERS-CHANGED")
    b_bul = sum(before.count(m) for m in BULLET_MARKERS)
    a_bul = sum(after.count(m) for m in BULLET_MARKERS)
    if b_bul > a_bul:
        tags.append("BULLETS-STRIPPED")
    elif a_bul > b_bul:
        tags.append("BULLETS-ADDED")
    if ADVISORY_RE.search(before) and not ADVISORY_RE.search(after):
        tags.append("ADVISORY-REMOVED")
    delta = len(_norm(after)) - len(_norm(before))
    if delta < -20:
        tags.append(f"SHORTENED({delta})")
    elif delta > 20:
        tags.append(f"LENGTHENED(+{delta})")
    if not tags:
        tags.append("cosmetic-only")
    return tags


def _validator_flags(entry: Dict[str, Any]) -> Tuple[int, int]:
    """(unsupported, total) from the STORED clause_reviews.

    Note these are the AUTHORITATIVE reviews -- fdd_utils/ai.py's
    verify_commentary layers a deterministic number check over the LLM's own
    judgment, so this is not purely what the model said.
    """
    validation = entry.get("agent_4_validation") or {}
    reviews = validation.get("clause_reviews") or []
    if not isinstance(reviews, list):
        return 0, 0
    total = len(reviews)
    unsupported = sum(1 for r in reviews if isinstance(r, dict) and not r.get("supported", True))
    return unsupported, total


def _llm_own_flags(entry: Dict[str, Any]) -> Tuple[int, int]:
    """(unsupported, total) from the model's OWN raw_response, before the
    deterministic layer overwrote it. The gap between this and
    _validator_flags is how much of the Validator's output is discarded."""
    validation = entry.get("agent_4_validation") or {}
    raw = validation.get("raw_response")
    if not isinstance(raw, str):
        return 0, 0
    try:
        import json
        parsed = json.loads(raw)
    except Exception:
        return 0, 0
    reviews = parsed.get("clause_reviews") or []
    if not isinstance(reviews, list):
        return 0, 0
    total = len(reviews)
    unsupported = sum(1 for r in reviews if isinstance(r, dict) and not r.get("supported", True))
    return unsupported, total


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("results_yml", nargs="?", help="path to a run's results.yml "
                                                  "(default: newest under fdd_utils/logs/)")
    ap.add_argument("--show-diffs", action="store_true",
                    help="print the actual before/after text for every stage that changed something")
    ap.add_argument("--noop-threshold", type=float, default=0.995,
                    help="similarity at or above which a stage counts as a no-op (default 0.995)")
    args = ap.parse_args()

    path = args.results_yml or _find_latest_results()
    if not path or not os.path.exists(path):
        print("❌ No results.yml found. Pass one explicitly, or run the pipeline first.")
        print("   Expected location: fdd_utils/logs/run_*/results.yml")
        return 1

    print(f"Analysing: {path}\n")
    with open(path, encoding="utf-8") as f:
        results = yaml.safe_load(f) or {}

    accounts = [k for k, v in results.items() if isinstance(v, dict)]
    if not accounts:
        print("❌ results.yml has no account entries.")
        return 1

    # ---- per-stage transition analysis -------------------------------------
    transitions = [(STAGES[i][0], STAGES[i][1], STAGES[i + 1][0], STAGES[i + 1][1])
                   for i in range(len(STAGES) - 1)]

    print("=" * 78)
    print("  PER-ACCOUNT: did each stage actually change its input?")
    print("=" * 78)
    stats: Dict[str, Dict[str, Any]] = {
        label_to: {"noop": 0, "changed": 0, "sims": [], "tags": {}}
        for _, _, _, label_to in transitions
    }

    for acct in sorted(accounts):
        entry = results[acct]
        print(f"\n{acct}")
        for key_from, label_from, key_to, label_to in transitions:
            before, after = entry.get(key_from), entry.get(key_to)
            if not isinstance(before, str) or not isinstance(after, str):
                print(f"  {label_to:<10} (missing output — stage skipped or errored)")
                continue
            nb, na = _norm(before), _norm(after)
            sim = _similarity(nb, na)
            stats[label_to]["sims"].append(sim)
            if nb == na or sim >= args.noop_threshold:
                stats[label_to]["noop"] += 1
                print(f"  {label_to:<10} NO-OP        (similarity {sim:6.1%}) — this call bought nothing")
            else:
                stats[label_to]["changed"] += 1
                tags = _classify_change(before, after)
                for t in tags:
                    base = t.split("(")[0]
                    stats[label_to]["tags"][base] = stats[label_to]["tags"].get(base, 0) + 1
                print(f"  {label_to:<10} changed      (similarity {sim:6.1%})  {', '.join(tags)}")
                if args.show_diffs:
                    for line in difflib.unified_diff(
                        [before], [after], fromfile=label_from, tofile=label_to, lineterm=""
                    ):
                        print(f"      {line}")

    # ---- summary -----------------------------------------------------------
    print()
    print("=" * 78)
    print("  SUMMARY — how often is each stage earning its runtime?")
    print("=" * 78)
    for _, _, _, label_to in transitions:
        s = stats[label_to]
        n = s["noop"] + s["changed"]
        if not n:
            continue
        avg = sum(s["sims"]) / len(s["sims"]) if s["sims"] else 0.0
        print(f"\n  {label_to}:")
        print(f"    no-op   : {s['noop']:>3}/{n}  ({s['noop']/n:.0%})  <- pure cost, zero effect")
        print(f"    changed : {s['changed']:>3}/{n}  ({s['changed']/n:.0%})")
        print(f"    mean similarity to its input: {avg:.1%}")
        if s["tags"]:
            print("    what it changed, by kind:")
            for tag, count in sorted(s["tags"].items(), key=lambda kv: -kv[1]):
                note = ""
                if tag == "BULLETS-STRIPPED":
                    note = "  ⚠️  destroys detail-table lead-in structure"
                elif tag == "cosmetic-only":
                    note = "  (no numbers, no advisory, no structure — wording only)"
                print(f"      {tag:<20} {count:>3}{note}")

    # ---- would the existing feedback loop ever fire? -----------------------
    print()
    print("=" * 78)
    print("  VALIDATOR FLAGS — would the (currently disabled) feedback loop fire?")
    print("=" * 78)
    print("  fdd_utils/ai.py already implements the retry loop: _run_feedback_loop_for_key")
    print("  re-runs Generator->Auditor->Validator when unsupported_ratio > threshold.")
    print("  Config: processing.feedback_loop.{enabled,max_retries,unsupported_threshold}")
    print("  Default is enabled: False, max_retries: 2, unsupported_threshold: 0.30\n")

    would_fire = 0
    total_final, total_unsup = 0, 0
    llm_total, llm_unsup = 0, 0
    for acct in sorted(accounts):
        unsup, total = _validator_flags(results[acct])
        l_unsup, l_total = _llm_own_flags(results[acct])
        total_final += total
        total_unsup += unsup
        llm_total += l_total
        llm_unsup += l_unsup
        ratio = (unsup / total) if total else 0.0
        fires = ratio > 0.30
        if fires:
            would_fire += 1
            print(f"  ⚠️  {acct}: {unsup}/{total} unsupported (ratio {ratio:.0%}) -> WOULD retry")

    if not would_fire:
        print("  ✅ No account exceeds the 0.30 threshold — the feedback loop would")
        print("     never fire on this run. Enabling it costs 0 extra calls here,")
        print("     but also delivers 0 benefit until the Validator actually flags")
        print("     something.")
    print(f"\n  Authoritative clause_reviews : {total_unsup}/{total_final} unsupported")
    print(f"  Validator LLM's OWN raw json : {llm_unsup}/{llm_total} unsupported")
    if llm_total and total_final and llm_total != total_final:
        print(f"  -> the stored reviews have {total_final} clauses vs the model's own {llm_total}:")
        print("     the deterministic grounding layer (verify_commentary, ai.py:1287)")
        print("     RE-SEGMENTS and RE-JUDGES the clauses, so much of what the")
        print("     Validator is asked to produce is overwritten before it is used.")
        print("     That output is the main reason this stage is the slowest one.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
