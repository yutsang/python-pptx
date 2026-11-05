#!/usr/bin/env python3
"""Grade fdd_utils/ai/repair.py by BREAKING known-good text and asking it to restore — free, offline.

Two modes, both zero-token:

  --induce (default)  take the archived CLEAN finals of a run, corrupt exactly
                      one token per account (an amount x100 and x1000, a date, a
                      direction word), and ask the repair module to put it back.
  --census            replay the verifier over the archived finals AS THEY ARE,
                      tally the typed defects, and report how many the
                      DETERMINISTIC repair path fixes with the LLM disabled.

WHY RESTORATION, NOT "RE-VERIFIES CLEAN". A coincidental scale match re-verifies
clean by construction — that is exactly the failure mode the whole guard stack
exists for — so "the account came back green" proves nothing. The only
mechanical test that catches it is comparing the patched token against the value
that was there before the corruption. This tool therefore corrupts by rendering
the token through repair.render_amount_in_shape, the SAME helper the patch
renders with, so "restored" is well defined by construction rather than by a
second, independently written formatter. Pass condition, per the plan:

    the restored token parses to the ORIGINAL value AND carries the same
    currency prefix / unit suffix shape

Byte inequality is reported separately, as a formatting-drift signal, not as a
failure — nothing in the pipeline enforces a canonical number format.

DETECTION IS ASSERTED BEFORE ANYTHING IS COUNTED. A corruption the verifier never
flags is not a repair failure, and folding it into the denominator silently
flatters (or damns) the patcher for the pool's behaviour. Every corruption is
therefore classified first: escaped / flagged-but-not-typed / ambiguous /
repairable, and only the last is scored.

Dates are DETECTION-ONLY here. DATE_UNSUPPORTED routes to fact_patch_llm (a
model has to choose between the account's real period ends), so with the LLM
disabled there is nothing to restore; and `expected` is only the NEAREST source
date, which is a candidate, not the original. Direction likewise cannot arrive
through clause_reviews today — _direction_reviews is report-only — so this tool
feeds the finding in review shape to exercise the mechanism, and says so.

The databook is a REQUIRED positional argument for the same reason it is in
replay_verification.py: a run folder carries no workbook identity, and pairing
the wrong one does not fail, it silently grounds one project's commentary
against another project's numbers.

Usage (PYTHONPATH=. from the repo root):
    python ad-hoc/workbench/induce_repair_defects.py <databook.xlsx> --run <id> [--run <id> ...]
    python ad-hoc/workbench/induce_repair_defects.py <databook.xlsx> --runs-file local_runs.txt --census
"""

from __future__ import annotations

import argparse
import json
import sys
from collections import Counter
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))
if str(Path(__file__).resolve().parent) not in sys.path:
    sys.path.insert(0, str(Path(__file__).resolve().parent))

from fdd_utils.ai import repair as R  # noqa: E402
from fdd_utils.ai.validator import (  # noqa: E402
    SourceIndex,
    collect_direction_findings,
    extract_amount_spans,
    segment_clauses,
    verify_commentary,
)
from replay_verification import (  # noqa: E402
    account_text,
    build_dfs,
    check_pairing,
    load_results,
    resolve_run_dir,
    sibling_dfs_for_account,
)
from fdd_utils.ai import get_prompt_engine  # noqa: E402

_SCALE_FACTORS = (100.0, 1000.0)
_DIRECTION_FLIPS = [
    ("增加", "减少"), ("增长", "下降"), ("上升", "下降"),
    ("减少", "增加"), ("下降", "上升"),
    ("increased", "decreased"), ("decreased", "increased"),
    ("rose", "fell"), ("fell", "rose"), ("declined", "increased"),
]


def defective(reviews) -> List[Dict[str, Any]]:
    return [
        r for r in reviews
        if isinstance(r, dict) and not r.get("supported")
        and str(r.get("category") or "").lower() == "hallucination"
    ]


def find_target_entry(reviews, code: str, value: float) -> Optional[Dict[str, Any]]:
    """The review whose flagged token IS the corruption we introduced."""
    for review in defective(reviews):
        if str(review.get("code") or "") != code:
            continue
        for entry in review.get("amounts") or []:
            if entry.get("matched"):
                continue
            got = entry.get("value")
            if isinstance(got, (int, float)) and abs(float(got) - value) <= 1e-6 * max(abs(value), 1.0):
                return review
    return None


# --------------------------------------------------------------------------
# amount corruptions
# --------------------------------------------------------------------------

def induce_amount(base: str, df, sibs, source: SourceIndex, factor: float) -> Optional[Dict[str, Any]]:
    """Corrupt the first GROUNDED amount in `base` by `factor` and try to restore it."""
    for value, start, end in extract_amount_spans(base):
        if source.matches(value) is None:
            continue  # only corrupt a figure the verifier currently accepts
        span_text = base[start:end]
        broken_token = R.render_amount_in_shape(span_text, value, value * factor)
        if broken_token is None:
            continue
        broken = base[:start] + broken_token + base[end:]
        reviews = verify_commentary(broken, df, None, sibling_dfs=sibs)
        review = find_target_entry(reviews, "AMOUNT_SCALE_ERROR", value * factor)
        outcome: Dict[str, Any] = {
            "factor": factor, "original_token": span_text, "broken_token": broken_token,
            "original_value": value,
        }
        if review is None:
            # Either the pool still accepts the corrupted figure, or it was
            # flagged as plain AMOUNT_UNSUPPORTED (no unique scale factor).
            unsup = find_target_entry(reviews, "AMOUNT_UNSUPPORTED", value * factor)
            outcome["status"] = "not_typed_as_scale" if unsup else "escaped"
            return outcome
        entry = next(
            e for e in review["amounts"]
            if not e.get("matched") and isinstance(e.get("value"), (int, float))
            and abs(float(e["value"]) - value * factor) <= 1e-6 * max(abs(value * factor), 1.0)
        )
        if entry.get("ambiguous"):
            outcome["status"] = "ambiguous"
            return outcome
        outcome["status"] = "repairable"
        result = R.repair_content(
            content=broken, clause_reviews=reviews, df=df, sibling_dfs=sibs,
            language="Chi" if any("一" <= ch <= "鿿" for ch in base) else "Eng",
            allow_llm=False,
        )
        verified = [item for item in result["log"] if item.get("verified")]
        outcome["patched"] = bool(result["changed"] and verified)
        outcome["refusals"] = [item.get("reason") for item in result["log"] if not item.get("verified")]
        if not outcome["patched"]:
            return outcome
        restored = verified[0]["after"]
        parsed = R._parse_one_amount(restored)
        shape_old, shape_new = R._shape(span_text), R._shape(restored)
        outcome["restored_token"] = restored
        outcome["restores_value"] = parsed is not None and R._close(parsed, value, rel=0.005)
        outcome["same_shape"] = bool(
            shape_old and shape_new and shape_old[0] == shape_new[0] and shape_old[2] == shape_new[2]
        )
        outcome["byte_identical"] = restored == span_text
        return outcome
    return None


# --------------------------------------------------------------------------
# date corruptions (detection only — see the module docstring)
# --------------------------------------------------------------------------

def induce_date(base: str, df, sibs) -> Optional[Dict[str, Any]]:
    dates = R._dates_in_text(base)
    if not dates:
        return None
    year, month, day = sorted(dates)[0]
    for pattern, before, after in (
        ("%04d年%02d月%02d日" % (year, month, day), None, None),
        ("%04d年%d月%d日" % (year, month, day), None, None),
        ("%04d-%02d-%02d" % (year, month, day), None, None),
    ):
        if pattern in base:
            broken_token = pattern.replace(str(year), str(year - 7), 1)
            broken = base.replace(pattern, broken_token, 1)
            reviews = verify_commentary(broken, df, None, sibling_dfs=sibs)
            hit = [
                r for r in defective(reviews)
                if str(r.get("code")) == "DATE_UNSUPPORTED"
                and any(str(e.get("value", "")).startswith("%04d-" % (year - 7)) for e in r.get("amounts") or [])
            ]
            return {"original": pattern, "broken": broken_token, "detected": bool(hit)}
    return None


# --------------------------------------------------------------------------
# direction corruptions
# --------------------------------------------------------------------------

def induce_direction(base: str, df, sibs) -> Optional[Dict[str, Any]]:
    """Flip one direction word and ask the deterministic patcher to flip it back.

    Fed in review shape deliberately: _direction_reviews is report-only, so this
    can never arrive through clause_reviews until M3 is promoted. What it grades
    is the patcher, not the routing.
    """
    for start, end, clause in segment_clauses(base):
        for wrong, right in _DIRECTION_FLIPS:
            if wrong not in clause:
                continue
            broken = base[:start] + clause.replace(wrong, right) + base[end:]
            findings = collect_direction_findings(broken, df)
            if not findings:
                continue
            finding = findings[0]
            review = dict(finding, supported=False, category="hallucination")
            result = R.repair_content(
                content=broken, clause_reviews=[review], df=df, sibling_dfs=sibs,
                language="Chi" if any("一" <= ch <= "鿿" for ch in base) else "Eng",
                allow_llm=False,
            )
            verified = [item for item in result["log"] if item.get("verified")]
            return {
                "flip": "%s->%s" % (wrong, right),
                "detected": True,
                "patched": bool(verified),
                "restored": result["content"] == base,
                "refusals": [item.get("reason") for item in result["log"] if not item.get("verified")],
            }
    return None


# --------------------------------------------------------------------------
# modes
# --------------------------------------------------------------------------

def run_induce(runs: List[Path], dfs, prompt_manager) -> Dict[str, Any]:
    tally: Counter = Counter()
    drift: List[Tuple[str, str]] = []
    failures: List[Dict[str, Any]] = []
    for run_dir in runs:
        results = load_results(run_dir)
        for key, result in results.items():
            df = dfs.get(key)
            if df is None or not isinstance(result, dict):
                continue
            base, _field = account_text(result)
            if not str(base or "").strip():
                continue
            sibs = sibling_dfs_for_account(key, dfs, prompt_manager)
            reviews = verify_commentary(base, df, None, sibling_dfs=sibs)
            if defective(reviews):
                tally["accounts skipped (final is not clean)"] += 1
                continue
            tally["accounts used"] += 1
            source = SourceIndex.from_df(df, sibling_dfs=sibs)

            for factor in _SCALE_FACTORS:
                outcome = induce_amount(base, df, sibs, source, factor)
                if outcome is None:
                    tally["amount: no groundable amount to corrupt"] += 1
                    continue
                label = "amount x%d" % factor
                tally["%s: corruptions" % label] += 1
                status = outcome["status"]
                if status != "repairable":
                    tally["%s: %s" % (label, status)] += 1
                    continue
                tally["%s: detected + uniquely typed" % label] += 1
                if not outcome["patched"]:
                    tally["%s: PATCH REFUSED" % label] += 1
                    failures.append({"key": key, "label": label, **outcome})
                    continue
                tally["%s: patched" % label] += 1
                if outcome["restores_value"] and outcome["same_shape"]:
                    tally["%s: RESTORED (value + shape)" % label] += 1
                else:
                    tally["%s: PATCHED BUT NOT RESTORED" % label] += 1
                    failures.append({"key": key, "label": label, **outcome})
                if not outcome["byte_identical"]:
                    drift.append((outcome["original_token"], outcome.get("restored_token", "")))

            date_outcome = induce_date(base, df, sibs)
            if date_outcome is None:
                tally["date: no date to corrupt"] += 1
            else:
                tally["date: corruptions"] += 1
                tally["date: detected" if date_outcome["detected"] else "date: escaped"] += 1

            dir_outcome = induce_direction(base, df, sibs)
            if dir_outcome is None:
                tally["direction: no flippable clause"] += 1
            else:
                tally["direction: corruptions"] += 1
                tally["direction: detected"] += 1
                if dir_outcome["patched"] and dir_outcome["restored"]:
                    tally["direction: RESTORED"] += 1
                elif dir_outcome["patched"]:
                    tally["direction: PATCHED BUT NOT RESTORED"] += 1
                    failures.append({"key": key, "label": "direction", **dir_outcome})
                else:
                    tally["direction: PATCH REFUSED"] += 1
                    failures.append({"key": key, "label": "direction", **dir_outcome})
    return {"tally": tally, "drift": drift, "failures": failures}


def run_census(runs: List[Path], dfs, prompt_manager) -> Dict[str, Any]:
    """The archived finals as they are: what is defective, and what a
    deterministic-only repair pass would fix."""
    tally: Counter = Counter()
    refusals: Counter = Counter()
    for run_dir in runs:
        results = load_results(run_dir)
        tally["runs"] += 1
        for key, result in results.items():
            df = dfs.get(key)
            if df is None or not isinstance(result, dict):
                continue
            base, _field = account_text(result)
            if not str(base or "").strip():
                continue
            sibs = sibling_dfs_for_account(key, dfs, prompt_manager)
            reviews = verify_commentary(base, df, None, sibling_dfs=sibs)
            bad = defective(reviews)
            tally["accounts"] += 1
            for review in bad:
                tally["defect: %s" % (review.get("code") or "NO_CODE")] += 1
            repairable = R.repairable_reviews(bad)
            if not repairable:
                continue
            tally["accounts with a repairable defect"] += 1
            for review in repairable:
                tally["repairable: %s" % review.get("code")] += 1
            outcome = R.repair_content(
                content=base, clause_reviews=reviews, df=df, sibling_dfs=sibs,
                language="Chi" if any("一" <= ch <= "鿿" for ch in base) else "Eng",
                allow_llm=False,
            )
            for item in outcome["log"]:
                if item.get("verified"):
                    tally["FIXED (deterministic): %s" % item.get("code")] += 1
                else:
                    refusals["%s: %s" % (item.get("code"), item.get("reason"))] += 1
    return {"tally": tally, "refusals": refusals}


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("databook", help="the databook this run was produced from")
    parser.add_argument("--run", action="append", default=[], help="run folder / timestamp (repeatable)")
    parser.add_argument("--runs-file", help="file with one run id per line (keep it local — run ids pair with client files)")
    parser.add_argument("--entity", default="")
    parser.add_argument("--sheet", default=None)
    parser.add_argument("--census", action="store_true", help="tally real defects and deterministic fixes instead of inducing")
    parser.add_argument("--json", help="write the tally to this path")
    args = parser.parse_args()

    run_args = list(args.run)
    if args.runs_file:
        run_args += [line.strip() for line in Path(args.runs_file).read_text().splitlines() if line.strip()]
    if not run_args:
        sys.exit("❌ pass at least one --run or a --runs-file")

    dfs = build_dfs(args.databook, args.entity, args.sheet)
    prompt_manager = get_prompt_engine()
    runs: List[Path] = []
    for run_arg in run_args:
        run_dir = resolve_run_dir(run_arg)
        results = load_results(run_dir)
        try:
            check_pairing(results, dfs, args.databook, run_dir)
        except SystemExit:
            print("   skipped %s (does not pair with this databook)" % run_dir.name)
            continue
        runs.append(run_dir)
    if not runs:
        sys.exit("❌ no run paired with this databook")

    out = run_census(runs, dfs, prompt_manager) if args.census else run_induce(runs, dfs, prompt_manager)
    print("\n%s over %d run(s), databook %s" % (
        "CENSUS" if args.census else "INDUCED DEFECTS", len(runs), Path(args.databook).name))
    for label, count in sorted(out["tally"].items()):
        print("   %-52s %6d" % (label, count))
    if args.census:
        print("\n   refusal reasons (deterministic path, LLM disabled):")
        for reason, count in out["refusals"].most_common(15):
            print("      %-70s %4d" % (reason[:70], count))
    else:
        drift = out["drift"]
        print("\n   formatting drift (restored != original bytes): %d" % len(drift))
        for original, restored in drift[:5]:
            print("      %r -> %r" % (original, restored))
        if out["failures"]:
            print("\n   FAILURES (%d):" % len(out["failures"]))
            for failure in out["failures"][:10]:
                print("      %s" % json.dumps(failure, ensure_ascii=False, default=str)[:220])
    if args.json:
        Path(args.json).write_text(json.dumps(
            {"tally": dict(out["tally"]), "refusals": dict(out.get("refusals") or {})},
            ensure_ascii=False, indent=2), encoding="utf-8")
        print("\n   wrote %s" % args.json)


if __name__ == "__main__":
    main()
