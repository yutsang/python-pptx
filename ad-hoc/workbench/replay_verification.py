#!/usr/bin/env python3
"""Replay the deterministic verifier over an archived run — free, offline, repeatable.

This is the instrument every verifier-side change is graded with. A run under
fdd_utils/logs/run_<ts>/ stores the text the pipeline shipped; this tool rebuilds
the SAME source data from a databook, re-runs verify_commentary over that stored
text, and prints what the verifier says today. Nothing here calls an LLM.

Why it needs the databook as a REQUIRED argument: an archived run folder carries
no workbook identity at all (no path, no entity, no sheet hash). Pairing the
wrong databook does not fail — it silently grounds one project's commentary
against another project's numbers and reports a diff that means nothing. So the
path is positional, and the replay refuses to run when the results keys are not
a subset of the rebuilt dfs keys.

Modes
  (default)   replay one run and print the defect table
  --snapshot  also write the verdicts to JSON, for a before/after pair
  --diff A B  compare two snapshots; no databook needed
  --decoys    discrimination test: deform real cell values and report the share
              the grounding pool still accepts

Typical grading loop for a pool change:
    python ad-hoc/workbench/replay_verification.py <book.xlsx> --run <id> \
        --snapshot /tmp/before.json --decoys
    ... make the change ...
    python ad-hoc/workbench/replay_verification.py <book.xlsx> --run <id> \
        --snapshot /tmp/after.json --decoys
    python ad-hoc/workbench/replay_verification.py --diff /tmp/before.json /tmp/after.json

Run with PYTHONPATH=. from the repo root.
"""

from __future__ import annotations

import argparse
import json
import os
import random
import re
import sys
import time
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import yaml

REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from fdd_utils.ai import get_prompt_engine, verify_commentary  # noqa: E402
# _to_float is private, imported deliberately: the decoy test must read cells with
# exactly the coercion SourceIndex uses, or it would compare against a pool built
# from a different set of numbers than the one it sampled from.
from fdd_utils.ai.validator import INTERNAL_ROW_KEY, SourceIndex, _to_float  # noqa: E402
from fdd_utils.financial_common import get_pipeline_result_text  # noqa: E402
from fdd_utils.workbook import process_workbook_data  # noqa: E402

LOGS_DIR = REPO_ROOT / "fdd_utils" / "logs"


# --------------------------------------------------------------------------
# sibling_dfs — the one piece of production logic this tool must reproduce
# --------------------------------------------------------------------------

def sibling_dfs_for_account(mapping_key: str, dfs: Dict[str, Any], prompt_manager) -> Optional[List[Any]]:
    """Every OTHER account's df of the same statement type, or None.

    Reproduces pipeline.py:696-706 (the subagent_2/subagent_4 verify call) and
    its duplicate at :1336-1341 (_ensure_clause_reviews_on_final). Both build the
    same set the same way; the plan (M1 step 6) folds them into one function that
    later becomes RunState.siblings. It lives here rather than in pipeline.py
    because this tool is not allowed to edit pipeline.py yet — when that lands,
    delete this and import the shared one.

    This must stay byte-identical in behaviour to production: a different sibling
    set is a different grounding pool, so a replay built on a looser or tighter
    set would report verdict changes that the real pipeline never made, and every
    before/after diff downstream would be measuring this function instead of the
    change under test.

    Returns None (not []) when the account has no statement type, because that is
    what production passes and SourceIndex.from_df treats them the same only by
    accident of `sibling_dfs or []`.
    """
    statement_type = prompt_manager.get_mapping_component(mapping_key, component="type")
    if not statement_type:
        return None
    return [
        other_df for other_key, other_df in dfs.items()
        if other_key != mapping_key
        and prompt_manager.get_mapping_component(other_key, component="type") == statement_type
    ]


# --------------------------------------------------------------------------
# loading
# --------------------------------------------------------------------------

def resolve_run_dir(run_arg: str) -> Path:
    """Accept a full path, a folder name, or a bare timestamp."""
    for candidate in (Path(run_arg), LOGS_DIR / run_arg, LOGS_DIR / f"run_{run_arg}"):
        if candidate.is_dir():
            return candidate
    sys.exit(f"❌ No such run folder: {run_arg!r} (looked in {LOGS_DIR})")


def load_results(run_dir: Path) -> Dict[str, Any]:
    results_path = run_dir / "results.yml"
    if not results_path.exists():
        sys.exit(f"❌ {results_path} does not exist — this run archived no results (see list_archived_runs.py).")
    with open(results_path, encoding="utf-8") as fh:
        data = yaml.safe_load(fh) or {}
    return {k: v for k, v in data.items() if isinstance(v, dict)}


def build_dfs(databook: str, entity: str, sheet: Optional[str]) -> Dict[str, Any]:
    """Rebuild the dfs the pipeline was given, the way production builds them.

    process_workbook_data returns the detail_analysis variant as state["dfs"] —
    the same dict both the Streamlit path and inspect_databook.py hand to
    run_ai_pipeline_with_progress, so the df each account is grounded against is
    the same object shape the run itself used.
    """
    if sheet is None:
        import inspect_databook  # heavy, and only needed when the sheet is not given
        found = inspect_databook._resolve_financials_sheets(pd.ExcelFile(databook))
        sheet = found[0] if found else None
    state = process_workbook_data(
        temp_path=databook, entity_name=entity, selected_sheet=sheet, debug=False,
    )
    return state.get("dfs") or {}


def check_pairing(results: Dict[str, Any], dfs: Dict[str, Any], databook: str, run_dir: Path) -> None:
    """Refuse to replay a run against a databook that cannot be its source.

    The archive records no workbook identity, so this key-subset test is the only
    evidence available that the pair is the right one. It is necessary, not
    sufficient: two workbooks from the same template share an account vocabulary.
    """
    missing = sorted(set(results) - set(dfs))
    if missing:
        sys.exit(
            f"❌ REFUSING TO REPLAY — wrong databook for this run.\n"
            f"   run:      {run_dir.name} ({len(results)} accounts)\n"
            f"   databook: {databook} ({len(dfs)} accounts)\n"
            f"   {len(missing)} account(s) in the results have no df here: {missing}\n"
            f"   Available: {sorted(dfs)}\n"
            f"   Grounding this run's text against these numbers would produce a\n"
            f"   silently mismatched diff. Pass the databook the run actually used."
        )


# --------------------------------------------------------------------------
# replay
# --------------------------------------------------------------------------

def account_text(result: Dict[str, Any]) -> Tuple[str, str]:
    """(text, which_field). Spans and clause offsets are defined against
    agent_4_validation["final_content"] — the stored string, which for some
    accounts differs from results["final"] — so prefer it, exactly as the
    plan's span-coordinate rule requires. Fall back to the usual priority chain
    for accounts the Validator skipped."""
    validation = result.get("agent_4_validation") or {}
    stored = validation.get("final_content") if isinstance(validation, dict) else None
    if isinstance(stored, str) and stored.strip():
        return stored, "agent_4_validation.final_content"
    return get_pipeline_result_text(result), "get_pipeline_result_text"


def replay_run(
    results: Dict[str, Any],
    dfs: Dict[str, Any],
    *,
    use_archived_llm_reviews: bool = True,
) -> List[Dict[str, Any]]:
    """Re-run verify_commentary over every account's final text.

    use_archived_llm_reviews: the archived clause_reviews are the COMBINED
    verdicts — the pipeline overwrites the LLM's own list with verify_commentary's
    output (pipeline.py:707), so the raw LLM reviews are not recoverable from the
    archive. Feeding the combined list back in as llm_clause_reviews is the
    closest available reconstruction: _combine_verdict preserves an LLM reasoning
    flag on a clause whose numbers check out, which is exactly the half that would
    otherwise vanish. The one thing it cannot restore is an LLM hallucination flag
    that the deterministic layer had already dropped. Pass False to see the purely
    deterministic verdicts.
    """
    prompt_manager = get_prompt_engine()
    records: List[Dict[str, Any]] = []
    for key in sorted(results):
        result = results[key]
        text, source_field = account_text(result)
        if not str(text or "").strip():
            continue
        df = dfs.get(key)
        if df is None:
            continue
        archived = ((result.get("agent_4_validation") or {}).get("clause_reviews")
                    if isinstance(result.get("agent_4_validation"), dict) else None)
        reviews = verify_commentary(
            text, df,
            archived if use_archived_llm_reviews else None,
            sibling_dfs=sibling_dfs_for_account(key, dfs, prompt_manager),
        )
        for i, review in enumerate(reviews):
            records.append({
                "account": key,
                "idx": i,
                "clause": review.get("clause", ""),
                "supported": bool(review.get("supported")),
                "category": review.get("category"),
                "reason": review.get("reason", ""),
                "text_source": source_field,
            })
    return records


def archived_agreement(results: Dict[str, Any], records: List[Dict[str, Any]]) -> Tuple[int, int]:
    """(agreeing, comparable) between the replay and what the run shipped.

    A replay on unchanged code should reproduce the archive almost exactly. When
    it does not, the gap is code drift since the run — which is worth knowing
    BEFORE reading a before/after diff, because drift and the change under test
    land in the same column.
    """
    replayed = {(r["account"], _clause_key(r["clause"])): r["supported"] for r in records}
    agree = comparable = 0
    for key, result in results.items():
        validation = result.get("agent_4_validation") or {}
        for review in (validation.get("clause_reviews") or []) if isinstance(validation, dict) else []:
            ck = (key, _clause_key(review.get("clause", "")))
            if ck in replayed:
                comparable += 1
                agree += int(replayed[ck] == bool(review.get("supported")))
    return agree, comparable


def _clause_key(clause: str) -> str:
    return re.sub(r"\s+", " ", str(clause or "")).strip()


# --------------------------------------------------------------------------
# reporting
# --------------------------------------------------------------------------

def print_defect_table(records: List[Dict[str, Any]]) -> None:
    unsupported = [r for r in records if not r["supported"]]
    by_account = defaultdict(list)
    for r in unsupported:
        by_account[r["account"]].append(r)

    print("\n--- DEFECT TABLE (account × category × reason) ---")
    if not unsupported:
        print("  (no unsupported clauses)")
    for account in sorted(by_account):
        rows = by_account[account]
        total = sum(1 for r in records if r["account"] == account)
        print(f"\n  {account}  ({len(rows)} unsupported of {total} reviews)")
        for r in rows:
            clause = _clause_key(r["clause"])
            print(f"    [{r['category']}] {clause[:88]}")
            print(f"        reason: {_clause_key(r['reason'])[:110]}")

    print("\n--- TOTALS ---")
    accounts = {r["account"] for r in records}
    print(f"  accounts replayed : {len(accounts)}")
    print(f"  clause reviews    : {len(records)}")
    print(f"  unsupported       : {len(unsupported)}"
          f"  ({100.0 * len(unsupported) / max(len(records), 1):.1f}%)")
    print("  by category       :")
    for category, count in Counter(r["category"] for r in records).most_common():
        flagged = sum(1 for r in unsupported if r["category"] == category)
        print(f"      {str(category):<16} {count:>5} total, {flagged:>4} unsupported")


# --------------------------------------------------------------------------
# snapshots and diff
# --------------------------------------------------------------------------

def write_snapshot(path: str, records: List[Dict[str, Any]], meta: Dict[str, Any]) -> None:
    payload = {"meta": meta, "verdicts": records}
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=1)
    print(f"\n💾 snapshot written: {path}  ({len(records)} verdicts)")


def diff_snapshots(before_path: str, after_path: str) -> None:
    def load(p):
        with open(p, encoding="utf-8") as fh:
            return json.load(fh)

    before, after = load(before_path), load(after_path)
    b_meta, a_meta = before.get("meta", {}), after.get("meta", {})
    print("--- SNAPSHOT DIFF ---")
    print(f"  before: {before_path}  run={b_meta.get('run')} databook={b_meta.get('databook')}")
    print(f"  after : {after_path}  run={a_meta.get('run')} databook={a_meta.get('databook')}")
    if b_meta.get("run") != a_meta.get("run") or b_meta.get("databook") != a_meta.get("databook"):
        print("  ⚠️  These snapshots are NOT the same run/databook pair. The diff below "
              "mixes the change under test with a different corpus.")

    # Keyed on (account, clause text): the clause is a verbatim substring of the
    # final content and the verifier is deterministic, so the same clause on the
    # same account is the same unit of judgement across the two runs. Index
    # position is NOT usable — a pool change can add or drop composition/date
    # reviews, which shifts every index after it.
    b_map = {(r["account"], _clause_key(r["clause"])): r for r in before["verdicts"]}
    a_map = {(r["account"], _clause_key(r["clause"])): r for r in after["verdicts"]}

    cleared, newly_flagged, recategorized = [], [], []
    for key in sorted(b_map.keys() & a_map.keys()):
        b, a = b_map[key], a_map[key]
        if b["supported"] and not a["supported"]:
            newly_flagged.append((key, b, a))
        elif not b["supported"] and a["supported"]:
            cleared.append((key, b, a))
        elif b["category"] != a["category"]:
            recategorized.append((key, b, a))
    disappeared = sorted(b_map.keys() - a_map.keys())
    appeared = sorted(a_map.keys() - b_map.keys())

    def show(title, rows):
        print(f"\n  {title}: {len(rows)}")
        for (account, clause), b, a in rows:
            print(f"    {account}: {clause[:80]}")
            print(f"        {b['category']}/{b['supported']} -> {a['category']}/{a['supported']}")
            print(f"        was: {_clause_key(b['reason'])[:100]}")
            print(f"        now: {_clause_key(a['reason'])[:100]}")

    show("CLEARED (unsupported -> supported)", cleared)
    show("NEWLY FLAGGED (supported -> unsupported)", newly_flagged)
    show("RECATEGORIZED (verdict same, category moved)", recategorized)
    for title, keys in (("CLAUSES ONLY IN BEFORE", disappeared), ("CLAUSES ONLY IN AFTER", appeared)):
        print(f"\n  {title}: {len(keys)}")
        for account, clause in keys:
            print(f"    {account}: {clause[:80]}")

    b_uns = sum(1 for r in before["verdicts"] if not r["supported"])
    a_uns = sum(1 for r in after["verdicts"] if not r["supported"])
    print("\n  NET")
    print(f"    reviews     {len(before['verdicts'])} -> {len(after['verdicts'])}")
    print(f"    unsupported {b_uns} -> {a_uns}  ({a_uns - b_uns:+d})")
    if not (cleared or newly_flagged or recategorized or appeared or disappeared):
        print("    VERDICT-IDENTICAL — nothing changed.")


# --------------------------------------------------------------------------
# decoys — the discrimination test
# --------------------------------------------------------------------------

def real_cell_values(df) -> List[float]:
    """The account's own numeric cells, the nested analysis frame included.

    Deliberately NOT SourceIndex's values: those already contain column totals
    and every 2-4 row adjacent window sum, so a decoy built from one would be
    testing the pool against its own derived arithmetic. A decoy must start from
    a figure a human could actually have read off the sheet.
    """
    frames = [df]
    analysis = df.attrs.get("prompt_analysis_df") if hasattr(df, "attrs") else None
    if analysis is not None and hasattr(analysis, "columns"):
        frames.append(analysis)
    values: List[float] = []
    for frame in frames:
        for col in frame.columns:
            if col == INTERNAL_ROW_KEY:
                continue
            for cell in frame[col].tolist():
                v = _to_float(cell) if isinstance(cell, (int, float, str)) else None
                # Below 1000 the pool's flat 500 tolerance floor makes almost any
                # deformation land inside the window by construction; those are a
                # known, deliberate design choice (万-rounding), not a finding.
                if v is not None and abs(v) >= 1000:
                    values.append(float(v))
    return values


# Each mode is (label, factor sampler). A pool that discriminates should accept
# almost none of these; the share it accepts is the false-accept rate, and that
# number is how a pool change is graded — a change that clears real flags but
# also raises this is buying its clean run by accepting anything.
DECOY_MODES = {
    "scale_x10": lambda rng: 10.0,
    "scale_div10": lambda rng: 0.1,
    "transposed_digits": lambda rng: rng.uniform(1.09, 1.30),
    "wrong_by_half": lambda rng: 0.5,
    "random_jitter": lambda rng: rng.choice([-1, 1]) * rng.uniform(0.20, 0.80) + 1.0,
}


def run_decoys(
    results: Dict[str, Any],
    dfs: Dict[str, Any],
    *,
    seed: int,
    per_account: int,
) -> None:
    prompt_manager = get_prompt_engine()
    rng = random.Random(seed)
    per_mode = Counter()
    per_mode_accepted = Counter()
    per_account_rows = []

    print(f"\n--- DECOY DISCRIMINATION TEST (seed={seed}, up to {per_account} real values per account) ---")
    for key in sorted(results):
        df = dfs.get(key)
        if df is None:
            continue
        siblings = sibling_dfs_for_account(key, dfs, prompt_manager)
        source = SourceIndex.from_df(df, sibling_dfs=siblings)
        values = real_cell_values(df)
        if not values:
            continue
        # Sanity leg: the pool MUST accept the undeformed value it was built
        # from. If this is not 100%, the decoy result below is meaningless.
        sample = values if len(values) <= per_account else rng.sample(values, per_account)
        truth_accepted = sum(1 for v in sample if source.matches(abs(v)))
        tried = accepted = 0
        for v in sample:
            for mode, sampler in DECOY_MODES.items():
                decoy = abs(v) * sampler(rng)
                hit = source.matches(decoy)
                tried += 1
                accepted += int(hit)
                per_mode[mode] += 1
                per_mode_accepted[mode] += int(hit)
        per_account_rows.append(
            (key, len(siblings or []), len(source.values), len(sample), truth_accepted, tried, accepted)
        )

    # pool is printed because it is the explanatory variable: the pool holds every
    # cell, every column total and every 2-4 row adjacent window sum of this df,
    # its nested analysis frame, AND all same-statement siblings. On a BS account
    # of the <databook> book that is ~10,400 values against ~660 for the account's
    # own frame — at ±5% tolerance a pool that size covers the number line densely
    # enough that "not in source" stops being a meaningful statement.
    print(f"\n  {'account':<28} {'sibs':>4} {'pool':>7} {'sampled':>7} {'truth✓':>7} "
          f"{'decoys':>7} {'accepted':>9} {'rate':>7}")
    for key, sibs, pool, n, truth, tried, accepted in per_account_rows:
        print(f"  {key[:28]:<28} {sibs:>4} {pool:>7} {n:>7} {truth:>7} {tried:>7} {accepted:>9} "
              f"{100.0 * accepted / max(tried, 1):>6.1f}%")

    total_tried = sum(per_mode.values())
    total_accepted = sum(per_mode_accepted.values())
    print(f"\n  by deformation mode:")
    for mode in DECOY_MODES:
        tried, accepted = per_mode[mode], per_mode_accepted[mode]
        print(f"    {mode:<20} {accepted:>5}/{tried:<5} accepted  "
              f"{100.0 * accepted / max(tried, 1):>5.1f}%")
    print(f"\n  OVERALL FALSE-ACCEPT RATE: {total_accepted}/{total_tried} = "
          f"{100.0 * total_accepted / max(total_tried, 1):.1f}%")
    truth_total = sum(r[4] for r in per_account_rows)
    truth_n = sum(r[3] for r in per_account_rows)
    print(f"  (sanity: undeformed real values accepted {truth_total}/{truth_n} = "
          f"{100.0 * truth_total / max(truth_n, 1):.1f}% — must be 100%)")


# --------------------------------------------------------------------------

def main() -> None:
    ap = argparse.ArgumentParser(
        description="Replay the deterministic verifier over an archived run.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    ap.add_argument("databook", nargs="?",
                    help="path to the databook the run was produced from (REQUIRED except for --diff)")
    ap.add_argument("--run", help="run folder: a path, 'run_20260713_113545', or '20260713_113545'")
    ap.add_argument("--entity", default="", help="entity name filter, if the workbook has multiple entities")
    ap.add_argument("--sheet", default=None, help="Financials sheet name (auto-resolved when omitted)")
    ap.add_argument("--snapshot", help="write the verdicts to this JSON path")
    ap.add_argument("--diff", nargs=2, metavar=("BEFORE", "AFTER"),
                    help="compare two snapshots and exit; needs no databook")
    ap.add_argument("--decoys", action="store_true", help="run the decoy discrimination test")
    ap.add_argument("--seed", type=int, default=1234, help="decoy RNG seed (default 1234)")
    ap.add_argument("--decoy-samples", type=int, default=25,
                    help="real values sampled per account for decoys (default 25)")
    ap.add_argument("--no-llm-reviews", action="store_true",
                    help="replay with the deterministic layer only, ignoring the archived reviews")
    ap.add_argument("--quiet-table", action="store_true", help="totals only, skip the per-clause table")
    args = ap.parse_args()

    if args.diff:
        diff_snapshots(*args.diff)
        return

    if not args.databook:
        ap.error("the databook path is required — an archived run carries no workbook identity, "
                 "and pairing the wrong one produces a silently mismatched diff")
    if not args.run:
        ap.error("--run is required (see: python ad-hoc/workbench/list_archived_runs.py)")
    if not os.path.exists(args.databook):
        sys.exit(f"❌ No such databook: {args.databook}")

    run_dir = resolve_run_dir(args.run)
    results = load_results(run_dir)
    print(f"run:      {run_dir}  ({len(results)} accounts)")
    print(f"databook: {args.databook}")

    started = time.perf_counter()
    dfs = build_dfs(args.databook, args.entity, args.sheet)
    print(f"rebuilt {len(dfs)} dfs in {time.perf_counter() - started:.1f}s")
    check_pairing(results, dfs, args.databook, run_dir)

    records = replay_run(results, dfs, use_archived_llm_reviews=not args.no_llm_reviews)
    agree, comparable = archived_agreement(results, records)
    print(f"\nbaseline fidelity: replay reproduces {agree}/{comparable} archived verdicts "
          f"({100.0 * agree / max(comparable, 1):.1f}%) — a gap here is code drift since the run, "
          f"not the change under test")

    if args.quiet_table:
        unsupported = [r for r in records if not r["supported"]]
        print(f"\nreviews={len(records)} unsupported={len(unsupported)} "
              f"by_category={dict(Counter(r['category'] for r in records))}")
    else:
        print_defect_table(records)

    if args.snapshot:
        write_snapshot(args.snapshot, records, {
            "run": run_dir.name,
            "databook": os.path.abspath(args.databook),
            "entity": args.entity,
            "llm_reviews": not args.no_llm_reviews,
            "captured": time.strftime("%Y-%m-%d %H:%M:%S"),
        })

    if args.decoys:
        run_decoys(results, dfs, seed=args.seed, per_account=args.decoy_samples)


if __name__ == "__main__":
    main()
