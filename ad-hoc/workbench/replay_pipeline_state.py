#!/usr/bin/env python3
"""Run the REAL stage loop with the LLM replaced by an archived run — free, offline.

replay_verification.py re-runs the verifier over stored FINAL text. It never
executes the stage loop, so it produces no transitions, no selective-Validator
decision, no defect gate and no arbiter. This tool does: it monkeypatches
`_run_ai_call` to hand back each account's archived per-agent output from
`fdd_utils/logs/<run>/data.yml` and then calls the production
`run_ai_pipeline_with_progress` over real dfs. Every stage boundary, every
degradation branch and every retry runs for real; the only thing faked is the
provider. Zero tokens.

Why the archive cannot be replayed the other way round: every archived run
predates the current code (the selective Validator, Auditor-side grounding and
the deterministic sweep all landed after the newest folder here), so an archived
`results.yml` is not a baseline for today's pipeline. Its per-agent OUTPUT still
is — that is text the model actually produced for this account — which is
exactly what this tool feeds back in.

The databook is a REQUIRED positional argument, for the same reason it is in
replay_verification.py: a run folder carries no workbook identity, and pairing
the wrong databook does not fail, it silently grounds one project's commentary
against another project's numbers.

What it asserts (plan M1's acceptance list):
  1. every account reaches a terminal phase
  2. every transition's `from` matches the previous entry's `to`, and the first
     `from` is PLANNED
  3. stages_done eligibility picks exactly the accounts the old
     `previous_agent not in results[key]` membership test picked, at every stage
     (run_agent_stage computes both and tallies the comparison in
     pipeline._ELIGIBILITY_PARITY)
  4. the deterministic sweep records ZERO transitions — it `continue`s on any
     account that already has clause_reviews, and the Auditor attaches them to
     every account it processes, so a non-zero count means a stage failed
  5. audit.jsonl is one parseable JSON object per account

Usage (PYTHONPATH=. from the repo root):
    python ad-hoc/workbench/replay_pipeline_state.py <databook.xlsx> --run <run_id>
    python ad-hoc/workbench/replay_pipeline_state.py <databook.xlsx> --run <run_id> --out /tmp/replay
"""

from __future__ import annotations

import argparse
import json
import sys
import tempfile
import threading
from collections import Counter
from pathlib import Path
from typing import Any, Dict, Tuple

import yaml

REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))
if str(Path(__file__).resolve().parent) not in sys.path:
    sys.path.insert(0, str(Path(__file__).resolve().parent))

import fdd_utils.ai.pipeline as pipeline  # noqa: E402
from fdd_utils.ai.logging import PipelineRunLogger  # noqa: E402
# build_dfs / resolve_run_dir are the SAME loaders replay_verification uses.
# Duplicating them would let the two tools drift onto different source data and
# quietly stop being comparable.
from replay_verification import build_dfs, resolve_run_dir  # noqa: E402


# --------------------------------------------------------------------------
# the archive, as a fake provider
# --------------------------------------------------------------------------

def load_archived_outputs(run_dir: Path) -> Tuple[Dict[Tuple[str, str], str], Dict[str, Any]]:
    """{(mapping_key, agent_name): output} out of data.yml, plus run metadata.

    `output` is logging.py's own record of the content the provider returned
    (log_agent_complete), so feeding it back through _run_ai_call reproduces the
    exact bytes the pipeline parsed the first time — before clean_agent_output,
    before humanise_*, before verify_commentary.
    """
    path = run_dir / "data.yml"
    if not path.exists():
        sys.exit(f"❌ {path} does not exist — this run archived no per-agent output.")
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    outputs: Dict[Tuple[str, str], str] = {}
    # Accounts come from processing_results, NOT from the outputs: a DEAD run
    # (every call a hard 400) archives 26 accounts and zero outputs, and that is
    # the most useful run in the archive for exercising the degradation
    # branches. Reading the account list off the outputs would silently reduce
    # it to nothing.
    meta: Dict[str, Any] = {"language": None, "model_type": None, "accounts": []}
    for key, agents in (data.get("processing_results") or {}).items():
        if not isinstance(agents, dict):
            continue
        meta["accounts"].append(str(key))
        for agent, record in agents.items():
            if not isinstance(record, dict) or record.get("error"):
                continue
            content = record.get("output")
            if isinstance(content, str) and content.strip():
                outputs[(str(key), str(agent))] = content
                meta["language"] = meta["language"] or record.get("language")
                meta["model_type"] = meta["model_type"] or record.get("model_type")
    return outputs, meta


class ArchiveProvider:
    """Stands in for _run_ai_call. Thread-safe; counts what it could not serve."""

    def __init__(self, outputs: Dict[Tuple[str, str], str]) -> None:
        self._outputs = outputs
        self._lock = threading.Lock()
        self.served = 0
        self.missing: Counter = Counter()
        # _run_ai_call is given no mapping_key — only the two prompt strings —
        # so the account is carried across from the wrapper below on a
        # thread-local. Every LLM call in this pipeline happens inside
        # process_single_agent_item on the calling thread, so this is exact.
        self.local = threading.local()

    def __call__(self, ai_helper, user_prompt, system_prompt, agent_name, timeout=30):
        key = getattr(self.local, "mapping_key", None)
        content = self._outputs.get((str(key), str(agent_name)))
        if content is None:
            with self._lock:
                self.missing["%s/%s" % (key, agent_name)] += 1
            raise RuntimeError(
                "archive has no %s output for %r" % (agent_name, key)
            )
        with self._lock:
            self.served += 1
        return {
            "content": content,
            "duration": 0.0,
            "tokens_used": 0,
            "mode": "archive-replay",
            "model": "archive",
            "model_type": "archive",
            "language": getattr(ai_helper, "language", None),
            # PipelineRunLogger.finalize sums these across every record with a
            # bare sum(), so a None here is a TypeError at the very end of the
            # run. A real provider response always carries them.
            "estimated_prompt_tokens": 0,
            "estimated_output_tokens": 0,
            "estimated_total_tokens": 0,
        }


def install_patches(outputs: Dict[Tuple[str, str], str], out_dir: Path) -> ArchiveProvider:
    provider = ArchiveProvider(outputs)
    real_item = pipeline.process_single_agent_item

    def traced_item(agent_name, mapping_key, *args, **kwargs):
        # A pass-through. Its only job is to publish the account to the thread
        # the LLM call will be made on; production behaviour is untouched.
        provider.local.mapping_key = mapping_key
        return real_item(agent_name, mapping_key, *args, **kwargs)

    class ScratchLogger(PipelineRunLogger):
        """Keeps replay folders out of fdd_utils/logs — that directory is the
        corpus, and salting it with synthetic runs would corrupt every future
        measurement taken from it."""

        def __init__(self, debug_mode: bool = False, **_kw: Any) -> None:
            super().__init__(
                log_dir=str(out_dir), output_dir=str(out_dir), debug_mode=debug_mode,
            )

    pipeline._run_ai_call = provider
    pipeline.process_single_agent_item = traced_item
    pipeline.PipelineRunLogger = ScratchLogger
    return provider


# --------------------------------------------------------------------------
# assertions
# --------------------------------------------------------------------------

def check_terminal(state: Dict[str, Any]) -> Tuple[bool, str]:
    stranded = sorted(
        key for key, acct in state["accounts"].items() if not acct["terminal"]
    )
    return (not stranded), ("every account terminal" if not stranded else f"stranded: {stranded}")


def check_path_continuity(state: Dict[str, Any]) -> Tuple[bool, str]:
    breaks = []
    for key, acct in sorted(state["accounts"].items()):
        expected = "PLANNED"
        for index, hop in enumerate(acct["transitions"]):
            if hop["from"] != expected:
                breaks.append("%s[%s]: from=%s expected=%s" % (key, index, hop["from"], expected))
                break
            expected = hop["to"]
        if not breaks and expected != acct["phase"]:
            breaks.append("%s: path ends at %s but phase is %s" % (key, expected, acct["phase"]))
    return (not breaks), ("all paths continuous" if not breaks else "; ".join(breaks[:5]))


def check_eligibility_parity(parity: Dict[str, int]) -> Tuple[bool, str]:
    ok = parity["mismatches"] == 0 and parity["checks"] > 0
    return ok, "%s stage eligibility check(s), %s mismatch(es)" % (
        parity["checks"], parity["mismatches"],
    )


def check_no_deterministic_sweep(state: Dict[str, Any], healthy: bool) -> Tuple[bool, str]:
    """Zero is the assertion ON A HEALTHY RUN only.

    The sweep `continue`s on any account that already has clause_reviews, and
    the Auditor attaches them to every account it processes — so a non-zero
    count means a stage failed. Replay a DEAD archived run and every account
    goes through it, correctly: that is the sweep doing its job, not a
    regression, so the check reports rather than fails.
    """
    count = sum(
        1
        for acct in state["accounts"].values()
        for hop in acct["transitions"]
        if hop["cause"] == "grounding_recovered_after_stage_failure"
    )
    if not healthy:
        return True, "%s (degraded replay — not asserted; the sweep is meant to fire here)" % count
    return count == 0, "deterministic-sweep transitions: %s" % count


#: Keys the pipeline writes that NO archive can contain, because they were
#: added after the newest archived run folder. Listed explicitly so a genuinely
#: new key still fails the check instead of hiding behind a loose baseline.
#: `<agent>_metadata` is where subagent_1's used_fallback/fallback_reason lands
#: — it used to be dropped on the floor in _store_agent_result.
_POST_ARCHIVE_KEYS = {
    "feedback_arbiter",
    "subagent_1_metadata", "subagent_2_metadata", "subagent_4_metadata",
}


def archived_key_vocabulary(limit: int) -> set:
    """Every key name the pipeline has ever written into `results[key]`,
    harvested from the newest `limit` archived results.yml files.

    One run is not a baseline — a run that never triggered the feedback loop
    simply has no `feedback_retry_*` keys, and reading its absence as "this key
    is new" is a false alarm. Nor are the newest forty: MEASURED, none of them
    retried anything, and the `feedback_retry_*` vocabulary only appears deeper
    in the archive. So the default reads all of it (~16s, offline, free);
    `limit > 0` trades coverage for speed and will produce false alarms.
    """
    known = set(_POST_ARCHIVE_KEYS)
    paths = sorted((REPO_ROOT / "fdd_utils" / "logs").glob("run_*/results.yml"), reverse=True)
    for path in (paths[:limit] if limit > 0 else paths):
        try:
            data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
        except Exception:
            continue
        for record in data.values():
            if isinstance(record, dict):
                known.update(record.keys())
    return known


def check_results_shape(results: Dict[str, Any], vocab: set) -> Tuple[bool, str]:
    """`results[key]` must be exactly what it was before state tracking existed.

    Nine consumers read that dict, so the milestone's whole claim is that it is
    untouched. Two specific ways this could go wrong, both checked here:
    `metadata["phase_events"]` leaking through into a stored record (it is
    popped in _store_agent_result, and again at the three sites that bypass it),
    and a new key name appearing.
    """
    leaked = []
    seen: set = set()
    for key, record in results.items():
        if key.startswith("__") or not isinstance(record, dict):
            continue
        seen.update(record.keys())
        if "phase_events" in record:
            leaked.append(key)
        for value in record.values():
            if isinstance(value, dict) and "phase_events" in value:
                leaked.append("%s.<metadata>" % key)
    if leaked:
        return False, "phase_events leaked into results for: %s" % sorted(set(leaked))

    novel = sorted(seen - vocab)
    if novel:
        return False, "new key(s) in results[key]: %s" % novel
    return True, "%s key name(s), all known to the archive" % len(seen)


def check_audit_log(path: str, expected: int) -> Tuple[bool, str]:
    if not path or not Path(path).exists():
        return False, "audit.jsonl not written"
    lines = [ln for ln in Path(path).read_text(encoding="utf-8").splitlines() if ln.strip()]
    try:
        parsed = [json.loads(ln) for ln in lines]
    except Exception as exc:
        return False, "audit.jsonl is not parseable: %s" % exc
    if len(parsed) != expected:
        return False, "audit.jsonl has %s line(s), expected %s" % (len(parsed), expected)
    missing = [p.get("account") for p in parsed if not p.get("transitions")]
    if missing:
        return False, "line(s) with no state path: %s" % missing
    return True, "%s parseable account path(s)" % len(parsed)


# --------------------------------------------------------------------------

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("databook", help="the .xlsx this run was produced from (REQUIRED — see module docstring)")
    ap.add_argument("--run", required=True, help="run folder, folder name, or bare timestamp")
    ap.add_argument("--entity", default="", help="entity name passed to process_workbook_data")
    ap.add_argument("--sheet", default=None, help="Financials sheet; auto-resolved when omitted")
    ap.add_argument("--out", default=None, help="where the replay's run folder goes (default: a temp dir)")
    ap.add_argument("--no-threads", action="store_true", help="run the stages serially")
    ap.add_argument("--vocab-runs", type=int, default=0,
                    help="how many archived results.yml to read for the results-key baseline (0 = all)")
    args = ap.parse_args()

    run_dir = resolve_run_dir(args.run)
    outputs, meta = load_archived_outputs(run_dir)
    archived_keys = sorted(set(meta["accounts"]) | {key for key, _agent in outputs})
    print(f"archive: {run_dir.name} — {len(archived_keys)} account(s), {len(outputs)} agent output(s)")

    dfs = build_dfs(args.databook, args.entity, args.sheet)
    mapping_keys = [k for k in archived_keys if k in dfs]
    unpaired = [k for k in archived_keys if k not in dfs]
    if not mapping_keys:
        sys.exit(
            "❌ REFUSING TO REPLAY — not one archived account has a df in this databook.\n"
            f"   archive: {archived_keys}\n   databook: {sorted(dfs)}"
        )
    if unpaired:
        print(f"⚠️  {len(unpaired)} archived account(s) absent from this databook, skipped: {unpaired}")

    out_dir = Path(args.out) if args.out else Path(tempfile.mkdtemp(prefix="replay_pipeline_state_"))
    out_dir.mkdir(parents=True, exist_ok=True)
    provider = install_patches(outputs, out_dir)
    pipeline._ELIGIBILITY_PARITY.update({"checks": 0, "mismatches": 0})

    language = meta.get("language") or "Chi"
    model_type = meta.get("model_type") or "deepseek"
    print(f"replaying {len(mapping_keys)} account(s) | language={language} | model_type={model_type}")

    results = pipeline.run_ai_pipeline_with_progress(
        mapping_keys=mapping_keys,
        dfs=dfs,
        model_type=model_type,
        language=language,
        use_multithreading=not args.no_threads,
    )

    run_record = results.get(pipeline.RUN_STATE_KEY) or {}
    state = run_record.get("state") or {}
    health = results.get(pipeline.RUN_HEALTH_KEY) or {}

    print()
    print(f"--- replay run folder: {run_record.get('run_folder')}")
    print(f"provider: {provider.served} archived output(s) served, "
          f"{sum(provider.missing.values())} call(s) the archive could not answer")
    if provider.missing:
        print(f"          unanswered: {dict(provider.missing)}")
    print(f"health:   {health.get('calls_succeeded')} succeeded / {health.get('calls_failed')} failed")
    # A replay of a DEAD archived run degrades on purpose — every branch below
    # the LLM call is exactly what it is there to exercise — so the two
    # healthy-run assertions report instead of failing.
    healthy = bool(health.get("calls_succeeded")) and not state.get("degraded_accounts")
    if not healthy:
        print("          DEGRADED replay: the healthy-run assertions are reported, not enforced")

    print()
    print("phase tally:      %s" % state.get("phase_tally"))
    print("degraded:         %s" % (state.get("degraded_accounts") or "none"))
    print("defect codes:     %s" % (run_record.get("defect_codes") or "none"))
    causes = Counter(
        hop["cause"]
        for acct in state.get("accounts", {}).values()
        for hop in acct["transitions"]
    )
    print("transition causes:")
    for cause, count in causes.most_common():
        print("   %4d  %s" % (count, cause))

    checks = [
        ("every account reaches a terminal phase", check_terminal(state)),
        ("transition `from` matches previous `to`", check_path_continuity(state)),
        ("stages_done eligibility == old membership test",
         check_eligibility_parity(run_record.get("eligibility_parity") or {"checks": 0, "mismatches": 1})),
        ("deterministic-sweep transitions == 0 (healthy)",
         check_no_deterministic_sweep(state, healthy)),
        ("audit.jsonl is one state path per account",
         check_audit_log(run_record.get("audit_log") or "", len(state.get("accounts") or {}))),
        ("results[key] unchanged (no leak, no new keys)",
         check_results_shape(results, archived_key_vocabulary(args.vocab_runs))),
    ]
    print()
    failed = 0
    for label, (ok, detail) in checks:
        print("%s  %-46s  %s" % ("PASS" if ok else "FAIL", label, detail))
        failed += 0 if ok else 1

    print()
    print(f"audit log: {run_record.get('audit_log')}")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
