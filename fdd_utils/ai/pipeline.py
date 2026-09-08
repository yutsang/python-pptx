from __future__ import annotations

# re-added: bound by an import in another section of the pre-split module
from ..financial_common import (
    get_pipeline_result_text,
    load_yaml_file,
    package_file_path,
    visible_descriptions,
)
import time

"""
Unified AI pipeline and prompt-loading surface for FDD.
"""

from .config import FDDConfig, get_safe_default_data_format, normalize_language_code
from .english import _iso_to_long_date, polish_english_commentary
from .validator import format_validator_feedback_for_reprompt, parse_validator_response, strip_thinking, verify_commentary
from .logging import PipelineRunLogger, coerce_plain
from .prompts import PromptEngine, PromptStylePack, _DEFAULT_MAPPINGS_FILE, _DEFAULT_PROMPTS_FILE, get_prompt_engine, resolve_prompt_asset_path
from .client import AIClient
from .repair import repair_content, repairable_reviews


import hashlib
import multiprocessing
import os
import re
import threading
import concurrent.futures
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Any, Callable, Dict, List, Optional, Tuple

import pandas as pd
import yaml


class _StageCircuitBreaker:
    """Per-stage circuit breaker for LLM calls.

    Tracks consecutive failures per agent stage. After ``threshold`` consecutive
    failures (across all worker threads), the breaker is OPEN for that stage —
    further calls raise immediately instead of waiting for retries to time out.
    The breaker resets on the first success or when ``reset_stage`` is called
    at the start of each new stage.

    Rationale: when the LLM endpoint is stressed, every call wastes
    ``timeout × retries`` (30s × 3 = 90s) before falling back. After 3-5
    failures in a row that's clearly the new normal — fail fast for the rest
    of the batch and let the deterministic fallback take over.
    """

    def __init__(self, threshold: int = 4):
        self._threshold = threshold
        self._consecutive_failures: Dict[str, int] = {}
        self._lock = threading.Lock()

    def record_success(self, stage: str) -> None:
        with self._lock:
            self._consecutive_failures[stage] = 0

    def record_failure(self, stage: str) -> None:
        with self._lock:
            self._consecutive_failures[stage] = self._consecutive_failures.get(stage, 0) + 1

    def is_open(self, stage: str) -> bool:
        with self._lock:
            return self._consecutive_failures.get(stage, 0) >= self._threshold

    def reset_stage(self, stage: str) -> None:
        with self._lock:
            self._consecutive_failures[stage] = 0


_PIPELINE_BREAKER = _StageCircuitBreaker(threshold=4)


#: Sentinel key under which run_ai_pipeline_with_progress files its health
#: tally on the results dict. Follows the existing __BS_summary__ /
#: __IS_summary__ precedent: it is not a mapping_key, so every consumer that
#: resolves a key through mappings.yml (build_pptx_structured_payloads,
#: the UI's summary loop) skips it, and the ones that filter on "final"
#: (extract_final_contents) never see it either.
RUN_HEALTH_KEY = "__run_health__"


class _RunHealth:
    """How much of this run's commentary actually came from the LLM.

    Measured across 255 archived runs: five of them finished normally having
    made ZERO successful LLM calls and shipped 17-19 deterministic fallback
    bullets as the deck's entire commentary (run_20260714_223603,
    run_20260715_081202, run_20260715_090949, run_20260717_001202,
    run_20260717_001303). The cause was a hard 400 — a model name the
    provider did not recognise — not rate limiting, so retrying could never
    help. Nothing in the log summary, the return value or the deck told those
    runs apart from a clean one, because ``process_single_agent_item``
    degrades in four separate places and every one of them returns normally.

    Thread-safe: the stages run in a ThreadPoolExecutor.
    """

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self.calls_succeeded = 0
        self.calls_failed = 0  # individual attempts, so retries count separately
        self.calls_skipped_breaker = 0
        self.stage_successes: Dict[str, int] = {}
        self.fallback_accounts: Dict[str, str] = {}
        self.passthrough_accounts: set = set()
        self.error_text_accounts: set = set()
        self.no_prompt_accounts: set = set()
        self.first_failure = ""

    def record_success(self, agent_name: str) -> None:
        with self._lock:
            self.calls_succeeded += 1
            self.stage_successes[agent_name] = self.stage_successes.get(agent_name, 0) + 1

    def record_failed_attempt(self, exc: BaseException) -> None:
        with self._lock:
            self.calls_failed += 1
            if not self.first_failure:
                # The first failure is the diagnostic one. In all five dead
                # runs every later failure was a verbatim copy of it.
                self.first_failure = str(exc)[:200]

    def record_breaker_skip(self) -> None:
        with self._lock:
            self.calls_skipped_breaker += 1

    def record_fallback(self, mapping_key: str, reason: str) -> None:
        with self._lock:
            self.fallback_accounts[mapping_key] = reason

    def record_passthrough(self, mapping_key: str) -> None:
        with self._lock:
            self.passthrough_accounts.add(mapping_key)

    def record_error_text(self, mapping_key: str) -> None:
        with self._lock:
            self.error_text_accounts.add(mapping_key)

    def record_no_prompt(self, mapping_key: str) -> None:
        with self._lock:
            self.no_prompt_accounts.add(mapping_key)

    def stage_succeeded(self, agent_name: str) -> int:
        with self._lock:
            return self.stage_successes.get(agent_name, 0)

    def as_dict(self) -> Dict[str, Any]:
        """Plain types only — this rides on the results dict, which is
        yaml.dump'd into both results.yml and the archived run log."""
        with self._lock:
            return {
                "calls_succeeded": self.calls_succeeded,
                "calls_failed": self.calls_failed,
                "calls_skipped_breaker": self.calls_skipped_breaker,
                "stage_successes": dict(self.stage_successes),
                "accounts_on_fallback_bullet": sorted(self.fallback_accounts),
                "accounts_on_passthrough": sorted(self.passthrough_accounts),
                "accounts_on_error_text": sorted(self.error_text_accounts),
                "accounts_without_prompt": sorted(self.no_prompt_accounts),
                "first_failure": self.first_failure,
                "zero_successful_calls": self.calls_succeeded == 0,
                # The one flag a caller may refuse to export on. Deliberately
                # NOT an exception: a degraded deck is sometimes exactly what
                # the user asked to look at, so the decision stays with them.
                "safe_to_export": self.calls_succeeded > 0,
            }


def _log_run_health(logger: PipelineRunLogger, health: Dict[str, Any], total_items: int) -> None:
    """One line saying how much of the run was real. ERROR when none of it was."""
    line = (
        "Run health: %s LLM call(s) succeeded, %s attempt(s) failed, %s skipped by an "
        "open circuit breaker | of %s account(s): %s on a deterministic fallback bullet, "
        "%s on stage passthrough, %s on error placeholder text, %s with no prompt"
        % (
            health["calls_succeeded"],
            health["calls_failed"],
            health["calls_skipped_breaker"],
            total_items,
            len(health["accounts_on_fallback_bullet"]),
            len(health["accounts_on_passthrough"]),
            len(health["accounts_on_error_text"]),
            len(health["accounts_without_prompt"]),
        )
    )
    if health["zero_successful_calls"]:
        logger.logger.error(
            "%s | NOT ONE LLM CALL SUCCEEDED — every bullet in this run is deterministic "
            "filler and the deck is not a deliverable. First failure: %s",
            line, health["first_failure"] or "none recorded",
        )
    else:
        logger.logger.info(line)


#: Sentinel key under which run_ai_pipeline_with_progress files the run folder
#: and the serialized RunState. Same convention and the same safety argument as
#: RUN_HEALTH_KEY above: it is not a mapping_key, so build_pptx_structured_payloads
#: (find_mapping_key returns None → continue) and the UI's summary loop skip it,
#: and it carries no "final" so extract_final_contents never sees it.
#:
#: It exists because run_ai_pipeline_with_progress returns ONLY `results` and
#: never exposes the run folder, so without a sentinel no caller could reach the
#: state or the audit log at all.
RUN_STATE_KEY = "__run__"

#: Running proof that the stages_done eligibility test in run_agent_stage picks
#: exactly the accounts the `previous_agent not in results[key]` membership
#: probe used to. Read by ad-hoc/workbench/replay_pipeline_state.py; a mismatch
#: is also logged at ERROR and falls back to the old probe.
_ELIGIBILITY_PARITY = {"checks": 0, "mismatches": 0}


# ---------------------------------------------------------------------------
# Explicit run state (plan M1 steps 1-4)
# ---------------------------------------------------------------------------
#
# TWO AXES, DELIBERATELY SEPARATED — collapsing them is what made an earlier
# draft of this unimplementable.
#
#   stages_done  which stages have run. Order-independent, and a direct
#                restatement of the membership test run_agent_stage used to do
#                against `results[key]`. This is what stage ELIGIBILITY reads.
#   phase        the state of the account's TEXT. This is what a router reads.
#
# A single phase test cannot decide eligibility: the Generator's entry phase is
# PLANNED, the Auditor's is DRAFTED and the Validator's would be GROUNDED, and
# _get_agent_stage_context has nothing to derive the requirement from.
PHASE_PLANNED = "PLANNED"
PHASE_DRAFTED = "DRAFTED"
#: GROUNDED means "arithmetic has been applied", and it is reached at the
#: AUDITOR stage, not after it: verify_commentary fires for subagent_2 as well
#: as subagent_4 (see process_single_agent_item), so an account is grounded one
#: stage earlier than a naive draft→audit→verify reading assumes. The Validator
#: therefore RE-ENTERS GROUNDED (a self-edge) rather than advancing; whether the
#: LLM Validator also ran is a transition CAUSE, not a separate phase. There is
#: deliberately no AUDITED or VERIFIED phase — "which stages ran" is stages_done.
PHASE_GROUNDED = "GROUNDED"
PHASE_DEFECTIVE = "DEFECTIVE"
#: Unreachable until the local-repair milestone (M4) lands. Named now so the
#: vocabulary is fixed and the audit log does not change shape when it does.
PHASE_REPAIRING = "REPAIRING"
PHASE_REGENERATING = "REGENERATING"
PHASE_ACCEPTED = "ACCEPTED"
PHASE_ARBITRATED = "ARBITRATED"
#: FAILED is ONLY the two placeholder-text branches at the bottom of
#: process_single_agent_item ("Content generation failed / incomplete for ...").
#: The deterministic fallback bullet is NOT a failure: that account goes on
#: through the Auditor and Validator and produces real output, so it stays on
#: the normal path with a `degraded` cause recorded instead. Marking it FAILED
#: would both lose it from the ACCEPTED tally and satisfy "every account reached
#: a terminal phase" for the wrong reason. Note even the placeholder text is
#: stored like any other content and the account keeps going — FAILED is a label
#: on the outcome, not a stop.
PHASE_FAILED = "FAILED"

#: DEFECTIVE is in here because the final sweep uses it as a VERDICT: reviews
#: present, defects present, and the feedback loop either disabled or unable to
#: clear them. Mid-run it is not terminal — it is the gate the retry path leaves
#: from — but nothing is mid-run by the time the sweep has finished.
TERMINAL_PHASES = frozenset({PHASE_ACCEPTED, PHASE_ARBITRATED, PHASE_FAILED, PHASE_DEFECTIVE})


class AccountState:
    """The *why* half of one account's run. `results[key]` remains the *what*.

    `results[key]` is untouched by design — nine consumers read it — so nothing
    here duplicates its contents. What lives here is the record the dict cannot
    hold: the ordered path the account took, why each step happened, and how it
    degraded along the way.

    THREADING: exactly one thread ever writes one AccountState. In the stage
    loop that is the main thread (workers accumulate metadata["phase_events"]
    and _store_agent_result drains them — see process_single_agent_item); in the
    feedback loop it is that key's own worker, which no other thread shares.
    """

    def __init__(self, mapping_key: str) -> None:
        self.mapping_key = mapping_key
        self.stages_done: set = set()
        self.phase = PHASE_PLANNED
        #: Non-terminal degradation: fallback_bullet | stage_passthrough_after_error
        #: | no_prompt | circuit_breaker. The account still produced output.
        self.degraded: List[str] = []
        self.attempts: List[Dict[str, Any]] = []
        self.defects: List[Dict[str, Any]] = []
        self.repairs: List[Dict[str, Any]] = []  # M4
        self.contract: Dict[str, Any] = {}       # M5
        self.transitions: List[Dict[str, str]] = []

    def transition(self, to_phase: str, cause: str) -> None:
        """Record a move. `from` is always the current phase, so a path read
        top to bottom is always continuous — that continuity is the audit
        log's only structural guarantee and it is free as long as nothing
        assigns `phase` directly."""
        self.transitions.append({"from": self.phase, "to": to_phase, "cause": str(cause)})
        self.phase = to_phase

    def note_degraded(self, cause: str) -> None:
        cause = str(cause)
        if cause not in self.degraded:
            self.degraded.append(cause)

    def mark_stage(self, agent_name: str) -> None:
        self.stages_done.add(agent_name)

    def as_dict(self) -> Dict[str, Any]:
        return coerce_plain({
            "mapping_key": self.mapping_key,
            "phase": self.phase,
            "terminal": self.phase in TERMINAL_PHASES,
            "stages_done": sorted(self.stages_done),
            "degraded": list(self.degraded),
            "transitions": list(self.transitions),
            "attempts": list(self.attempts),
            "defects": list(self.defects),
            "repairs": list(self.repairs),
            "contract": dict(self.contract),
        })


class RunState:
    """Every derived fact for the duration of ONE run, plus one slot per account.

    NOT cross-run memory — nothing here survives the process, and nothing from a
    previous run influences this one. It also unlocks NO resumability: `results`
    and `run_data` are both written once, in PipelineRunLogger.finalize, so a run
    that dies at the eighteenth account still restarts from zero. An in-memory
    RunState dies with the process exactly as `results` does. It makes a resume
    predicate expressible, nothing more.

    The fact half is built ONCE, before the stage loop, and is read-only during
    it. Most of it is reserved for later milestones and is empty today; the names
    are fixed now so the shape does not move under them:

        profile    the workbook semantic profile (N1)
        siblings   {key -> [df]}, replacing the two duplicated build sites (M2)
        evidence   {key -> EvidenceIndex} (M2)
        facts      {(key, ISO period) -> total} (N2)
        links      verified cross-account edges, each with its numeric test (N2)

    `peer` is built here rather than per call site because _build_peer_context is
    silently None on the retry and reprompt paths today — but the call sites are
    NOT switched over in this milestone: doing so changes prompts, and this
    milestone's gate is verdict-identical output.
    """

    def __init__(
        self,
        dfs: Optional[Dict[str, pd.DataFrame]],
        mapping_keys: List[str],
        run_folder: str = "",
    ) -> None:
        # ---- fact half ----
        self.dfs = dfs or {}
        self.profile: Optional[Dict[str, Any]] = None
        self.siblings: Dict[str, Any] = {}
        self.evidence: Dict[str, Any] = {}
        self.facts: Dict[Any, Any] = {}
        try:
            self.peer = _build_peer_context(dfs)
        except Exception:  # pragma: no cover - defensive; a missing peer is not fatal
            self.peer = None
        self.links: List[Dict[str, Any]] = []

        # ---- process half ----
        # Same key set create_result_shell uses. An account in mapping_keys but
        # absent from dfs gets no result shell today, so it must get no
        # AccountState either — otherwise it sits in PLANNED forever and breaks
        # the "every account reached a terminal phase" check for a reason that
        # has nothing to do with the pipeline.
        self.accounts: Dict[str, AccountState] = {
            key: AccountState(key) for key in mapping_keys if key in self.dfs
        }
        self.deck: Dict[str, Any] = {}
        self.run_folder = run_folder

    def account(self, mapping_key: str) -> Optional[AccountState]:
        return self.accounts.get(mapping_key)

    def phase_tally(self) -> Dict[str, int]:
        return dict(Counter(state.phase for state in self.accounts.values()))

    def transition_causes(self) -> Dict[str, int]:
        return dict(Counter(
            t["cause"] for state in self.accounts.values() for t in state.transitions
        ))

    def as_dict(self) -> Dict[str, Any]:
        return coerce_plain({
            "run_folder": self.run_folder,
            "phase_tally": self.phase_tally(),
            "accounts_not_terminal": sorted(
                key for key, state in self.accounts.items()
                if state.phase not in TERMINAL_PHASES
            ),
            "degraded_accounts": {
                key: list(state.degraded)
                for key, state in self.accounts.items() if state.degraded
            },
            "accounts": {key: state.as_dict() for key, state in self.accounts.items()},
        })


def _record_recovered_grounding(state: Optional[AccountState], cause: str) -> None:
    """Grounding attached by a late recovery pass, without reviving a FAILED
    account.

    MEASURED, replaying an archived dead run (every LLM call a hard 400): 8
    accounts shipped the error-placeholder string, the recovery pass then
    grounded that string — it contains no amount, so it has no unsupported
    clause — and the final sweep read "reviews present, no defect" and called
    all 26 ACCEPTED, on a deck that was entirely deterministic filler. Running
    arithmetic over a placeholder is not grounding; the reviews are still
    attached (the deck's highlighting reads them), only the phase is held.
    """
    if state is None:
        return
    if state.phase == PHASE_FAILED:
        state.transition(PHASE_FAILED, "%s: placeholder text, phase held" % cause)
    else:
        state.transition(PHASE_GROUNDED, cause)


def _queue_phase_event(metadata: Dict[str, Any], **event: Any) -> Dict[str, Any]:
    """Park a state change a WORKER observed, for the main thread to record.

    THE RULE THAT MAKES THIS SAFE: no phase transition is ever recorded from a
    worker thread. verify_commentary and every degradation branch below run
    inside process_single_agent_item on a ThreadPoolExecutor worker and return
    BEFORE _store_agent_result runs on the main thread, so recording at the
    observation site would append transitions out of ORDER for every account —
    the path would stop being a path.

    So the worker queues; _store_agent_result drains, in order, one writer.
    The key is popped back off in _store_agent_result before anything is filed,
    so `results` never sees it and stays byte-identical to what it was.
    """
    metadata = dict(metadata or {})
    metadata.setdefault("phase_events", []).append(event)
    return metadata


# Active pipeline stages, in order. The Refiner (subagent_3 / 3_Refiner) is
# DORMANT by design — its config/prompts are retained for reference but it is
# deliberately omitted here, so the runtime is a 3-stage pipeline despite the
# "4 subagent" naming elsewhere. Add ("subagent_3", "Refiner") between Auditor
# and Validator to re-enable it.
SUBAGENT_SEQUENCE = [
    ("subagent_1", "Generator"),
    ("subagent_2", "Auditor"),
    ("subagent_4", "Validator"),
]

def _get_prompt_manager(prompts_file: str, mappings_file: str) -> PromptEngine:
    resolved_prompts = resolve_prompt_asset_path(
        prompts_file,
        _DEFAULT_PROMPTS_FILE,
        get_prompt_engine().prompts_path,
    )
    resolved_mappings = resolve_prompt_asset_path(
        mappings_file,
        _DEFAULT_MAPPINGS_FILE,
        get_prompt_engine().mappings_path,
    )
    return get_prompt_engine(
        prompts_path=resolved_prompts,
        mappings_path=resolved_mappings,
    )


def map_value_to_component(
    value: str,
    component: Optional[str] = None,
    file_path: str = _DEFAULT_MAPPINGS_FILE,
) -> Any:
    """Look up mapping metadata from the configured mappings file."""
    mappings_path = resolve_prompt_asset_path(
        file_path,
        _DEFAULT_MAPPINGS_FILE,
        get_prompt_engine().mappings_path,
    )
    manager = get_prompt_engine(mappings_path=mappings_path)
    return manager.get_mapping_component(value, component=component)


def load_prompts_and_format(
    agent_name: str,
    language: str,
    mapping_key: str,
    df: pd.DataFrame,
    prompts_file: str = _DEFAULT_PROMPTS_FILE,
    mappings_file: str = _DEFAULT_MAPPINGS_FILE,
    **kwargs,
) -> Tuple[str, str]:
    """Render prompts for a specific agent/account pair."""
    manager = _get_prompt_manager(prompts_file, mappings_file)
    data_format = get_safe_default_data_format(language=language)
    return manager.render_prompt(
        agent_name=agent_name,
        language=language,
        mapping_key=mapping_key,
        df=df,
        data_format=data_format,
        **kwargs,
    )


def clean_agent_output(content: str) -> str:
    """Remove common meta-commentary from agent outputs."""
    content = strip_thinking(content)  # drop any Qwen3 <think> block first
    prefixes_to_remove = [
        r"^verified\s+output:\s*",
        r"^corrected\s+output:\s*",
        r"^refined\s+output:\s*",
        r"^formatted\s+output:\s*",
        r"^final\s+output:\s*",
        r"^after\s+verification[,:]?\s*",
        r"^after\s+refining[,:]?\s*",
        r"^final\s+formatted\s+content:\s*",
        r"^the\s+corrected\s+output\s+is:\s*",
        r"^here\s+is\s+the\s+(corrected|refined|verified)\s+output:\s*",
        r"^已验证输出：\s*",
        r"^已更正输出：\s*",
        r"^精炼后的输出：\s*",
        r"^格式化后的输出：\s*",
        r"^经过验证[，,]\s*",
        r"^经过精炼后[，,]\s*",
    ]
    cleaned = (content or "").strip()
    for pattern in prefixes_to_remove:
        cleaned = re.sub(pattern, "", cleaned, flags=re.IGNORECASE)

    if (cleaned.startswith('"') and cleaned.endswith('"')) or (
        cleaned.startswith("'") and cleaned.endswith("'")
    ):
        cleaned = cleaned[1:-1]

    end_patterns = [
        r"\s*I\s+(?:verified|corrected|refined|checked).*$",
        r"\s*(?:Corrections?|Verifications?)\s+made:.*$",
        r"\s*我(?:验证|更正|精炼|检查)了.*$",
        r"\s*所做更正：.*$",
    ]
    for pattern in end_patterns:
        cleaned = re.sub(pattern, "", cleaned, flags=re.IGNORECASE)
    return cleaned.strip()


# "余额为0元" and its variants. The prompts already forbid this (mappings.yml
# for the Generator, prompts.yml for the Auditor, plus its checklist), but a
# rule that only exists in prose is not a guarantee -- this is the same rule
# as code, applied to every agent's output.
#
# The negative lookahead is what keeps it safe: a real amount that merely
# STARTS with a zero digit (0.5万元) must not match, so the 0 may not be
# followed by another digit, a decimal point, or a full-width period.
# Two guards, both earned on a real seven-entity run.
#
#   * The unit. The old lookahead blocked a digit or a decimal point but not a
#     UNIT, so "分别为0万元" matched at "为0" and the deck shipped
#     "银行手续费分别未发生万元、1.8万元、1.4万元及1.3万元". The zero there is
#     0万元, a figure, not a bare "为0元".
#   * The enumeration. Even with the right unit, substituting one item inside
#     "分别为A、B、C" leaves a list whose first item is a verb phrase. A trailing
#     list separator followed by another figure means this zero is one item of
#     several, and the sentence has to be restructured rather than substituted --
#     which is prompt work, not regex work, and is why the 0万元 rewrite was
#     dropped once already.
# 万/亿/千/百 only. NOT 元 -- "为0元" is exactly the bare-zero form this rule
# exists to rewrite, whereas "为0万元" is a figure that happens to be zero.
_ZERO_UNITS = "万亿千百"
# The enumeration guard is a lookahead placed BEFORE the optional 元, not after
# it. Written as a trailing lookahead it does nothing: the engine simply
# backtracks, gives up the 元 it had consumed, and matches the shorter "为0" --
# turning "分别为0元、1.8万元" into "分别未发生元、1.8万元". Checking from the
# zero itself, with 元 optional inside the lookahead, leaves nothing to
# backtrack into.
_ZERO_NOT_A_LIST = r"(?!\s*元?\s*[、，,]\s*-?[\d])"
_ZERO_BALANCE_RE = re.compile(
    r"(?:余额)?(?:合计)?为(?:人民币)?0(?![\d.．" + _ZERO_UNITS + r"])"
    + _ZERO_NOT_A_LIST + r"\s*(?:元)?"
)
# Same idea for the other mechanical form: "无/未有余额" is fine, "余额为零" is not.
_ZERO_BALANCE_CHAR_RE = re.compile(
    r"(?:余额)?(?:合计)?为(?:人民币)?零(?![" + _ZERO_UNITS + r"])"
    + _ZERO_NOT_A_LIST + r"\s*(?:元)?"
)


# Databook working-note vocabulary that should never reach a deliverable.
# 1:1 replacements only -- no deletion, so there is no way for this to lose a
# fact the way the reverted enumeration dedupe did.
_WORKPAPER_WORDS = (
    (re.compile(r"checking(?=不平|不符|有差异)?"), "核对"),
    (re.compile(r"tie[- ]?out"), "核对"),
    # "管报package" / "管理报表package" -- the English tail is the databook's
    # own filing word and says nothing the Chinese does not already say.
    (re.compile(r"(?<=管报)\s*package"), ""),
    (re.compile(r"(?<=管理报表)\s*package"), ""),
)

# "根据备注" / "根据备注说明" cites the DATABOOK's remarks column. A deliverable
# states the fact; the reader has no access to the working file and does not
# care where it was written down.
#
# "根据管理层说明" / "管理层表示" are deliberately NOT here: a management
# representation is a real, citable source in due diligence, and the prompts
# explicitly ask for those to be kept.
_SOURCE_META_RE = re.compile(r"(?:^|(?<=[。；;\n]))\s*根据备注(?:说明)?[，,、]\s*")

# A negative retained-earnings balance reads as 未弥补亏损, not as a minus sign.
# Anchored on the account name so it cannot touch an ordinary negative amount,
# and it only ever rewrites the sign into words -- the figure is untouched.
_NEGATIVE_RE_RE = re.compile(r"(未分配利润|留存收益)(为|是)(?:人民币)?-\s*([\d,]+(?:\.\d+)?)")

# NOT rewritten here: "分别为16.1万元、-16.1万元、0万元及0万元".
# Turning the zeros into a bare "0" was tried and the result -- "0及0" -- reads
# no better than what it replaced. What a report actually writes is
# "分别为16.1万元及-16.1万元，2025年度及2026年1至3月未发生", which means
# splitting the enumeration and re-attaching the period labels. That is a
# restructure, not a substitution, and restructuring finished sentences is what
# deleted real figures last time (see docs/failed-attempts.md). It stays a
# prompt rule, with a worked example, where the model can see the whole
# sentence and decide.


def humanise_report_language(text: str) -> str:
    """Fixed substitutions for phrasing a deliverable never uses.

    Every rule here is a REPLACEMENT, never a deletion of content -- the one
    exception being the "根据备注" lead-in, which is a citation of the working
    file rather than a fact, and whose removal leaves the sentence intact.
    That restraint is deliberate: the enumeration dedupe that did guess at
    removable text deleted real figures on six accounts and was reverted
    (see docs/failed-attempts.md).
    """
    body = str(text or "")
    if not body:
        return body
    for pattern, replacement in _WORKPAPER_WORDS:
        body = pattern.sub(replacement, body)
    body = _SOURCE_META_RE.sub("", body)
    body = _NEGATIVE_RE_RE.sub(lambda m: f"{m.group(1)}{m.group(2)}未弥补亏损{m.group(3)}", body)
    return body


def humanise_zero_balance(text: str, statement_type: str = "") -> str:
    """Rewrite the mechanical zero-balance forms the deliverable never uses.

    A real report does not write "余额为0元". Balance-sheet accounts read
    "无余额"; income-statement (period) accounts read "未发生". statement_type
    picks between them and defaults to the balance-sheet wording, which is the
    safer of the two to apply to an unknown account -- it is at worst slightly
    less idiomatic, where "未发生" on a balance would be wrong.
    """
    body = str(text or "")
    if not body:
        return body
    replacement = "未发生" if str(statement_type or "").strip().upper() == "IS" else "无余额"
    body = _ZERO_BALANCE_RE.sub(replacement, body)
    body = _ZERO_BALANCE_CHAR_RE.sub(replacement, body)
    return body


# REMOVED 2026-08-11: dedupe_enumeration_prefix, which stripped a shared
# leading name repeated across items of one enumeration so a cash list named
# its bank once instead of on every line.
#
# IT DELETED REAL TEXT, on six accounts of a single real run. Two design
# errors, either of which is fatal on financial prose:
#
#   * "," was in the separator set, and a thousands separator is a comma. It
#     split "1,034.3万元" into "1" and "034.3万元", after which the second half
#     matched an earlier occurrence and was stripped -- the deck shipped
#     "余额为1,。我方核对了".
#   * "the longest prefix of this item that appears ANYWHERE earlier" is far
#     too loose. It ate "土地使用税" out of "房产税及土地使用税分别减少" and
#     "Loan-" out of "Loan-2 phase", because both had appeared earlier in the
#     same paragraph for unrelated reasons.
#
# The repeated bank name it was written for is a cosmetic complaint. Silent
# deletion inside a financial deliverable is not a fair trade, and no narrowing
# of the rule makes "delete text the model wrote" safe by default. If this is
# attempted again it belongs in the PROMPT, where the model decides what to
# omit, not in a post-processor that cuts finished sentences.


def create_result_shell(mapping_keys: List[str], dfs: Dict[str, pd.DataFrame]) -> Dict[str, Dict[str, str]]:
    return {key: {} for key in mapping_keys if key in dfs}


def _get_agent_stage_context(agent_name: str) -> Tuple[int, str, Optional[str]]:
    for agent_num, (name, label) in enumerate(SUBAGENT_SEQUENCE, start=1):
        if name == agent_name:
            previous_agent = None if agent_num == 1 else SUBAGENT_SEQUENCE[agent_num - 2][0]
            return agent_num, label, previous_agent
    raise ValueError(f"Unknown agent stage: {agent_name}")


def _store_agent_result(
    results: Dict[str, Dict[str, str]],
    mapping_key: str,
    agent_name: str,
    content: str,
    metadata: Dict[str, Any],
    state: Optional[AccountState] = None,
) -> None:
    # The ONLY place a stage transition is written. Popped before anything is
    # filed, so `results` is byte-identical with and without state tracking --
    # in particular an otherwise-empty metadata stays empty and still writes no
    # "<agent>_metadata" record.
    phase_events = (metadata or {}).pop("phase_events", None) if isinstance(metadata, dict) else None

    results[mapping_key][agent_name] = content
    if agent_name == "subagent_4":
        results[mapping_key]["agent_4_validation"] = metadata
        results[mapping_key]["final"] = content
    elif agent_name == "subagent_2" and (metadata or {}).get("clause_reviews"):
        # The Auditor now carries deterministic grounding too (see
        # process_single_agent_item). Filed under the SAME key rather than a
        # new one so the deck's highlighting, _apply_deterministic_verification
        # and the retry gate all keep reading one record -- the Validator
        # overwrites it when it runs. "final" is deliberately NOT set here: the
        # Validator, when it runs, produces the later text.
        results[mapping_key]["agent_4_validation"] = metadata
    elif metadata:
        # Every OTHER agent's metadata used to be dropped on the floor right
        # here. subagent_1 is the only stage that can emit a deterministic
        # fallback bullet, and its used_fallback/fallback_reason died at this
        # line -- which is a large part of why five archived runs shipped a
        # deck whose entire commentary was fallback bullets with nothing in
        # the results saying so. Filed under its own key so it can never
        # collide with the single agent_4_validation record everything else
        # reads.
        results[mapping_key]["%s_metadata" % agent_name] = metadata

    if state is None:
        return

    state.mark_stage(agent_name)
    if agent_name == "subagent_1":
        state.transition(PHASE_DRAFTED, "generator")
    elif agent_name == "subagent_2" and (metadata or {}).get("clause_reviews"):
        # Conditional on the reviews actually being there: the Auditor reaches
        # this line with empty metadata when its LLM call failed and the worker
        # passed the previous stage's text through, and nothing was grounded
        # then. Same condition as the agent_4_validation write above.
        state.transition(PHASE_GROUNDED, "arithmetic")
    elif agent_name == "subagent_4":
        state.transition(PHASE_GROUNDED, "llm_validator")

    if (metadata or {}).get("used_fallback"):
        # Degraded, NOT failed: this account keeps going and produces real
        # output. See PHASE_FAILED.
        state.note_degraded("fallback_bullet")

    # Drained after the stage transition, so a placeholder that the Generator
    # produced reads PLANNED -> DRAFTED -> FAILED rather than jumping.
    for event in phase_events or []:
        if event.get("to"):
            state.transition(event["to"], event.get("cause", ""))
        if event.get("degraded"):
            state.note_degraded(event["degraded"])


def apply_house_style(content: str, language: str, statement_type: str = "") -> str:
    """House style enforced as code, not only as prompt text. Both rules exist
    in the prompts already; these are the same rules where they cannot be
    ignored. Neither touches an amount, so number-grounding is unaffected.

    Lifted out of _finalize_agent_content (which is its only production caller)
    so a repair can re-apply EXACTLY this and nothing else after splicing a
    patch. What deliberately stays outside it: parse_validator_response, which
    would re-parse a bullet as a validator JSON envelope, and clean_agent_output,
    which a patch applies itself to the model's own reply before the splice.

    Measured idempotent on real output: re-applying polish_english_commentary to
    24 English archived finals and humanise_zero_balance + humanise_report_
    language to 104 Chinese ones produced zero differences -- which is what lets
    a repair compare styled-once content against styled-twice content directly.
    """
    if language != "Eng":
        content = humanise_zero_balance(content, statement_type)
        content = humanise_report_language(content)
    if language == "Eng":
        content = polish_english_commentary(content)
    return content


def _finalize_agent_content(
    *,
    agent_name: str,
    raw_content: str,
    previous_output: str,
    language: str,
    statement_type: str = "",
) -> Tuple[str, Dict[str, Any]]:
    metadata: Dict[str, Any] = {}
    if agent_name == "subagent_4":
        parsed = parse_validator_response(raw_content, fallback_content=previous_output)
        content = clean_agent_output(parsed["final_content"])
        metadata = {
            "final_content": content,
            "raw_response": parsed.get("raw_response", raw_content),
            "clause_reviews": parsed.get("clause_reviews", []),
        }
    else:
        content = clean_agent_output(raw_content)
    content = apply_house_style(content, language, statement_type)
    if agent_name == "subagent_4" and metadata:
        metadata["final_content"] = content
    return content, metadata


def _notify_stage_progress(
    progress_callback,
    *,
    agent_num: int,
    agent_label: str,
    completed: int,
    total_eligible: int,
    total_items: int,
    mapping_key: str,
) -> None:
    if not progress_callback:
        return
    progress_callback(
        agent_num,
        agent_label,
        completed,
        total_eligible,
        ((agent_num - 1) * total_items) + completed,
        mapping_key,
    )


def _run_ai_call(ai_helper, user_prompt: str, system_prompt: str, agent_name: str, timeout: int = 30):
    result_container = {"response": None, "error": None, "completed": False}
    agent_cfg = ai_helper.get_agent_settings(agent_name)

    def call_ai():
        try:
            result_container["response"] = ai_helper.get_response(
                user_prompt,
                system_prompt,
                temperature=agent_cfg.get("temperature"),
                max_tokens=agent_cfg.get("max_tokens"),
                top_p=agent_cfg.get("top_p"),
                frequency_penalty=agent_cfg.get("frequency_penalty"),
                presence_penalty=agent_cfg.get("presence_penalty"),
                allow_thinking=agent_cfg.get("allow_thinking"),
                reasoning_effort=agent_cfg.get("reasoning_effort"),
            )
            result_container["completed"] = True
        except Exception as exc:  # pragma: no cover - defensive
            result_container["error"] = exc
            result_container["completed"] = True

    thread = threading.Thread(target=call_ai, daemon=True)
    thread.start()
    thread.join(timeout=timeout)

    if not result_container["completed"]:
        raise TimeoutError(f"AI call timeout after {timeout} seconds")
    if result_container["error"]:
        raise result_container["error"]
    return result_container["response"]


def _agent_prompt_kwargs(
    agent_name: str,
    mapping_key: str,
    prompt_manager: PromptEngine,
    previous_output: str,
    agent_config: Optional[Dict[str, Any]] = None,
) -> Dict[str, str]:
    if agent_name == "subagent_1" and str(previous_output or "").strip():
        return {"previous_content": previous_output}
    if agent_name == "subagent_2":
        return {
            "account": prompt_manager.get_mapping_component(mapping_key, component="type") or mapping_key,
            "output": previous_output,
        }
    if agent_name == "subagent_3":
        cfg = agent_config or {}
        default_target = int(cfg.get("reduction_target_pct", 64))
        statement_type = str(prompt_manager.get_mapping_component(mapping_key, component="type") or "").strip().upper()
        if statement_type == "BS" and cfg.get("reduction_target_pct_bs") is not None:
            reduction_target = int(cfg["reduction_target_pct_bs"])
        else:
            reduction_target = default_target
        return {
            "previous_content": previous_output,
            "original_length": len(previous_output or ""),
            "reduction_target_pct": str(reduction_target),
        }
    if agent_name == "subagent_4":
        cfg = agent_config or {}
        return {
            "content": previous_output,
            "materiality_threshold_pct": str(cfg.get("materiality_threshold_pct", 5)),
        }
    return {}


_REVENUE_KEY_NEEDLES = ("operating income", "revenue", "营业收入")
_NOT_REVENUE_KEY_NEEDLES = ("non-operating", "营业外", "cost", "成本")


def _build_peer_context(dfs: Optional[Dict[str, pd.DataFrame]]) -> Optional[Dict[str, Any]]:
    """Revenue growth for the latest comparable period, so an expense
    account's prompt can point out that it grew out of proportion to
    revenue. Every account is otherwise generated in isolation from its own
    DataFrame, so no single account's commentary can currently make that
    observation at all -- which is one of the project team's asks.

    Computed once from whichever account looks like operating revenue, and
    only over FULL periods: a partial tail period against a full year would
    read as a collapse that is purely an artefact of period length.
    """
    if not dfs:
        return None
    for key, df in dfs.items():
        low = str(key).lower()
        if any(n in low for n in _NOT_REVENUE_KEY_NEEDLES):
            continue
        if not any(n in low for n in _REVENUE_KEY_NEEDLES):
            continue
        if not isinstance(df, pd.DataFrame) or df.empty:
            continue
        attrs = df.attrs or {}
        integrity = attrs.get("integrity") or {}
        cols = [
            c for c in list(df.columns)[1:]
            if str(c) != "__source_row_idx" and not str(c).endswith("_formatted")
        ]
        if len(cols) < 2:
            continue
        row_types = attrs.get("row_types_by_description") or {}
        desc_col = df.columns[0]
        total_idx = None
        for idx, row in df.iterrows():
            if str(row_types.get(str(row[desc_col]), "")).lower() in ("total", "subtotal"):
                total_idx = idx
        try:
            if total_idx is None:
                vals = [float(df[c].fillna(0).sum()) for c in cols]
            else:
                vals = [float(df.loc[total_idx, c] or 0) for c in cols]
        except Exception:
            continue
        months = attrs.get("annualization_months") or integrity.get("annualization_months")
        dropped_tail = False
        if isinstance(months, (int, float)) and 0 < months < 12 and len(vals) > 2:
            # cols is truncated WITH vals: the period the surviving figure was
            # measured over has to travel with it, or a ratio built downstream
            # would divide by the wrong column's revenue.
            vals, cols = vals[:-1], cols[:-1]
            dropped_tail = True
        if len(vals) < 2:
            continue
        scale = max((abs(v) for v in vals), default=0.0)
        prev, curr = vals[-2], vals[-1]
        if scale <= 0:
            continue
        # How long the surviving latest period covers. Dropping a partial tail
        # leaves a full period behind; keeping one (only two periods, so there
        # was nothing to fall back to) means the tail's own length stands.
        period_months = (
            float(months)
            if not dropped_tail and isinstance(months, (int, float)) and 0 < months < 12
            else 12.0
        )
        # Growth stays optional -- a base too small for a percentage to mean
        # anything used to abandon the whole peer context, taking the revenue
        # LEVEL with it. The level is what a ratio needs, and it is fine even
        # when growth is not, so the two are now reported independently.
        growth = (
            (curr - prev) / abs(prev) * 100 if abs(prev) >= scale * 0.01 else None
        )
        return {
            "revenue_growth_pct": growth,
            "revenue_key": key,
            "revenue_latest": curr,
            "revenue_period": str(cols[-1]),
            "revenue_months": period_months,
            # The prior period too, so a ratio can be stated as a MOVEMENT.
            # "税金及附加相当于收入的148%" is a number; "占收入比由14%升至148%"
            # is the finding, and it is the one that can actually be attributed.
            "revenue_prev": prev,
            "revenue_prev_period": str(cols[-2]),
        }
    return None


def process_single_agent_item(
    agent_name: str,
    mapping_key: str,
    df: Optional[pd.DataFrame],
    ai_helper,
    prompt_manager: PromptEngine,
    logger: PipelineRunLogger,
    previous_output: str = "",
    user_comment: str = "",
    dfs: Optional[Dict[str, pd.DataFrame]] = None,
    health: Optional[_RunHealth] = None,
    run_state: Optional[RunState] = None,
) -> Tuple[str, str, Dict[str, Any]]:
    """Run one account through a single agent stage.

    `run_state` is ADDITIVE — `dfs` stays. It was left alone through the
    state-machine milestone because three call sites omitted `dfs` entirely (the
    peer_context/siblings gap) and closing that gap changes prompts, which that
    milestone's verdict-identical gate did not allow. All three now pass it: the
    two feedback-retry calls, and the Reprompt path's Generator, which was the
    last one left and is fixed alongside the repair path.

    Nothing here writes a transition — see _queue_phase_event for why.
    """
    # A caller that wants no tally still gets one; it is simply discarded.
    # Keeps the eight recording sites below free of None checks.
    health = health or _RunHealth()
    # "the LLM was never even asked" is a different degradation from "the LLM
    # was asked and failed", and only the breaker path can tell them apart.
    _breaker_skipped = False
    try:
        logger.log_agent_start(agent_name, mapping_key)
        agent_cfg = ai_helper.get_agent_settings(agent_name)

        system_prompt, user_prompt = prompt_manager.render_prompt(
            agent_name=agent_name,
            language=ai_helper.language,
            mapping_key=mapping_key,
            df=df,
            data_format=ai_helper.data_format,
            user_comment=user_comment,
            peer_context=_build_peer_context(dfs),
            cross_account_facts=(run_state.facts if run_state is not None else None),
            analysis_thresholds=(getattr(ai_helper, "full_config", None) or {}).get("analysis"),
            **_agent_prompt_kwargs(agent_name, mapping_key, prompt_manager, previous_output, agent_config=agent_cfg),
        )

        if logger.debug_mode:
            logger.log_debug("PROMPT_SYSTEM", mapping_key, "Agent=%s len=%s" % (agent_name, len(system_prompt)), system_prompt)
            logger.log_debug("PROMPT_USER", mapping_key, "Agent=%s len=%s" % (agent_name, len(user_prompt)), user_prompt)

        if agent_name == "subagent_1" and (not system_prompt or not user_prompt):
            placeholder = f"Content generation skipped for {mapping_key}: No prompts available"
            health.record_no_prompt(mapping_key)
            return mapping_key, placeholder, _queue_phase_event({}, degraded="no_prompt")

        if not system_prompt or not user_prompt:
            health.record_no_prompt(mapping_key)
            return mapping_key, previous_output, _queue_phase_event({}, degraded="no_prompt")

        # Auto-reprompt on timeout. The user does not want to see "AI call
        # timed out" placeholder text in the commentary; retry up to twice
        # before falling back. Retries use the SAME prompt — if the API is
        # truly unresponsive the timeout will fire again, but typically a
        # transient slow response succeeds on retry.
        # Circuit breaker: when the API is stressed, repeatedly retrying just
        # wastes wall-clock time. After N consecutive failures across the
        # current stage, skip remaining LLM calls in this stage and fall
        # through to the deterministic fallback. Reset between stages so each
        # stage gets a fresh chance.
        if _PIPELINE_BREAKER.is_open(agent_name):
            health.record_breaker_skip()
            logger.logger.warning(
                "[%s] %s: circuit breaker OPEN — skipping LLM call, using fallback",
                agent_name, mapping_key,
            )
            _breaker_skipped = True
            raise RuntimeError("Circuit breaker open for this stage")

        response = None
        last_exc: Optional[Exception] = None
        # Exponential backoff between retries: 0s, 1s, 2s — gives the API a
        # brief moment to recover without compounding total wall-time too much.
        # Earlier (0, 2, 5) added 7s per failed call which compounded to
        # several extra minutes on a stressed run; (0, 1, 2) saves ~4s/call.
        retry_backoffs = [0.0, 1.0, 2.0]
        for attempt in range(1, 4):  # 1 initial + 2 retries
            if attempt > 1 and retry_backoffs[attempt - 1] > 0:
                time.sleep(retry_backoffs[attempt - 1])
            try:
                # Local models (Qwen3-32B etc.) are slow and emit <think> tokens, so the
                # default must give the SAME headroom to every stage. Previously only the
                # Validator set call_timeout=90 in config; Generator/Auditor fell to 30s and
                # timed out on large prompts. Derive a model-aware default so all stages match.
                _default_timeout = 90 if getattr(ai_helper, "model_type", "") == "local" else 30
                call_timeout = int(agent_cfg.get("call_timeout", _default_timeout))
                response = _run_ai_call(ai_helper, user_prompt, system_prompt, agent_name, timeout=call_timeout)
                last_exc = None
                _PIPELINE_BREAKER.record_success(agent_name)
                health.record_success(agent_name)
                if attempt > 1:
                    logger.logger.info(
                        "[%s] %s: succeeded on retry %s/2",
                        agent_name, mapping_key, attempt - 1,
                    )
                break
            except (TimeoutError, Exception) as exc:
                last_exc = exc
                health.record_failed_attempt(exc)
                logger.logger.warning(
                    "[%s] %s: AI call attempt %s failed (%s); %s",
                    agent_name, mapping_key, attempt, str(exc)[:80],
                    "retrying" if attempt < 3 else "giving up",
                )
        if response is None:
            _PIPELINE_BREAKER.record_failure(agent_name)
            raise last_exc if last_exc is not None else RuntimeError("AI call failed with no exception captured")

        raw_content = response["content"].strip().replace("\n\n", "\n").replace("\n \n", "\n")

        if logger.debug_mode:
            logger.log_debug("RAW_OUTPUT", mapping_key, "Agent=%s len=%s" % (agent_name, len(raw_content)), raw_content)

        statement_type = ""
        try:
            statement_type = prompt_manager.get_mapping_component(mapping_key, component="type") or ""
        except Exception:
            pass

        content, metadata = _finalize_agent_content(
            agent_name=agent_name,
            raw_content=raw_content,
            previous_output=previous_output,
            language=ai_helper.language,
            statement_type=statement_type,
        )

        # Deterministic hallucination/reasoning verification: layer Python
        # number-grounding over the LLM's soft judgement. This catches
        # fabricated figures the weak local model misses AND drops its false
        # positives on figures that actually match the source. Wrapped so a
        # verifier error never breaks the pipeline (keeps the LLM clause_reviews).
        #
        # Runs after the AUDITOR as well as the Validator. Attaching it only to
        # the Validator was an accident of implementation, not a design
        # decision: it is pure arithmetic over the account's own data, costs no
        # tokens, and does not care which agent last touched the text. Attached
        # here too, every account carries clause_reviews after stage 2 -- so the
        # deck's highlighting and the retry gate both have something to read on
        # accounts the Validator will skip, and an Auditor-stage defect can
        # actually trigger a retry. The Validator overwrites this when it runs
        # (its text is later, and it merges its own judgement in).
        if agent_name in ("subagent_2", "subagent_4") and df is not None:
            try:
                sibling_dfs = None
                if dfs and statement_type:
                    sibling_dfs = [
                        other_df for other_key, other_df in dfs.items()
                        if other_key != mapping_key
                        and prompt_manager.get_mapping_component(other_key, component="type") == statement_type
                    ]
                reviews = verify_commentary(
                    content, df, metadata.get("clause_reviews"),
                    sibling_dfs=sibling_dfs,
                )
                if agent_name == "subagent_2":
                    # The Auditor has no metadata of its own; give it the same
                    # shape the Validator produces so _store_agent_result can
                    # file it under the one validation record everything reads.
                    metadata = dict(metadata or {})
                    metadata["final_content"] = content
                metadata["clause_reviews"] = reviews
            except Exception as exc:  # pragma: no cover - defensive
                logger.logger.warning("[verify_commentary] %s: %s", mapping_key, exc)

        if logger.debug_mode and agent_name == "subagent_4" and metadata.get("clause_reviews"):
            reviews = metadata["clause_reviews"]
            supported_count = sum(1 for r in reviews if r.get("supported"))
            unsupported = [r for r in reviews if not r.get("supported")]
            logger.log_debug("VALIDATION", mapping_key,
                "Clauses: %s total, %s supported, %s unsupported" % (len(reviews), supported_count, len(unsupported)))
            for r in unsupported:
                logger.log_debug("UNSUPPORTED_CLAUSE", mapping_key,
                    '"%s" -- %s' % (str(r.get("clause", ""))[:80], str(r.get("reason", ""))[:120]))

        logger.log_agent_complete(
            agent_name,
            mapping_key,
            response,
            system_prompt,
            user_prompt,
            prompt_context=prompt_manager.build_prompt_context_snapshot(
                df,
                language=ai_helper.language,
                user_comment=user_comment,
                previous_output=previous_output,
            ),
        )
        return mapping_key, content, metadata
    except Exception as exc:
        logger.log_error(agent_name, mapping_key, exc)
        # Queued, not recorded: this still runs on a worker thread.
        events: Dict[str, Any] = {}
        if _breaker_skipped:
            events = _queue_phase_event(events, degraded="circuit_breaker")
        if agent_name == "subagent_1":
            fallback = _build_deterministic_fallback_bullet(mapping_key, df, ai_helper.language)
            if fallback:
                logger.logger.warning(
                    "[%s] %s: AI unavailable after retries; using deterministic data-only fallback",
                    agent_name, mapping_key,
                )
                health.record_fallback(mapping_key, str(exc)[:120])
                events.update({"used_fallback": True, "fallback_reason": str(exc)[:120]})
                return mapping_key, fallback, events
            health.record_error_text(mapping_key)
            return (
                mapping_key,
                f"Content generation failed for {mapping_key}: {str(exc)[:100]}",
                _queue_phase_event(events, to=PHASE_FAILED, cause="error_placeholder_text"),
            )
        if previous_output and str(previous_output).strip():
            health.record_passthrough(mapping_key)
            return mapping_key, previous_output, _queue_phase_event(
                events, degraded="stage_passthrough_after_error",
            )
        health.record_error_text(mapping_key)
        return (
            mapping_key,
            f"Content generation incomplete for {mapping_key}: {str(exc)[:100]}",
            _queue_phase_event(events, to=PHASE_FAILED, cause="error_placeholder_text"),
        )


def _build_deterministic_fallback_bullet(
    mapping_key: str,
    df: Optional[pd.DataFrame],
    language: str,
) -> str:
    """Produce a minimal data-only bullet when the LLM is unreachable.

    The bullet uses the latest period's total from the dataframe (last numeric
    column header is treated as the latest reporting date). Output format
    mirrors the project's reference style; downstream stages may still polish.
    """
    if df is None or df.empty:
        return ""
    try:
        from ..financial_display_format import format_number_chinese
    except Exception:
        return ""

    numeric_cols = [c for c in df.columns[1:] if not str(c).endswith("_formatted")]
    if not numeric_cols:
        return ""
    latest_col = numeric_cols[-1]
    try:
        if not pd.api.types.is_numeric_dtype(df[latest_col]):
            return ""
        total = float(df[latest_col].fillna(0).sum())
    except Exception:
        return ""
    if total == 0:
        return ""

    date_label = str(latest_col).strip()
    iso_match = re.match(r"(\d{4})-(\d{2})-(\d{2})", date_label)
    if iso_match:
        if language == "Chi":
            date_label = f"{iso_match.group(1)}年{int(iso_match.group(2))}月{int(iso_match.group(3))}日"
        else:
            date_label = _iso_to_long_date(iso_match)

    formatted = format_number_chinese(total, language)
    if language == "Chi":
        return f"截至{date_label}余额合计{formatted}。（自动摘要 - AI 暂时不可用）"
    return (
        f"the balance as at {date_label} totalled {formatted}. "
        "(Auto-summary while AI service was unavailable; please refresh this account when service recovers.)"
    )


def _resolve_max_workers(ai_helper, max_workers: Optional[int]) -> int:
    """Single source of truth for worker-count defaults across every thread
    pool in the pipeline (main stages, feedback loop, ensure-validation
    re-runs). An explicit caller-supplied value always wins; otherwise the
    provider's own <provider>.max_workers in config.yml (e.g. workbench: 4,
    validated against that gateway — see test_workbench_concurrency.py)
    wins; otherwise "local" (the self-hosted model) defaults to 1 — this
    project's local server serves one request effectively serially, so
    concurrent requests just queue rather than reduce wall time — and every
    other provider defaults to 4 (parallel by default, no UI toggle needed).
    Both defaults are overridable per-provider via config.yml."""
    if max_workers is not None:
        return max_workers
    _configured = (getattr(ai_helper, "config_details", None) or {}).get("max_workers")
    if _configured:
        return int(_configured)
    _model_type = getattr(ai_helper, "model_type", "")
    return 1 if _model_type == "local" else 4


def run_agent_stage(
    agent_name: str,
    mapping_keys: List[str],
    dfs: Dict[str, pd.DataFrame],
    results: Dict[str, Dict[str, str]],
    ai_helper,
    prompt_manager: PromptEngine,
    logger: PipelineRunLogger,
    use_multithreading: bool = True,
    max_workers: Optional[int] = None,
    progress_callback=None,
    total_items: int = 0,
    user_comments: Optional[Dict[str, str]] = None,
    health: Optional[_RunHealth] = None,
    run_state: Optional[RunState] = None,
):
    """Run all items for a single agent stage."""
    max_workers = _resolve_max_workers(ai_helper, max_workers)

    # Reset the circuit breaker at the start of each stage so a fresh
    # opportunity is given even if a prior stage tripped it.
    _PIPELINE_BREAKER.reset_stage(agent_name)

    agent_num, agent_label, previous_agent = _get_agent_stage_context(agent_name)

    # Eligibility reads stages_done, not phase. stages_done is a direct
    # restatement of the membership probe below -- "has this account been
    # through the previous stage?" -- so equality with the old behaviour is
    # provable rather than argued, and the old probe is kept beside it to prove
    # it on every real run rather than once in a harness. It is a dict lookup
    # per account per stage; the cost is not measurable.
    legacy_eligible_keys = []
    eligible_keys = []
    for key in mapping_keys:
        if key not in dfs or key not in results:
            continue
        if not previous_agent or previous_agent in results[key]:
            legacy_eligible_keys.append(key)
        state = run_state.account(key) if run_state is not None else None
        if state is None:
            continue
        if previous_agent and previous_agent not in state.stages_done:
            continue
        eligible_keys.append(key)

    if run_state is None:
        eligible_keys = legacy_eligible_keys
    else:
        _ELIGIBILITY_PARITY["checks"] += 1
        if eligible_keys != legacy_eligible_keys:
            _ELIGIBILITY_PARITY["mismatches"] += 1
            logger.logger.error(
                "[%s] stages_done eligibility disagrees with the membership test it "
                "replaced: %s vs %s. Falling back to the membership test",
                agent_name, eligible_keys, legacy_eligible_keys,
            )
            eligible_keys = legacy_eligible_keys

    if use_multithreading and len(eligible_keys) > 1:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {}
            for key in eligible_keys:
                future = executor.submit(
                    process_single_agent_item,
                    agent_name,
                    key,
                    dfs.get(key),
                    ai_helper,
                    prompt_manager,
                    logger,
                    results[key].get(previous_agent, "") if previous_agent else "",
                    (user_comments or {}).get(key, ""),
                    dfs,
                    health,
                    run_state,
                )
                futures[future] = key

            completed = 0
            for future in as_completed(futures):
                mapping_key, content, metadata = future.result()
                _store_agent_result(
                    results, mapping_key, agent_name, content, metadata,
                    state=run_state.account(mapping_key) if run_state is not None else None,
                )
                completed += 1
                _notify_stage_progress(
                    progress_callback,
                    agent_num=agent_num,
                    agent_label=agent_label,
                    completed=completed,
                    total_eligible=len(eligible_keys),
                    total_items=total_items,
                    mapping_key=mapping_key,
                )
    else:
        completed = 0
        for key in eligible_keys:
            mapping_key, content, metadata = process_single_agent_item(
                agent_name,
                key,
                dfs.get(key),
                ai_helper,
                prompt_manager,
                logger,
                results[key].get(previous_agent, "") if previous_agent else "",
                (user_comments or {}).get(key, ""),
                dfs,
                health,
                run_state,
            )
            _store_agent_result(
                results, mapping_key, agent_name, content, metadata,
                state=run_state.account(mapping_key) if run_state is not None else None,
            )
            completed += 1
            _notify_stage_progress(
                progress_callback,
                agent_num=agent_num,
                agent_label=agent_label,
                completed=completed,
                total_eligible=len(eligible_keys),
                total_items=total_items,
                mapping_key=mapping_key,
            )


def set_final_fallbacks(results: Dict[str, Dict[str, str]]):
    """Populate `final` from the latest successful agent output."""
    for key in results:
        results[key]["final"] = get_pipeline_result_text(results[key])


def settle_subtable_selection(
    mapping_keys: List[str], dfs: Dict[str, pd.DataFrame],
) -> List[Tuple[str, str]]:
    """Decide which accounts get a subtable BEFORE a single prompt is built,
    and strip the table off the frames that lose.

    The selection used to run at export, long after the commentary was
    written. So the model was told a table existed for EVERY account that
    could have one, wrote a lead-in ending "明细如下：" and left the components
    to the table -- and then the table was taken away. A real deck shipped
    "管理费用主要包括管理费用-行政及管理层调整。明细如下：" followed by prose,
    a promise pointing at nothing, and 固定资产 enumerated nine components in
    the sentence beside a table already listing thirteen.

    Deciding here fixes both at the source: an account with no table is never
    told it has one, so it writes ordinary prose that names its own
    components, and an account WITH a table still gets the guidance not to
    repeat it. This is also the honest answer to "can the model judge whether
    a table is warranted" -- it can only judge that if it is told, and it
    cannot be told by a decision taken after it has finished writing.

    Returns the (account, reason) pairs that lost, for the caller to log.
    """
    # Local import: the AI package does not otherwise depend on the PPTX one,
    # and the rule must not be duplicated -- the deck is capped by exactly the
    # same function that caps it at export (which then finds nothing left to
    # do, since the losers no longer carry a table).
    from ..pptx.helpers import _select_deck_subtables
    from ..pptx.payloads import _load_pptx_settings

    by_statement: Dict[str, List[Dict[str, Any]]] = {}
    stubs: List[Tuple[str, Dict[str, Any]]] = []
    for key in mapping_keys:
        df = dfs.get(key)
        if df is None or not hasattr(df, "attrs"):
            continue
        table = (df.attrs or {}).get("presentation_detail_table")
        if not table or not table.get("rows"):
            continue
        statement = str(((df.attrs.get("integrity") or {}).get("statement_type") or "")).strip().upper()
        stub = {"mapping_key": key, "account_name": key, "financial_data": df}
        by_statement.setdefault(statement or "?", []).append(stub)
        stubs.append((key, stub))

    if not stubs:
        return []

    try:
        settings = _load_pptx_settings()
    except Exception:
        settings = {}
    _select_deck_subtables(
        [(s, by_statement.get(s) or []) for s in ("IS", "BS", "?")], settings,
    )

    dropped: List[Tuple[str, str]] = []
    for key, stub in stubs:
        if stub.get("_subtable_ok"):
            continue
        dropped.append((key, "not selected for a subtable"))
        try:
            # Removed, not flagged: every downstream reader -- the prompt
            # builder, _presentation_table_for_account, the packer -- already
            # treats "no table" correctly, and a second flag they would all
            # have to honour is one more thing to get out of step.
            dfs[key].attrs.pop("presentation_detail_table", None)
        except Exception:
            pass
    return dropped


def run_ai_pipeline_with_progress(
    mapping_keys: List[str],
    dfs: Dict[str, pd.DataFrame],
    model_type: str = "deepseek",
    language: str = "Eng",
    use_heuristic: bool = False,
    use_multithreading: bool = True,
    max_workers: Optional[int] = None,
    progress_callback: Optional[Callable[..., None]] = None,
    user_comments: Optional[Dict[str, str]] = None,
    model_name: Optional[str] = None,
) -> Dict[str, Dict[str, str]]:
    """Run the 4-agent FDD pipeline with optional progress callbacks."""
    # Normalise UI language codes ("Chn" → "Chi") to match prompt-file keys.
    language = normalize_language_code(language)
    settle_subtable_selection(mapping_keys, dfs)
    total_items = len([key for key in mapping_keys if key in dfs])
    fdd_config = FDDConfig(language=language, model_type=model_type)
    debug_mode = fdd_config.get_debug_mode()
    logger = PipelineRunLogger(debug_mode=debug_mode)
    prompt_manager = get_prompt_engine()
    ai_helper = AIClient(
        model_type=model_type,
        agent_name="content_pipeline",
        language=language,
        use_heuristic=use_heuristic,
        model_name=model_name,
    )
    results = create_result_shell(mapping_keys, dfs)
    # Built HERE and not at the top of the function: settle_subtable_selection
    # mutates df.attrs (it strips the detail table off the frames that lose a
    # slot), and the fact half has to be built from the settled frames or it
    # describes a workbook the prompts were never given.
    run_state = RunState(dfs, mapping_keys, run_folder=logger.run_folder)
    # Cross-account facts and the relationships that survived a numeric test.
    # Built ONCE here, off the settled frames, and read from RunState by every
    # prompt render -- the alternative is what peer_context still does, which is
    # to rebuild the same thing on every call and come out differently on the
    # paths that forget to pass dfs. Only verified links reach a prompt; a
    # hypothesis that failed its test goes to the internal insight summary and
    # nowhere near the deck.
    try:
        from .facts import build_cross_account_facts

        run_state.facts = build_cross_account_facts(
            dfs, type_lookup=lambda k: prompt_manager.get_mapping_component(k, component="type"),
        )
    except Exception as exc:  # a fact table is an enrichment, never a gate
        logger.logger.warning("[CrossAccountFacts] skipped: %s", exc)
        run_state.facts = {}
    health = _RunHealth()

    logger.logger.info(
        "Starting FDD pipeline with %s items | model=%s | language=%s | multithreading=%s",
        total_items,
        ai_helper.model_type,
        language,
        use_multithreading,
    )

    validator_mode = fdd_config.get_validator_mode()
    for agent_name, agent_label in SUBAGENT_SEQUENCE:
        stage_keys = mapping_keys
        if agent_name == "subagent_4" and validator_mode == "selective":
            # The Validator is the pipeline's most expensive stage (59% of a
            # measured 964s run) because it re-emits the whole bullet PLUS a
            # per-clause review array -- most of which verify_commentary then
            # overwrites with its own deterministic number-grounding (174
            # stored clauses vs the model's own 73 on a real run).
            #
            # The one thing it can do that determinism cannot is judge a
            # CAUSAL claim. So run it only where there is a causal claim to
            # judge, and let the deterministic verifier handle the rest.
            stage_keys = [
                k for k in mapping_keys
                if commentary_asserts_inference(get_pipeline_result_text(results.get(k) or {}))
            ]
            skipped = len(mapping_keys) - len(stage_keys)
            logger.logger.info(
                "Validator (selective): %s of %s account(s) assert a causal claim; "
                "%s handled by deterministic verification only",
                len(stage_keys), len(mapping_keys), skipped,
            )
            # A no-op transition, deliberately: "the Validator did not run here,
            # and it was a decision" is a different fact from "the Validator
            # never got to this account", and only a recorded cause separates
            # them. Main thread, before the stage fans out.
            _selected = set(stage_keys)
            for _key in mapping_keys:
                _state = run_state.account(_key)
                if _state is not None and _key not in _selected:
                    _state.transition(_state.phase, "validator_skipped: no_causal_claim")
        if not stage_keys:
            continue
        logger.logger.info("Running %s stage", agent_label)
        run_agent_stage(
            agent_name=agent_name,
            mapping_keys=stage_keys,
            dfs=dfs,
            results=results,
            ai_helper=ai_helper,
            prompt_manager=prompt_manager,
            logger=logger,
            use_multithreading=use_multithreading,
            max_workers=max_workers,
            progress_callback=progress_callback,
            total_items=total_items,
            user_comments=user_comments,
            health=health,
            run_state=run_state,
        )

    if validator_mode == "selective":
        # Every account the Validator skipped still needs clause_reviews, both
        # for the deck's hallucination highlighting and for the feedback
        # loop's own gate. This is the deterministic half of the old
        # Validator -- number-grounding only, no LLM call, no rewrite.
        _apply_deterministic_verification(
            results=results, dfs=dfs, prompt_manager=prompt_manager, logger=logger,
            run_state=run_state,
        )

    # --- Feedback loop: re-run generator+validator for accounts with too many unsupported clauses ---
    feedback_config = fdd_config.get_feedback_loop_config()
    if feedback_config.get("enabled"):
        # Reset breakers: a stage tripped during the main run would otherwise stay
        # OPEN and make the feedback loop fail-fast (skip) every account silently.
        for _stage, _ in SUBAGENT_SEQUENCE:
            _PIPELINE_BREAKER.reset_stage(_stage)
        logger.logger.info(
            "Starting feedback loop (max_retries=%s, threshold=%.2f)",
            feedback_config["max_retries"],
            feedback_config["unsupported_threshold"],
        )
        eligible_keys = [k for k in mapping_keys if k in results and k in dfs]
        if use_multithreading and len(eligible_keys) > 1:
            fb_workers = _resolve_max_workers(ai_helper, max_workers)
            with ThreadPoolExecutor(max_workers=fb_workers) as executor:
                futures = {
                    executor.submit(
                        _run_feedback_loop_for_key,
                        key=key,
                        dfs=dfs,
                        results=results,
                        ai_helper=ai_helper,
                        prompt_manager=prompt_manager,
                        logger=logger,
                        feedback_config=feedback_config,
                        user_comments=user_comments,
                        progress_callback=progress_callback,
                        health=health,
                        run_state=run_state,
                    ): key
                    for key in eligible_keys
                }
                for future in as_completed(futures):
                    key = futures[future]
                    try:
                        retries = future.result()
                    except Exception as exc:
                        logger.logger.warning("[FeedbackLoop] %s: failed: %s", key, exc)
                        continue
                    if retries > 0:
                        logger.logger.info("[FeedbackLoop] %s: completed with %s retry(ies)", key, retries)
        else:
            for key in eligible_keys:
                retries = _run_feedback_loop_for_key(
                    key=key,
                    dfs=dfs,
                    results=results,
                    ai_helper=ai_helper,
                    prompt_manager=prompt_manager,
                    logger=logger,
                    feedback_config=feedback_config,
                    user_comments=user_comments,
                    progress_callback=progress_callback,
                    health=health,
                    run_state=run_state,
                )
                if retries > 0:
                    logger.logger.info("[FeedbackLoop] %s: completed with %s retry(ies)", key, retries)

    set_final_fallbacks(results)

    # Ensure every account with a final commentary has hallucination/reasoning
    # clause_reviews so the UI can highlight them. If the Validator stage
    # didn't produce clause_reviews (timeout, parse failure, etc.), run a
    # one-shot validator pass on the final text. Runs only for accounts that
    # need it, in parallel.
    #
    # It used to call _PIPELINE_BREAKER.reset_stage("subagent_4") first, so a
    # Validator stage that had tripped the breaker would not make this pass
    # fail-fast for every account. That reversed the breaker's whole purpose on
    # exactly the runs it existed for: measured over 255 archived runs this pass
    # fired 6 times, five of them after the Validator had already died on a hard
    # 400, and it then re-sent every account to an endpoint that had just failed
    # 84 times. The reset is gone; the health gates now live inside
    # _ensure_clause_reviews_on_final (M8.2).
    _ensure_clause_reviews_on_final(
        results=results,
        dfs=dfs,
        ai_helper=ai_helper,
        prompt_manager=prompt_manager,
        logger=logger,
        use_multithreading=use_multithreading,
        max_workers=max_workers,
        user_comments=user_comments,
        health=health,
        run_state=run_state,
    )

    _sweep_final_phases(results, run_state, logger)

    # Attached LAST, after set_final_fallbacks and the ensure pass have finished
    # walking `results` -- both iterate every key and would otherwise write a
    # "final" onto the tally. Same for the run sentinel below it.
    health_summary = health.as_dict()
    _log_run_health(logger, health_summary, total_items)
    results[RUN_HEALTH_KEY] = health_summary
    results[RUN_STATE_KEY] = {
        "run_folder": logger.run_folder,
        "audit_log": _write_run_audit_log(logger, run_state, results),
        "state": run_state.as_dict(),
        "defect_codes": _defect_code_frequency(run_state),
        "eligibility_parity": dict(_ELIGIBILITY_PARITY),
    }

    logger.finalize(results)
    return results


def _sweep_final_phases(
    results: Dict[str, Dict[str, str]],
    run_state: RunState,
    logger: PipelineRunLogger,
) -> None:
    """Give every account a terminal phase, on every run, gated on nothing.

    ACCEPTED has no site anywhere else in the pipeline, and the nearest
    candidate -- _evaluate_feedback_needed -- is reachable only inside
    `if feedback_config.get("enabled")`, a per-machine config value. An outcome
    tally that silently empties when one machine's config.yml differs from
    another's is not a tally, so this sweep is not gated on anything.

    Two phases are already decided and are NOT swept:

      ARBITRATED  the arbiter gave it a terminal phase and it still carries
                  defects, so the DEFECTIVE branch below would overwrite an
                  informative label with a coarser one.
      FAILED      MEASURED on a replay of an archived dead run (every LLM call a
                  hard 400): 8 accounts took the error-placeholder branch, the
                  deterministic sweep then attached clause_reviews to the
                  placeholder string, it contained no defective clause, and all
                  26 accounts came out ACCEPTED — on a deck that was entirely
                  filler. Grounding a placeholder is not acceptance.
    """
    for key, state in run_state.accounts.items():
        result = results.get(key) or {}
        reviews = ((result.get("agent_4_validation") or {}).get("clause_reviews") or []
                   if isinstance(result, dict) else [])
        # Harvested for EVERY account, before the skip below. An arbitrated
        # account is by definition the one that still carries defects, so it is
        # the one a reviewer opens the audit log for; leaving its defects out
        # because its phase was already decided would empty the per-run defect
        # table of exactly the rows that matter.
        state.defects = _clause_defect_records(reviews)
        if state.phase in (PHASE_ARBITRATED, PHASE_FAILED):
            continue
        if not reviews:
            # Nothing was ever grounded here. That is the same outcome as the
            # placeholder branches, so it gets the same label.
            state.transition(PHASE_FAILED, "final_sweep: no_clause_reviews")
            continue
        defective = count_defective_clauses(reviews)
        if defective:
            state.transition(
                PHASE_DEFECTIVE, "final_sweep: %s defective clause(s) unresolved" % len(defective),
            )
        else:
            state.transition(PHASE_ACCEPTED, "final_sweep: no defective clause")

    tally = run_state.phase_tally()
    stranded = [k for k, s in run_state.accounts.items() if s.phase not in TERMINAL_PHASES]
    logger.logger.info(
        "Run outcome: %s | %s account(s) with a non-terminal phase: %s",
        " ".join("%s=%s" % (phase, count) for phase, count in sorted(tally.items())),
        len(stranded), stranded or "none",
    )


def _clause_defect_records(reviews: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """The unsupported clauses, in the shape the audit log files them.

    `code`/`span`/`source_ref` are OPTIONAL and absent today -- the typed defect
    codes are a later step in this milestone, landing in validator.py. Reading
    them through .get() now means the audit line does not change shape when they
    arrive.
    """
    records = []
    for review in reviews or []:
        if not isinstance(review, dict) or review.get("supported"):
            continue
        records.append({
            "code": review.get("code"),
            "span": review.get("span"),
            "category": review.get("category"),
            "reason": str(review.get("reason") or "")[:400],
            "source_ref": review.get("source_ref"),
            "clause": str(review.get("clause") or "")[:200],
        })
    return records


def _defect_code_frequency(run_state: RunState) -> Dict[str, int]:
    """Per-run defect-code histogram. `null` until validator.py emits codes;
    the category is carried alongside so the table says something in the
    meantime."""
    counter: Counter = Counter()
    for state in run_state.accounts.values():
        for defect in state.defects:
            counter["%s/%s" % (defect.get("code"), defect.get("category"))] += 1
    return dict(counter)


def _account_audit_line(
    run_id: str,
    key: str,
    state: AccountState,
    result: Dict[str, Any],
) -> Dict[str, Any]:
    """One account's state path, self-contained enough to answer "why did this
    account end up like this?" read top to bottom."""
    validation = (result.get("agent_4_validation") or {}) if isinstance(result, dict) else {}
    reviews = validation.get("clause_reviews") or []
    # WHICH TEXT THE SPANS INDEX. Clause and amount offsets are defined against
    # agent_4_validation["final_content"] -- the stored string, which for some
    # accounts is NOT results["final"]. Naming and hashing it here is what lets
    # a later reader tell whether a stored span still means anything.
    spans_text = validation.get("final_content")
    spans_text = spans_text if isinstance(spans_text, str) else ""
    final_text = str((result.get("final") or "") if isinstance(result, dict) else "")
    line = dict(state.as_dict())
    line.update({
        "run_id": run_id,
        "account": key,
        "feedback_retries": (result.get("feedback_retries") if isinstance(result, dict) else None),
        "feedback_arbiter": (result.get("feedback_arbiter") if isinstance(result, dict) else None),
        "evidence": {
            "clauses_reviewed": len(reviews),
            "clauses_unsupported": sum(
                1 for r in reviews if isinstance(r, dict) and not r.get("supported")
            ),
            "clauses_defective": len(count_defective_clauses(reviews)),
        },
        "spans_index": {
            "text": "agent_4_validation.final_content",
            "sha1": hashlib.sha1(spans_text.encode("utf-8")).hexdigest() if spans_text else None,
            "length": len(spans_text),
            "same_as_final": bool(spans_text) and spans_text == final_text,
        },
    })
    return line


def _write_run_audit_log(
    logger: PipelineRunLogger,
    run_state: RunState,
    results: Dict[str, Dict[str, str]],
) -> str:
    lines = [
        _account_audit_line(logger.run_id, key, state, results.get(key) or {})
        for key, state in sorted(run_state.accounts.items())
    ]
    try:
        path = logger.write_audit_log(lines)
    except Exception as exc:  # pragma: no cover - a log is never worth losing a run over
        logger.logger.warning("[Audit] could not write audit.jsonl: %s", exc)
        return ""
    logger.logger.info(
        "[Audit] %s account path(s) -> %s | %s transition(s) | causes: %s",
        len(lines), path,
        sum(len(s.transitions) for s in run_state.accounts.values()),
        run_state.transition_causes(),
    )
    return path


def _ensure_clause_reviews_on_final(
    *,
    results: Dict[str, Dict[str, str]],
    dfs: Dict[str, pd.DataFrame],
    ai_helper,
    prompt_manager: PromptEngine,
    logger: PipelineRunLogger,
    use_multithreading: bool,
    max_workers: Optional[int],
    user_comments: Optional[Dict[str, str]] = None,
    health: Optional[_RunHealth] = None,
    run_state: Optional[RunState] = None,
) -> None:
    """Re-run Validator on accounts whose final commentary lacks clause_reviews."""
    needs_validation: List[str] = []
    for key, result in results.items():
        if not isinstance(result, dict):
            continue
        final_text = str(result.get("final") or "").strip()
        if not final_text:
            continue
        validation = result.get("agent_4_validation") or {}
        if isinstance(validation, dict) and validation.get("clause_reviews"):
            continue
        if key not in dfs:
            continue
        needs_validation.append(key)

    if not needs_validation:
        return

    # This pass was written for a Validator that failed TRANSIENTLY on a
    # handful of accounts. Both gates below exist because that is not what it
    # actually met: in five of the six archived runs where it fired, the
    # Validator had already died on a hard 400 for every single account, and
    # the pass simply paid for a fourth round of the same error.
    if health is not None and not health.stage_succeeded("subagent_4"):
        logger.logger.warning(
            "[EnsureValidation] skipped for %s account(s): the Validator made no "
            "successful call anywhere in this run, so a fourth pass can only repeat it",
            len(needs_validation),
        )
        return
    if _PIPELINE_BREAKER.is_open("subagent_4"):
        logger.logger.warning(
            "[EnsureValidation] skipped for %s account(s): the Validator circuit breaker "
            "is open. It is deliberately NOT reset here — resetting it is what turned a "
            "fail-fast into a full extra pass against a failing endpoint",
            len(needs_validation),
        )
        return

    corpus = sum(
        1 for r in results.values()
        if isinstance(r, dict) and str(r.get("final") or "").strip()
    )
    if corpus and len(needs_validation) / corpus > 0.20:
        logger.logger.warning(
            "[EnsureValidation] %s of %s account(s) (%.0f%%) finished with no "
            "clause_reviews. That is a SYSTEMIC Validator failure, not the few transient "
            "timeouts this pass was written for; it is running anyway, but the fourth "
            "pass is being paid for a problem that belongs upstream",
            len(needs_validation), corpus, len(needs_validation) / corpus * 100,
        )

    logger.logger.info(
        "[EnsureValidation] %s account(s) need fresh clause_reviews", len(needs_validation),
    )

    def _run_one(key: str):
        final_text = str((results.get(key) or {}).get("final") or "").strip()
        try:
            _k, content, metadata = process_single_agent_item(
                "subagent_4", key, dfs.get(key), ai_helper, prompt_manager, logger,
                previous_output=final_text,
                user_comment=(user_comments or {}).get(key, ""),
                dfs=dfs,
                health=health,
                run_state=run_state,
            )
            if isinstance(metadata, dict) and metadata.get("clause_reviews"):
                # Keep the original final text (don't overwrite); just attach
                # clause_reviews so highlighting works.
                metadata.pop("phase_events", None)  # this write bypasses _store_agent_result
                results[key]["agent_4_validation"] = metadata
                _state = run_state.account(key) if run_state is not None else None
                if _state is not None:
                    _state.mark_stage("subagent_4")
                    _record_recovered_grounding(_state, "ensure_validation_after_stage_failure")
                logger.logger.info(
                    "[EnsureValidation] %s: validated %s clause(s)", key,
                    len(metadata.get("clause_reviews", [])),
                )
        except Exception as exc:
            logger.logger.warning("[EnsureValidation] %s: failed: %s", key, exc)

    if use_multithreading and len(needs_validation) > 1:
        ev_workers = _resolve_max_workers(ai_helper, max_workers)
        with ThreadPoolExecutor(max_workers=ev_workers) as executor:
            list(executor.map(_run_one, needs_validation))
    else:
        for key in needs_validation:
            _run_one(key)


#: Clause categories that mean the model got a FACT wrong, as opposed to
#: drawing an inference the data doesn't spell out. Only these justify
#: spending another generation pass.
RETRIABLE_CLAUSE_CATEGORIES = ("hallucination",)

#: Wording that asserts a CAUSE, DRIVER or EXPECTATION rather than restating
#: a figure. These are the only clauses the LLM Validator can judge that
#: verify_commentary's deterministic number-grounding cannot: numbers either
#: tie to the data or they don't, but "主要系X所致" is a claim about WHY.
#: Measured on a real 23-account run, only these accounts produced any
#: reasoning flag at all, so gating the Validator on this marker set keeps
#: the whole capability while skipping the ~80% of accounts it had nothing
#: to say about.
#: NOTE "主要为" / "主要包括" are deliberately ABSENT. They introduce a
#: COMPOSITION ("主要为银行存款"), not a cause, and including them selected
#: 48% of a real 21-account set instead of 33% for no additional flag.
_INFERENCE_MARKERS_CHI = (
    "主要系", "主要由于", "所致", "导致", "预计", "推测", "预期",
    "反映", "结合其性质", "表明", "拉低", "带动", "归因", "驱动",
    "原因",
)
#: "受X影响" is a cause, but its two halves are separated by whatever X is, so
#: it cannot be a substring test. It and "系.*所致" sat in the tuple above as
#: literal strings and were tested with `in`, which means neither could ever
#: match anything — for as long as they had been there. "系.*所致" is simply
#: dropped: the bare "所致" already in the tuple catches every case it could.
#: This one is now a real search, bounded to a short span so it cannot bridge
#: a sentence break and pair an unrelated 受 with a later 影响.
#: NOTE this WIDENS Validator selection and therefore costs LLM calls that
#: were never being made.
_INFERENCE_PATTERNS_CHI = (re.compile(r"受[^。；;\n]{1,20}影响"),)
_INFERENCE_MARKERS_ENG = (
    "mainly due to", "driven by", "attributable to", "reflecting", "as a result of",
    "expected to", "indicates that", "suggests that", "because of", "owing to",
)


# Phrasings that say a cause is NOT available. They contain inference-looking
# words but make no claim, so they must not, on their own, buy an LLM call.
_NO_CAUSE_DISCLAIMERS = (
    "原因未在资料中说明", "原因未在数据中说明", "备注中未提供进一步解释",
    "未在资料中说明原因", "数据中未提供原因", "变动原因未在数据中明确说明",
    "the reason is not set out in the information provided",
    "no reason is given in the information provided",
)


def commentary_asserts_inference(text: str) -> bool:
    """Does this commentary make a causal/forward-looking claim?

    Cheap, deterministic pre-filter deciding whether an account is worth an
    LLM validation call at all. Deliberately over-inclusive: a false
    positive costs one avoidable call, a false negative loses a real
    reasoning flag from the deck's orange highlighting.
    """
    body = str(text or "")
    if not body.strip():
        return False
    # The prompts now REQUIRE the model to say so when the databook does not
    # explain a movement. That disclaimer is the opposite of an unsupported
    # causal claim -- it asserts nothing -- but it reads as one to the marker
    # scan, and adding the rule pushed the Validator from 173s to 560s on a
    # real run for a single extra flag. Strip the disclaimers before deciding.
    for _disclaimer in _NO_CAUSE_DISCLAIMERS:
        body = body.replace(_disclaimer, "")
    if not body.strip():
        return False
    lowered = body.lower()
    return (
        any(m in body for m in _INFERENCE_MARKERS_CHI)
        or any(p.search(body) for p in _INFERENCE_PATTERNS_CHI)
        or any(m in lowered for m in _INFERENCE_MARKERS_ENG)
    )


def count_defective_clauses(clause_reviews: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """The unsupported clauses worth re-generating for.

    NOT simply "every clause the validator marked unsupported". A real run
    flagged 3 clauses across 23 accounts and every one was
    category="reasoning" -- e.g. "预计主要系待抵扣进项税逐步抵扣所致". That is
    an FDD consultant drawing a supportable inference, which the deliverable
    explicitly wants and which the deck renders in orange as marked
    reasoning. Retrying on those would spend real tokens punishing good
    output and would push the model toward blander commentary.

    A "hallucination" is different in kind: a number, direction or fact that
    cannot be tied to the data at all. That is the only thing a re-run can
    actually fix.
    """
    return [
        r for r in (clause_reviews or [])
        if isinstance(r, dict)
        and not r.get("supported")
        and str(r.get("category", "")).strip().lower() in RETRIABLE_CLAUSE_CATEGORIES
    ]


def _apply_deterministic_verification(
    *,
    results: Dict[str, Dict[str, str]],
    dfs: Dict[str, pd.DataFrame],
    prompt_manager: PromptEngine,
    logger: PipelineRunLogger,
    run_state: Optional[RunState] = None,
) -> None:
    """Attach clause_reviews to every account the LLM Validator skipped.

    Pure `verify_commentary` -- the same deterministic number-grounding that
    is ALREADY authoritative over the Validator's own judgement on the
    accounts it does run (see process_single_agent_item's subagent_4 branch).
    Free, no LLM call, and it produces the identical clause_reviews shape the
    deck's highlighting and the feedback loop's gate both read.
    """
    done = 0
    for key, result in results.items():
        if not isinstance(result, dict):
            continue
        validation = result.get("agent_4_validation") or {}
        if isinstance(validation, dict) and validation.get("clause_reviews"):
            continue  # the LLM Validator already ran for this account
        content = get_pipeline_result_text(result)
        df = dfs.get(key)
        if not str(content or "").strip() or df is None:
            continue
        try:
            statement_type = prompt_manager.get_mapping_component(key, component="type")
            sibling_dfs = [
                other_df for other_key, other_df in dfs.items()
                if other_key != key
                and prompt_manager.get_mapping_component(other_key, component="type") == statement_type
            ] if statement_type else None
            reviews = verify_commentary(content, df, None, sibling_dfs=sibling_dfs)
        except Exception as exc:
            logger.logger.warning("[DeterministicVerify] %s: %s", key, exc)
            continue
        result["agent_4_validation"] = {"final_content": content, "clause_reviews": reviews}
        # Downstream (set_final_fallbacks, the feedback loop, the PPTX
        # payload builder) all read subagent_4 as "the validated text".
        # Nothing rewrote it here, so it is the Auditor's own output.
        result.setdefault("subagent_4", content)
        # EXPECT ZERO OF THESE ON A HEALTHY RUN. The loop above `continue`s on
        # any account that already has clause_reviews, and the Auditor attaches
        # them to every account it processes -- so reaching this line means a
        # stage failed for that account and the grounding was recovered here
        # instead. Recorded with its own cause so the count is assertable rather
        # than assumed; ad-hoc/workbench/replay_pipeline_state.py asserts it.
        _record_recovered_grounding(
            run_state.account(key) if run_state is not None else None,
            "grounding_recovered_after_stage_failure",
        )
        done += 1
    if done:
        logger.logger.info(
            "[DeterministicVerify] grounded %s account(s) with no LLM call", done,
        )


def _evaluate_feedback_needed(
    results: Dict[str, Dict[str, str]],
    key: str,
    unsupported_threshold: float,
) -> tuple[bool, float, List[Dict[str, Any]]]:
    """Whether this mapping_key's validator result warrants a retry.

    Gate is "any hallucination at all", not "a high enough RATIO of
    unsupported clauses". A ratio is the wrong shape of test here: one
    fabricated figure in a long, otherwise-correct bullet is a reportable
    error, but it dilutes to a ratio far below any sane threshold, so the
    old 0.30 gate could never fire on exactly the case it existed for.
    `unsupported_threshold` is still honoured as an ADDITIONAL trigger for
    output that is broadly unsupported rather than specifically wrong.
    """
    validation = (results.get(key) or {}).get("agent_4_validation") or {}
    clause_reviews = validation.get("clause_reviews") or []
    if not clause_reviews:
        return False, 0.0, []
    defective = count_defective_clauses(clause_reviews)
    unsupported = [r for r in clause_reviews if isinstance(r, dict) and not r.get("supported")]
    ratio = len(unsupported) / len(clause_reviews)
    needed = bool(defective) or ratio > unsupported_threshold
    return needed, ratio, (defective or unsupported)


def _sibling_dfs_for(
    key: str,
    dfs: Optional[Dict[str, pd.DataFrame]],
    prompt_manager: PromptEngine,
    statement_type: str,
) -> Optional[List[Any]]:
    """Every OTHER account's df of the same statement type — the sibling set the
    production verify_commentary call was given.

    Same construction as process_single_agent_item's own (:1137-1143). Repeated
    rather than shared because re-verifying a patch against a DIFFERENT sibling
    set is a different grounding pool, and that manufactures defects that were
    never in the text; M1 step 6 folds all three sites into RunState.siblings.
    """
    if not dfs or not statement_type:
        return None
    return [
        other_df for other_key, other_df in dfs.items()
        if other_key != key
        and prompt_manager.get_mapping_component(other_key, component="type") == statement_type
    ]


def _attempt_patch_first(
    *,
    key: str,
    dfs: Dict[str, pd.DataFrame],
    results: Dict[str, Dict[str, str]],
    ai_helper,
    prompt_manager: PromptEngine,
    logger: PipelineRunLogger,
    state: Optional[AccountState],
    snapshot: Callable[[str], Dict[str, Any]],
    attempts: List[Dict[str, Any]],
) -> bool:
    """Repair the proven defects in place before spending a full-chain retry.

    Returns True when the stored text was replaced by a verified patch. A retry
    today is 3 full calls; a patch is 0 (deterministic) or 1 ~300-token call.

    This is the ONLY caller of the repair module in production, and it is behind
    processing.feedback_loop.repair_mode, which defaults to "regenerate" — an
    unconfigured deployment never reaches this function.
    """
    entry = results.get(key) or {}
    validation = entry.get("agent_4_validation") or {}
    reviews = validation.get("clause_reviews") or []
    defective = count_defective_clauses(reviews)
    # A ratio-only trigger means nothing was PROVEN wrong — there is no defect
    # with a known correct value, so there is nothing a patch could act on.
    if not defective or not repairable_reviews(defective):
        return False
    content = str(validation.get("final_content") or "")
    df = dfs.get(key)
    if not content.strip() or df is None:
        return False
    try:
        statement_type = prompt_manager.get_mapping_component(key, component="type") or ""
    except Exception:  # pragma: no cover - defensive
        statement_type = ""

    codes = sorted({str(r.get("code")) for r in repairable_reviews(defective)})
    if state is not None:
        state.transition(PHASE_DEFECTIVE, "repair_gate: %s defective clause(s)" % len(defective))
        state.transition(PHASE_REPAIRING, "patch_first: %s" % ",".join(codes))
    try:
        outcome = repair_content(
            content=content,
            clause_reviews=reviews,
            df=df,
            sibling_dfs=_sibling_dfs_for(key, dfs, prompt_manager, statement_type),
            language=getattr(ai_helper, "language", "Eng"),
            statement_type=statement_type,
            mapping_key=key,
            ai_helper=ai_helper,
            logger=logger,
            style_pack=PromptStylePack(getattr(ai_helper, "language", "Eng")),
        )
    except Exception as exc:  # pragma: no cover - a repair must never break a run
        logger.logger.warning("[Repair] %s: %s", key, exc)
        if state is not None:
            state.transition(PHASE_REGENERATING, "repair_error: %s" % str(exc)[:80])
        return False

    if outcome["log"]:
        results[key].setdefault("repair_log", []).extend(outcome["log"])
        if state is not None:
            state.repairs.extend(outcome["log"])
    if not outcome["changed"]:
        reason = next((str(item.get("reason")) for item in reversed(outcome["log"]) if item.get("reason")), "no patch")
        logger.logger.info("[Repair] %s: no patch applied (%s) — falling through to regeneration", key, reason)
        if state is not None:
            state.transition(PHASE_REGENERATING, "repair_guard_failed: %s" % reason[:80])
        return False

    # All three keys, unconditionally, mirroring the arbiter's own triple below.
    # subagent_4 is not optional: _snapshot resolves content as subagent_4 ->
    # subagent_2 -> subagent_1 and never reads "final", and most accounts have no
    # subagent_4 at this point (the Validator is selective). Written BEFORE the
    # snapshot, or the attempt the arbiter ranks is the unpatched text.
    verified = [item for item in outcome["log"] if item.get("verified")]
    results[key]["subagent_4"] = outcome["content"]
    results[key]["agent_4_validation"] = {
        "final_content": outcome["content"],
        "clause_reviews": outcome["clause_reviews"],
    }
    results[key]["final"] = outcome["content"]
    attempts.append(snapshot("patch_%s" % len(verified)))
    if state is not None:
        state.transition(
            PHASE_GROUNDED,
            "patch_verified: %s" % ",".join(str(item.get("code")) for item in verified),
        )
    logger.logger.info(
        "[Repair] %s: %s patch(es) verified (%s) — retry avoided for those defect(s)",
        key, len(verified), ",".join(str(item.get("mode")) for item in verified),
    )
    return True


def _run_feedback_loop_for_key(
    key: str,
    dfs: Dict[str, pd.DataFrame],
    results: Dict[str, Dict[str, str]],
    ai_helper,
    prompt_manager: PromptEngine,
    logger: PipelineRunLogger,
    feedback_config: Dict[str, Any],
    user_comments: Optional[Dict[str, str]] = None,
    progress_callback: Optional[Callable[..., None]] = None,
    health: Optional[_RunHealth] = None,
    run_state: Optional[RunState] = None,
) -> int:
    """Run feedback loop for a single key. Returns number of retries performed.

    Transitions ARE written from here even though this can run on a worker
    thread — which does not break the one-writer rule: the pool fans out over
    KEYS, so exactly one thread ever touches this account's AccountState, and
    the ordering the rule protects is per-account.
    """
    max_retries = int(feedback_config.get("max_retries", 2))
    threshold = float(feedback_config.get("unsupported_threshold", 0.3))
    state = run_state.account(key) if run_state is not None else None

    def _snapshot(label: str) -> Dict[str, Any]:
        """One attempt, with the score the arbiter below ranks it by."""
        import copy  # module-scope import isn't available in this file section
        entry = results.get(key) or {}
        validation = entry.get("agent_4_validation") or {}
        reviews = validation.get("clause_reviews") or []
        return {
            "label": label,
            "content": entry.get("subagent_4") or entry.get("subagent_2") or entry.get("subagent_1") or "",
            "validation": copy.deepcopy(validation),
            "defects": len(count_defective_clauses(reviews)),
            "reviewed": len(reviews),
        }

    attempts: List[Dict[str, Any]] = [_snapshot("original")]

    # Patch-first: try to repair the proven defects in place BEFORE spending a
    # full three-call retry on them. Off by default (repair_mode "regenerate"),
    # so an unconfigured deployment behaves exactly as it does today. If the
    # patch clears the account, the loop below finds nothing to do and returns
    # 0 retries; if a guard refuses, the account is already in REGENERATING and
    # the loop re-reads the same defect on its way there.
    if str(feedback_config.get("repair_mode", "regenerate") or "regenerate").strip().lower() == "patch_first":
        _attempt_patch_first(
            key=key, dfs=dfs, results=results, ai_helper=ai_helper,
            prompt_manager=prompt_manager, logger=logger, state=state,
            snapshot=_snapshot, attempts=attempts,
        )

    for retry_num in range(1, max_retries + 1):
        needs_feedback, ratio, unsupported = _evaluate_feedback_needed(results, key, threshold)
        if not needs_feedback:
            if state is not None:
                state.attempts = [
                    {"label": a["label"], "defects": a["defects"], "reviewed": a["reviewed"]}
                    for a in attempts
                ]
            return retry_num - 1
        if state is not None:
            state.transition(
                PHASE_DEFECTIVE,
                "defect_gate: %s unsupported clause(s), ratio %.2f" % (len(unsupported), ratio),
            )

        logger.logger.info(
            "[FeedbackLoop] %s: retry %s/%s (unsupported_ratio=%.2f, threshold=%.2f, unsupported_count=%s)",
            key, retry_num, max_retries, ratio, threshold, len(unsupported),
        )

        feedback_text = format_validator_feedback_for_reprompt(
            (results[key].get("agent_4_validation") or {}).get("clause_reviews", []),
            ai_helper.language,
        )

        if logger.debug_mode:
            logger.log_debug("FEEDBACK_LOOP", key,
                "Retry %s/%s: feedback_text_len=%s" % (retry_num, max_retries, len(feedback_text)),
                feedback_text)

        previous_output = get_pipeline_result_text(results[key])
        base_user_comment = (user_comments or {}).get(key, "")
        combined_comment = ("%s\n\n%s" % (base_user_comment, feedback_text)).strip() if feedback_text else base_user_comment

        # Re-run generator with feedback.
        # dfs is passed for the same reason the Validator call below does: it is
        # what feeds _build_peer_context and the sibling grounding. Omitting it
        # here meant a retried account was re-generated from a STRICTLY SMALLER
        # prompt than the attempt that failed — no peer context, no siblings —
        # so the retry was handicapped exactly where it needed the most help.
        _key, gen_content, _meta = process_single_agent_item(
            "subagent_1", key, dfs.get(key), ai_helper, prompt_manager, logger,
            previous_output=previous_output,
            user_comment=combined_comment,
            dfs=dfs,
            health=health,
        )
        results[key]["subagent_1"] = gen_content
        results[key]["feedback_retry_%s_agent_1" % retry_num] = gen_content
        # These two writes bypass _store_agent_result, so the transitions have
        # to be explicit. Routing them through the store instead is NOT a
        # refactor: it would file an agent_4_validation record that the Auditor
        # write below currently discards, changing what the next
        # _evaluate_feedback_needed reads. This is not a rare path — 99 of 4,219
        # archived account records carry feedback_retries.
        if state is not None:
            state.transition(PHASE_REGENERATING, "feedback_retry_%s" % retry_num)
            state.transition(PHASE_DRAFTED, "generator_retry_%s" % retry_num)

        # Re-run Auditor (polish) so the validator sees refined output, not raw
        # generator output. Skipping this step caused more clauses to be flagged
        # as unsupported and triggered unnecessary further retries.
        _key, audit_content, _audit_meta = process_single_agent_item(
            "subagent_2", key, dfs.get(key), ai_helper, prompt_manager, logger,
            previous_output=gen_content,
            user_comment=base_user_comment,
            # Without dfs the Auditor loses both its peer context AND its
            # sibling_dfs, which is what verify_commentary grounds a figure
            # borrowed from another account against — so a retry produced
            # clause_reviews built on less evidence than the original pass.
            dfs=dfs,
            health=health,
        )
        results[key]["subagent_2"] = audit_content
        results[key]["feedback_retry_%s_agent_2" % retry_num] = audit_content
        if state is not None and (_audit_meta or {}).get("clause_reviews"):
            state.transition(PHASE_GROUNDED, "arithmetic_retry_%s" % retry_num)

        # Re-run validator on the polished output
        _key, val_content, val_metadata = process_single_agent_item(
            "subagent_4", key, dfs.get(key), ai_helper, prompt_manager, logger,
            previous_output=audit_content,
            user_comment=base_user_comment,
            dfs=dfs,
            health=health,
        )
        _store_agent_result(results, key, "subagent_4", val_content, val_metadata, state=state)
        results[key]["feedback_retry_%s_agent_4" % retry_num] = val_content
        results[key]["feedback_retries"] = retry_num
        attempts.append(_snapshot("retry_%s" % retry_num))

        if progress_callback:
            try:
                progress_callback(5, "FeedbackLoop-%s" % retry_num, 0, 0, 0, key)
            except Exception:
                pass

    # --- Final arbiter -----------------------------------------------------
    # Retries are exhausted and the output is STILL defective. Leaving
    # whatever the last pass happened to produce is wrong: a re-generation
    # is not monotonically better, and the last attempt can easily be the
    # worst of the three. Pick the attempt with the fewest hallucinated
    # clauses -- the deterministic number-grounding check in
    # verify_commentary, not an LLM judgment -- and break ties toward the
    # LATEST, which has had the most correction feedback applied.
    still_bad, _ratio, _unsupported = _evaluate_feedback_needed(results, key, threshold)
    if state is not None:
        state.attempts = [
            {"label": a["label"], "defects": a["defects"], "reviewed": a["reviewed"]}
            for a in attempts
        ]
    if still_bad and attempts:
        best = min(attempts, key=lambda a: (a["defects"], -attempts.index(a)))
        if state is not None:
            # The last retry's Validator left this GROUNDED; the gate has just
            # said it is still defective, so the path goes back through
            # DEFECTIVE before the arbiter terminates it.
            state.transition(
                PHASE_DEFECTIVE, "defect_gate_after_%s retry(ies)" % max_retries,
            )
            state.transition(
                PHASE_ARBITRATED,
                "retries_exhausted: kept '%s' (%s defect(s))" % (best["label"], best["defects"]),
            )
        results[key]["feedback_arbiter"] = {
            "chosen": best["label"],
            "scores": {a["label"]: a["defects"] for a in attempts},
        }
        if best["label"] != attempts[-1]["label"] and best["content"]:
            results[key]["subagent_4"] = best["content"]
            results[key]["agent_4_validation"] = best["validation"]
            # "final" too, or this whole arbiter is a no-op: _store_agent_result
            # set it to the LAST attempt when that attempt's validator result
            # was stored, and every consumer resolves text through
            # get_pipeline_result_text, whose priority reads "final" FIRST.
            # Overwriting only subagent_4 left the deck showing the attempt the
            # arbiter had just rejected.
            results[key]["final"] = best["content"]
            logger.logger.warning(
                "[FeedbackLoop] %s: %s retries did not clear all hallucinations; "
                "kept '%s' (%s defect(s)) over the final attempt (%s defect(s))",
                key, max_retries, best["label"], best["defects"], attempts[-1]["defects"],
            )
        else:
            logger.logger.warning(
                "[FeedbackLoop] %s: %s retries did not clear all hallucinations; "
                "keeping the final attempt (%s defect(s)) — needs human review",
                key, max_retries, attempts[-1]["defects"],
            )

    return max_retries


def run_ai_pipeline(
    mapping_keys: List[str],
    dfs: Dict[str, pd.DataFrame],
    model_type: str = "deepseek",
    language: str = "Eng",
    use_heuristic: bool = False,
    use_multithreading: bool = True,
    max_workers: Optional[int] = None,
    user_comments: Optional[Dict[str, str]] = None,
) -> Dict[str, Dict[str, str]]:
    """Simple wrapper without progress callback."""
    return run_ai_pipeline_with_progress(
        mapping_keys=mapping_keys,
        dfs=dfs,
        model_type=model_type,
        language=language,
        use_heuristic=use_heuristic,
        use_multithreading=use_multithreading,
        max_workers=max_workers,
        progress_callback=None,
        user_comments=user_comments,
    )


def run_generator_reprompt(
    mapping_keys: List[str],
    dfs: Dict[str, pd.DataFrame],
    existing_results: Optional[Dict[str, Dict[str, str]]] = None,
    model_type: str = "deepseek",
    language: str = "Eng",
    use_heuristic: bool = False,
    user_comments: Optional[Dict[str, str]] = None,
    model_name: Optional[str] = None,
) -> Dict[str, Dict[str, str]]:
    """Regenerate selected items, then immediately revalidate the revised output.

    IN SCOPE for run state, deliberately. This is a second live entry point that
    never touches run_ai_pipeline_with_progress — its own bare `results` dict,
    Generator then Validator with the Auditor skipped, direct writes that never
    reach _store_agent_result, no sweep and no feedback loop — and its output is
    merged into session_state.ai_results and exported. Leaving it stateless
    would mean an account could reach the deck with no recorded path at all.

    What it does NOT get is the RUN_STATE_KEY sentinel: this dict is MERGED into
    an existing ai_results, so a sentinel written here would travel into a
    results dict that already has one from the full run. The audit log is
    written to this reprompt's own run folder instead.
    """
    # Mirror run_ai_pipeline_with_progress: normalise "Chn" → "Chi" so the Chinese
    # reprompt path resolves prompts and applies Chinese styling (not English).
    language = normalize_language_code(language)
    logger = PipelineRunLogger()
    prompt_manager = get_prompt_engine()
    ai_helper = AIClient(
        model_type=model_type,
        agent_name="content_pipeline",
        language=language,
        use_heuristic=use_heuristic,
        model_name=model_name,
    )
    results: Dict[str, Dict[str, str]] = {}
    run_state = RunState(dfs, mapping_keys, run_folder=logger.run_folder)

    logger.logger.info(
        "Starting reprompt + validator flow with %s items | model=%s | language=%s",
        len([key for key in mapping_keys if key in dfs]),
        ai_helper.model_type,
        language,
    )

    for key in mapping_keys:
        if key not in dfs:
            continue
        existing_result = (existing_results or {}).get(key) or {}
        previous_output = ""
        if isinstance(existing_result, dict):
            for field in ("final", "subagent_4", "subagent_3", "subagent_2", "subagent_1"):
                candidate = existing_result.get(field)
                if candidate and str(candidate).strip():
                    previous_output = str(candidate)
                    break

        state = run_state.account(key)

        mapping_key, content, _metadata = process_single_agent_item(
            "subagent_1",
            key,
            dfs.get(key),
            ai_helper,
            prompt_manager,
            logger,
            previous_output=previous_output,
            user_comment=(user_comments or {}).get(key, ""),
            # dfs for the same reason the Validator call below already passes
            # it: without it this regeneration loses BOTH its peer context and
            # the sibling grounding, so a reprompted account was rebuilt from a
            # strictly smaller prompt than the pass it is replacing — the same
            # defect that was fixed on the feedback-retry path.
            dfs=dfs,
            run_state=run_state,
        )
        updated_result = dict(existing_result) if isinstance(existing_result, dict) else {}
        updated_result["subagent_1"] = content
        if state is not None:
            state.mark_stage("subagent_1")
            state.transition(PHASE_DRAFTED, "generator_reprompt")
            for _event in (_metadata or {}).pop("phase_events", None) or []:
                if _event.get("to"):
                    state.transition(_event["to"], _event.get("cause", ""))
                if _event.get("degraded"):
                    state.note_degraded(_event["degraded"])

        _validator_key, validator_content, validator_metadata = process_single_agent_item(
            "subagent_4",
            key,
            dfs.get(key),
            ai_helper,
            prompt_manager,
            logger,
            previous_output=content,
            user_comment=(user_comments or {}).get(key, ""),
            dfs=dfs,
            run_state=run_state,
        )
        _validator_events = (
            (validator_metadata or {}).pop("phase_events", None)
            if isinstance(validator_metadata, dict) else None
        )
        updated_result["subagent_4"] = validator_content
        updated_result["agent_4_validation"] = validator_metadata
        updated_result["final"] = validator_content
        updated_result["reprompt_mode"] = "generator_reprompt_validated"
        results[mapping_key] = updated_result
        if state is not None:
            state.mark_stage("subagent_4")
            # DRAFTED -> GROUNDED with the Auditor skipped: this path runs the
            # Generator then the Validator, and verify_commentary fires inside
            # the Validator stage, so the grounding still happens — one stage
            # later than in the full pipeline.
            state.transition(PHASE_GROUNDED, "llm_validator_reprompt")
            for _event in _validator_events or []:
                if _event.get("to"):
                    state.transition(_event["to"], _event.get("cause", ""))
                if _event.get("degraded"):
                    state.note_degraded(_event["degraded"])

    _sweep_final_phases(results, run_state, logger)
    _write_run_audit_log(logger, run_state, results)
    logger.finalize(results)
    return results


def save_results(results: Dict[str, Dict[str, str]], output_path: str = "fdd_utils/output/results.yml"):
    """Persist pipeline results to YAML."""
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as file:
        yaml.dump(results, file, default_flow_style=False, allow_unicode=True)


def extract_final_contents(results: Dict[str, Dict[str, str]]) -> Dict[str, str]:
    """Extract only the final content from the pipeline result payload."""
    return {
        key: value["final"]
        for key, value in results.items()
        if isinstance(value, dict) and "final" in value
    }


__all__ = [
    "RUN_HEALTH_KEY",
    "RUN_STATE_KEY",
    "SUBAGENT_SEQUENCE",
    "TERMINAL_PHASES",
    "AccountState",
    "RunState",
    "apply_house_style",
    "clean_agent_output",
    "extract_final_contents",
    "load_prompts_and_format",
    "map_value_to_component",
    "run_ai_pipeline",
    "run_ai_pipeline_with_progress",
    "run_generator_reprompt",
    "save_results",
]
