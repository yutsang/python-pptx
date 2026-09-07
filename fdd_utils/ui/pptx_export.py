from __future__ import annotations

# re-added: bound by an import in another section of the pre-split module
from ..ai import (
    FDDConfig,
    WORKBENCH_AVAILABLE_MODELS,
    build_highlighted_commentary_html,
    get_default_config_path,
    get_prompt_engine,
    is_provider_ready,
    load_yaml_config,
    parse_validator_response,
    run_ai_pipeline_with_progress,
    run_generator_reprompt,
)
from ..financial_common import extract_result_text_content, get_pipeline_result_text
import pandas as pd


from .views import derive_reconciliation_matched_keys, detect_statement_mode
from .ai_panel import build_selected_pipeline_dfs, effective_mappings_from_session
import datetime as dt_module
import logging
import os
import re
import time
from typing import Any, Callable, Dict, List, Optional

import streamlit as st

from ..pptx import build_pptx_structured_payloads
from ..workbook import find_mapping_key, get_effective_mappings, load_mappings

logger = logging.getLogger(__name__)


def generate_pptx_presentation(
    *,
    session_state: Any,
    pptx_available: bool,
) -> None:
    if not session_state.ai_results:
        st.error("❌ No AI results available. Generate AI content first.")
        return

    if not pptx_available:
        st.error("❌ PPTX generation not available. Missing required modules.")
        return

    project_name = session_state.get("project_name", "Project")
    entity_name = session_state.get("entity_name", project_name)
    language = session_state.get("language", "Eng")
    mappings = effective_mappings_from_session(session_state)

    template_path = None
    for template in ["fdd_utils/template.pptx", "template.pptx"]:
        if os.path.exists(template):
            template_path = template
            break
    if not template_path:
        st.error("❌ PowerPoint template not found. Please add `fdd_utils/template.pptx` or `template.pptx`.")
        return

    output_dir = "fdd_utils/output"
    os.makedirs(output_dir, exist_ok=True)

    timestamp = dt_module.datetime.now().strftime("%Y%m%d_%H%M%S")
    sanitized_entity = re.sub(r"[^\w\-_]", "_", str(entity_name)).strip("_") or "Project"
    selected_pipeline_dfs = build_selected_pipeline_dfs(session_state)

    try:
        combined_output_path = os.path.join(output_dir, f"{sanitized_entity}_{timestamp}.pptx")
        from fdd_utils.pptx import export_pptx_from_structured_data_combined

        structured_payloads = build_pptx_structured_payloads(
            ai_results=session_state.ai_results,
            mappings=mappings,
            bs_is_results=session_state.bs_is_results,
            dfs=selected_pipeline_dfs,
        )
        bs_data = structured_payloads.get("BS", [])
        is_data = structured_payloads.get("IS", [])
        logger.debug("PPTX payload account counts | BS=%s | IS=%s", len(bs_data), len(is_data))

        if not bs_data and not is_data:
            st.error("❌ No content generated for PPTX")
            logger.debug(
                "PPTX payload is empty | ai_results_keys=%s | dfs_keys=%s",
                list(session_state.ai_results.keys())[:10] if session_state.ai_results else "None",
                list(selected_pipeline_dfs.keys())[:10] if selected_pipeline_dfs else "None",
            )
            return

        # Demo mode: skip coSummaryShape AI so export is instant.
        _demo_cfg2 = (FDDConfig().config or {}).get("demo", {})
        _is_demo_pptx = bool(
            _demo_cfg2.get("filename") and
            str(session_state.get("uploaded_filename") or "").strip() == str(_demo_cfg2.get("filename") or "").strip()
        )
        # embed_financial_tables reads (temp_path, selected_sheet) for its
        # currency-unit-label detection and as a fresh-extraction fallback --
        # when this entity's financials came from an uploaded roll-up
        # workbook (the "進階：主表" expander), that source -- not this
        # entity's own file/selected_sheet -- is the one that actually holds
        # them. Mirrors process_workbook_data's own precedence (financials_from
        # or temp_path).
        _rollup_temp_path = session_state.get("rollup_temp_path")
        _financials_workbook_path = _rollup_temp_path or session_state.get("temp_path")
        _financials_sheet_name = (
            session_state.get("rollup_sheet") if _rollup_temp_path else session_state.get("selected_sheet")
        )

        with st.spinner("Generating PPTX…"):
            export_pptx_from_structured_data_combined(
                template_path,
                bs_data,
                is_data,
                combined_output_path,
                entity_name,
                language="chinese" if language == "Chn" else "english",
                temp_path=_financials_workbook_path,
                selected_sheet=_financials_sheet_name,
                is_chinese_databook=(language == "Chn"),
                bs_is_results=session_state.get("bs_is_results"),
                model_type=session_state.get("model_type", "local"),
                model_name=session_state.get("model_name"),
                skip_summary_ai=False,
                pre_generated_summaries=session_state.get("section_summaries") or None,
                mappings=mappings,
            )
        if os.path.exists(combined_output_path):
            with open(combined_output_path, "rb") as handle:
                session_state.pptx_download_data = handle.read()
            session_state.pptx_download_filename = os.path.basename(combined_output_path)
            session_state.pptx_download_mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            session_state.pptx_ready = True

    except Exception as exc:
        st.error(f"❌ PPTX generation failed: {exc}")
        import traceback

        st.code(traceback.format_exc())


def batch_extract_entity_data(
    *,
    temp_path: str,
    entity_name: str,
    selected_sheet: Optional[str] = None,
    financials_from: Optional[str] = None,
    financials_sheet: Optional[str] = None,
    mapping_overrides: Optional[Dict[str, str]] = None,
    language: Optional[str] = None,
) -> Dict[str, Any]:
    """Phase 1 of the batch entity pipeline (process + reconcile -- fast).

    Split out from what used to be one single batch_process_entity() call
    so a checkpoint-based batch UI can st.rerun() BETWEEN this and
    batch_run_ai_for_entity() (phase 2, slow). Streamlit only paints the
    browser once a script run returns control to it: an on_data_ready-style
    callback fired midway through one long blocking call can update
    session_state all it wants, but the switcher/data-view code living
    later in that SAME script run still can't actually render anything
    until the whole call returns -- and it doesn't return until AI
    generation is ALSO done, defeating the entire point of showing data
    before AI finishes. Only an actual rerun between two separately
    checkpointed phases makes the browser paint an intermediate state.

    financials_from/financials_sheet point BS/IS extraction at a sibling
    roll-up ("主表") workbook's named sheet when this entity's own file has
    no Financials-pattern sheet of its own — same mechanism
    process_workbook_data already exposes for the single-file flow.

    Returns {"status": "failed", "entity_name", "error"} on failure, or on
    success {"status": "ok", "entity_name", "data_summary": {the same
    "state"-shaped partial bundle a caller can swap into st.session_state
    and pass to render_data_tables_section for the full recon+breakdown
    view, plus match-count summaries}, "_internal": {everything
    batch_run_ai_for_entity needs to continue without re-deriving it}}.
    """
    from ..workbook import process_workbook_data

    result: Dict[str, Any] = {"entity_name": entity_name, "status": "ok"}

    try:
        state = process_workbook_data(
            temp_path=temp_path,
            entity_name=entity_name,
            selected_sheet=selected_sheet,
            mapping_overrides=mapping_overrides,
            financials_from=financials_from,
            financials_sheet=financials_sheet,
        )
    except Exception as exc:
        result["status"] = "failed"
        result["error"] = f"Processing failed: {exc}"
        return result

    dfs = state.get("dfs") or {}
    if not dfs:
        result["status"] = "failed"
        result["error"] = "No schedule tabs could be extracted from this databook."
        return result

    reconciliation = state.get("reconciliation")
    resolution = state.get("resolution")
    mappings = get_effective_mappings(load_mappings(), resolution)

    # Raw process_workbook_data language is "Eng"/"Chi" (workbook.py's own
    # detection convention); normalise to the UI's "Eng"/"Chn" convention so
    # this matches every == "Chn" check generate_pptx_presentation makes,
    # unless the caller already passed an explicit override in that form.
    if language:
        effective_language = language
    else:
        raw_language = str(state.get("language") or "Eng").strip()
        effective_language = "Chn" if raw_language in ("Chi", "Chn", "chinese", "Chinese") else "Eng"

    statement_mode = detect_statement_mode(reconciliation)
    if statement_mode in ("is_only", "bs_only"):
        target_type = "IS" if statement_mode == "is_only" else "BS"
        matched_mapping_keys = [
            k for k in dfs
            if mappings.get(find_mapping_key(k, mappings) or k, {}).get("type") == target_type
        ]
        if not matched_mapping_keys:
            matched_mapping_keys = list(dfs.keys())
    else:
        matched_mapping_keys = derive_reconciliation_matched_keys(reconciliation, dfs.keys(), resolution, dfs=dfs)
        has_reconciliation_data = bool(
            reconciliation and any(recon_df is not None and not recon_df.empty for recon_df in reconciliation)
        )
        if not has_reconciliation_data:
            matched_mapping_keys = list(dfs.keys())

    if not matched_mapping_keys:
        result["status"] = "failed"
        result["error"] = "No eligible accounts after reconciliation filtering."
        return result

    bs_recon, is_recon = (list(reconciliation) + [None, None])[:2] if reconciliation else (None, None)
    result["data_summary"] = {
        "entity_name": entity_name,
        "accounts_total": len(dfs),
        "accounts_matched": len(matched_mapping_keys),
        "bs_match_counts": bs_recon["Match"].value_counts().to_dict() if bs_recon is not None and not bs_recon.empty else {},
        "is_match_counts": is_recon["Match"].value_counts().to_dict() if is_recon is not None and not is_recon.empty else {},
        # Raw per-account reconciliation breakdowns (same DataFrames
        # render_reconciliation_section uses in the interactive single-file
        # flow) -- so a caller can show the actual account-by-account
        # table, not just the match-status counts.
        "bs_recon_df": bs_recon,
        "is_recon_df": is_recon,
        # Full session_state-shaped (minus ai_results/pptx) partial bundle
        # -- lets a caller swap this into st.session_state and call
        # render_data_tables_section() for the complete per-account
        # breakdown view (cash, investment properties, etc., not just
        # reconciliation), the same rich view a fully-finished entity
        # gets, while AI is still running (or hasn't started yet).
        "state": {
            "dfs": dfs,
            "display_dfs": state.get("display_dfs"),
            "workbook_list": state.get("workbook_list"),
            "display_workbook_list": state.get("display_workbook_list"),
            "language": effective_language,
            "bs_is_results": state.get("bs_is_results"),
            "reconciliation": reconciliation,
            "resolution": resolution,
            "entity_name": entity_name,
        },
    }
    result["_internal"] = {
        "raw_state": state,
        "dfs": dfs,
        "reconciliation": reconciliation,
        "resolution": resolution,
        "mappings": mappings,
        "matched_mapping_keys": matched_mapping_keys,
        "effective_language": effective_language,
        "entity_name": entity_name,
        "temp_path": temp_path,
        "selected_sheet": selected_sheet,
        "financials_from": financials_from,
        "financials_sheet": financials_sheet,
        "mapping_overrides": mapping_overrides,
    }
    return result


def build_section_summaries(
    *,
    ai_results: Dict[str, Any],
    mappings: Dict[str, Any],
    is_chinese_db: bool,
    model_type: Optional[str] = None,
    model_name: Optional[str] = None,
    label: str = "",
) -> Dict[str, str]:
    """One LLM-written executive summary per statement, keyed "BS"/"IS".

    Must be called BEFORE export. `export_pptx_from_structured_data_combined`
    deliberately makes no LLM call of its own -- an in-export call was
    reported to hang 10+ minutes on a flaky API -- so when nothing is passed
    to its `pre_generated_summaries`, the band falls back to
    `_generate_page_summary`, which splices each account's opening sentence.
    That fallback is why a summary can read as a verbatim copy of the first
    bullet on the page: it is not a summary at all, and nothing in the export
    log says so.

    Shared by the batch UI path and inspect_databook.py's CLI export, which
    is the whole reason it is a function rather than inline: the CLI had no
    equivalent and therefore never produced an AI summary at all.
    """
    from ..pptx import PowerPointGenerator

    section_summaries: Dict[str, str] = {}
    try:
        blobs: Dict[str, List[str]] = {"BS": [], "IS": []}
        for account_key, ai_result in (ai_results or {}).items():
            mapping_key = find_mapping_key(account_key, mappings)
            if not mapping_key or mapping_key not in mappings:
                continue
            atype = mappings[mapping_key].get("type")
            if atype not in blobs:
                continue
            text = extract_result_text_content(
                (ai_result or {}).get("final")
                or (ai_result or {}).get("subagent_4")
                or (ai_result or {}).get("subagent_2")
                or (ai_result or {}).get("subagent_1")
                or ""
            )
            if not text.strip():
                continue
            # A page/section summary needs each account's lead-in theme
            # only, never a table account's per-component "-"/"➢" detail
            # bullets after "明细如下：" -- see strip_table_detail_for_
            # summary's own docstring for the real corrupted-coSummaryShape
            # bug this prevents.
            blobs[atype].append(
                PowerPointGenerator.strip_table_detail_for_summary(text, is_chinese_db)
            )
        for stmt, blob in blobs.items():
            if not blob:
                continue
            try:
                from ..ai import _PIPELINE_BREAKER
                if any(_PIPELINE_BREAKER.is_open(stage) for stage in ("subagent_1", "subagent_2")):
                    continue
            except Exception:
                pass
            summary = PowerPointGenerator.generate_section_summary(
                "\n\n".join(blob),
                is_chinese=is_chinese_db,
                language=("chinese" if is_chinese_db else "english"),
                model_type=model_type,
                model_name=model_name,
            )
            if summary:
                section_summaries[stmt] = summary
    except Exception as exc:
        logger.warning(
            "Section summary generation failed for %s (PPTX summary band will fall "
            "back to spliced first sentences): %s", label or "?", exc,
        )
        return {}
    return section_summaries


# ---------------------------------------------------------------------------
# Internal insight summary (N3 v1) -- for the reviewer, NEVER for the deck
# ---------------------------------------------------------------------------

#: Sentinel key a caller may use if the summary must travel with `ai_results`.
#: Follows the __run_health__ / __BS_summary__ convention: it is not a
#: mapping_key, so build_pptx_structured_payloads' `find_mapping_key` guard
#: (payloads.py:484-486) skips it and no insight text can reach a slide.
#: build_insight_summary itself neither reads nor writes this key -- the CLI
#: keeps the summary beside the results, not inside them.
INSIGHT_SUMMARY_KEY = "__insight_summary__"

#: A resolution score this close to the 45.0 acceptance floor won a mapping by
#: a margin no one should trust. The floor itself is in the resolver; this is
#: only the width of the band that gets escalated to a human.
_RESOLUTION_FLOOR = 45.0
_RESOLUTION_WARN_BAND = 10.0

#: build_significant_movements ranks every row of the analysis frame, total and
#: subtotal rows included. Those restate the account's own movement, which
#: already has its own finding, so they are dropped from the question list.
_TOTAL_LABELS = ("total", "subtotal", "合计", "總計", "小计", "小計", "总计", "合計")

#: _change_direction's vocabulary, in words a reviewer can put in an email.
_MOVEMENT_VERBS = {
    "increase": "rise", "decrease": "fall", "flat": "stay flat",
    "new_increase": "appear from nil", "new_decrease": "appear from nil as a negative",
}


def _is_total_label(description: str) -> bool:
    low = str(description or "").strip().lower()
    return any(marker in low for marker in _TOTAL_LABELS)


def _sheet_name_of(entry: Any) -> str:
    """`unresolved_sheets` entries used to be bare names and are now per-sheet
    diagnostic dicts (sheet_name, best_score, floor_missed_by, reason ...).
    Both shapes appear in archived runs, so read the name out of either --
    printing the whole dict inline made a finding unreadable."""
    if isinstance(entry, dict):
        return str(entry.get("sheet_name") or entry.get("name") or entry)
    return str(entry)


def _insight_num(value: Any) -> Optional[float]:
    """float or None. numpy scalars coerced deliberately: this dict is meant to
    be yaml.dump'd into the run folder, and a numpy.int64 anywhere in it raises
    RepresenterError at the very end of a paid run (logging.py:243/247)."""
    try:
        if value is None or isinstance(value, bool):
            return None
        return float(value)
    except (TypeError, ValueError):
        return None


def build_insight_summary(
    *,
    ai_results: Dict[str, Any],
    dfs: Optional[Dict[str, Any]] = None,
    mappings: Optional[Dict[str, Any]] = None,
    reconciliation: Any = None,
    resolution: Optional[Dict[str, Any]] = None,
    evidence: Any = None,
    links: Any = None,
    language: str = "Eng",
    max_movements: int = 8,
) -> Dict[str, Any]:
    """What the run itself knows about its own weak spots — internal only.

    **v1 makes no LLM call.** Every field is assembled from material the
    pipeline already computes and then discards:

    - the run-health tally filed under ``__run_health__`` (how much of the
      commentary actually came from the model at all);
    - ``clause_reviews`` counts and categories per account;
    - ``feedback_retries`` / ``feedback_arbiter``;
    - ``claim_contract`` verdicts, when M5 is enabled (absent otherwise);
    - ``build_trend_summary`` / ``build_significant_movements`` over the same
      nested analysis frame the Generator prompt was built from — computed
      inside prompt rendering and thrown away there;
    - reconciliation ❌ Diff / ⚠️ Match rows;
    - resolution scores sitting within 10 points of the 45.0 floor;
    - ``build_account_mapping_diagnostics``, which has been written, exported
      through two ``__init__`` files and imported into ``ai_panel.py`` since it
      was added, and never once called;
    - ``presentation_detail_table["tie_status"]``, computed and never read.

    ``evidence`` (M2) and ``links`` (N2) are accepted and ignored until those
    milestones land — the plan's own rule is that a field sourced from a
    milestone that has not shipped is omitted, not faked.

    Returns ``{summary, visible_issues, client_questions, external_research,
    commentary_instruction, unverified_hypotheses}``. ``visible_issues`` entries
    are ``{issue, basis, evidence_ids, severity}`` where ``basis`` names the
    computation the claim came from and every ``evidence_id`` is a stable,
    self-describing handle (``clause:<account>#<n>``, ``recon:BS:<row>``,
    ``movement:<account>:<from>-><to>``, ...). When M2's EvidenceIndex lands
    these become its ids; until then they are resolvable by hand, which is the
    acceptance bar the plan set for the sources that have shipped.

    This must never enter the deck. It is not commentary: it names what the
    databook failed to answer, which is exactly what a client must not read.
    """
    from ..workbook import (
        build_account_mapping_diagnostics,
        build_significant_movements,
        build_trend_summary,
    )

    ai_results = ai_results or {}
    dfs = dfs or {}
    accounts = {k: v for k, v in ai_results.items()
                if isinstance(v, dict) and not str(k).startswith("__")}
    # With no AI results at all -- the free, no-LLM path -- the extraction-side
    # half still has everything it needs, so name the accounts from the frames
    # instead of reporting an empty run. Every AI-sourced section below is
    # keyed off `accounts` and simply produces nothing.
    account_names = sorted(accounts) or sorted(dfs)

    issues: List[Dict[str, Any]] = []
    questions: List[str] = []
    counts: Dict[str, int] = {}

    def add(issue: str, basis: str, evidence_ids: List[str], severity: str) -> None:
        issues.append({"issue": issue, "basis": basis,
                       "evidence_ids": list(evidence_ids), "severity": severity})
        counts[severity] = counts.get(severity, 0) + 1

    # -- 1. Run health --------------------------------------------------
    health = ai_results.get("__run_health__")
    if isinstance(health, dict):
        if health.get("zero_successful_calls"):
            add("Not one LLM call succeeded — every bullet in this run is deterministic "
                f"filler. First failure: {health.get('first_failure') or 'none recorded'}",
                "__run_health__.zero_successful_calls",
                ["run_health:zero_successful_calls"], "critical")
        for field, label in (
            ("accounts_on_fallback_bullet", "fell back to a deterministic bullet"),
            ("accounts_on_passthrough", "shipped an earlier stage's text unchanged"),
            ("accounts_on_error_text", "shipped error placeholder text"),
            ("accounts_without_prompt", "had no prompt at all"),
        ):
            hit = list(health.get(field) or [])
            if hit:
                add(f"{len(hit)} account(s) {label}: {', '.join(map(str, hit[:6]))}"
                    + (" ..." if len(hit) > 6 else ""),
                    f"__run_health__.{field}", [f"run_health:{field}"], "high")
        failed = int(health.get("calls_failed") or 0)
        if failed and not health.get("zero_successful_calls"):
            add(f"{failed} LLM attempt(s) failed and were retried or skipped; the deck is "
                "built from what came back.", "__run_health__.calls_failed",
                ["run_health:calls_failed"], "medium")

    # -- 2. Verifier defects -------------------------------------------
    for key, result in sorted(accounts.items()):
        reviews = ((result.get("agent_4_validation") or {}).get("clause_reviews")
                   if isinstance(result.get("agent_4_validation"), dict) else None) or []
        flagged = [(i, r) for i, r in enumerate(reviews)
                   if isinstance(r, dict) and not r.get("supported", True)]
        if not flagged:
            continue
        categories = sorted({str(r.get("category") or "?") for _i, r in flagged})
        severity = "high" if any("halluc" in c for c in categories) else "medium"
        add(f"{key}: {len(flagged)} of {len(reviews)} clause(s) unsupported ({', '.join(categories)})",
            "clause_reviews", [f"clause:{key}#{i}" for i, _r in flagged], severity)

    # -- 3. Retries and arbitration ------------------------------------
    retried = {k: v.get("feedback_retries") for k, v in accounts.items() if v.get("feedback_retries")}
    if retried:
        add(f"{len(retried)} account(s) needed a feedback retry: "
            + ", ".join(f"{k}×{v}" for k, v in sorted(retried.items())),
            "feedback_retries", [f"retry:{k}" for k in sorted(retried)], "medium")
    for key, result in sorted(accounts.items()):
        arbiter = result.get("feedback_arbiter")
        if isinstance(arbiter, dict) and arbiter:
            add(f"{key}: the arbiter kept an earlier attempt ({arbiter.get('reason') or arbiter})",
                "feedback_arbiter", [f"arbiter:{key}"], "medium")

    # -- 4. Claim contracts (M5; absent unless enabled) -----------------
    contract_failures: Dict[str, List[str]] = {}
    for key, result in accounts.items():
        verdicts = result.get("claim_contract")
        if not isinstance(verdicts, dict):
            continue
        for claim_id, verdict in verdicts.items():
            if str(verdict) == "fail":
                contract_failures.setdefault(str(claim_id).rsplit(":", 1)[-1], []).append(key)
    for detector, keys in sorted(contract_failures.items()):
        add(f"{len(keys)} account(s) did not carry out the injected '{detector}' instruction: "
            + ", ".join(sorted(keys)[:6]) + (" ..." if len(keys) > 6 else ""),
            "claim_contract", [f"claim:{k}:{detector}" for k in sorted(keys)], "medium")

    # -- 5. Movements the data shows (from the frame the prompt used) ---
    # build_trend_summary / build_significant_movements run inside
    # PromptEngine.render_prompt and are discarded there. Re-running them over
    # the same nested frame costs nothing and is the only place in this repo
    # that keeps what they found.
    movements: List[Dict[str, Any]] = []
    for key in sorted(dfs):
        df = dfs.get(key)
        attrs = getattr(df, "attrs", {}) or {}
        analysis = attrs.get("prompt_analysis_df")
        if analysis is None or getattr(analysis, "empty", True):
            continue
        # A question is only worth asking the client when the databook does not
        # already answer it. An account carrying notes or row-linked remarks
        # may well explain the movement in prose this layer cannot read, so
        # asserting "nothing states a cause" there would be a claim about
        # material never examined.
        if attrs.get("supporting_notes") or attrs.get("adjacent_detail_rows"):
            continue
        try:
            trend = build_trend_summary(analysis) or {}
            for item in build_significant_movements(analysis) or []:
                description = str(item.get("description") or "").strip()
                if _is_total_label(description):
                    # The account's own total already has its own finding (the
                    # material-movement claim); the useful question is which
                    # component moved.
                    continue
                movements.append({
                    "account": key,
                    "description": description,
                    "from_period": str(item.get("from_period") or trend.get("start_period") or ""),
                    "to_period": str(item.get("to_period") or trend.get("end_period") or ""),
                    "pct_change": _insight_num(item.get("pct_change")),
                    "delta": _insight_num(item.get("delta")),
                    "direction": str(item.get("direction") or trend.get("series_direction") or ""),
                })
        except Exception as exc:  # a malformed analysis frame must not kill the summary
            logger.debug("Insight summary: movement scan failed for %s: %s", key, exc)
    # Rank on the percentage where there is one, otherwise on the absolute
    # delta scaled below every percentage, so a from-nil movement (no
    # meaningful percentage, by build_significant_movements' own rule) still
    # sorts by size instead of collapsing to zero and sinking to the bottom.
    movements.sort(key=lambda m: (m.get("pct_change") is not None,
                                  abs(m.get("pct_change") or m.get("delta") or 0.0)),
                   reverse=True)
    for m in movements[:max_movements]:
        question = (
            f"{m['account']} — what drove {m['description'] or 'this line'} to "
            f"{_MOVEMENT_VERBS.get(m['direction'], 'move')} between {m['from_period']} and "
            f"{m['to_period']}"
            + (f" ({m['pct_change']:+.0f}%)" if m.get("pct_change") is not None else "")
            + "? This tab carries no notes or side remarks at all."
        )
        if question not in questions:
            questions.append(question)

    # -- 6. Reconciliation ---------------------------------------------
    # The four statuses mean four different things and must not be lumped:
    # '⚠️ Match' is NOT a break -- the values already agree, and the flag only
    # marks a zero sitting next to a non-zero period (the CLI's section 4
    # explains this at length). Reporting it as a finding sends a reviewer
    # chasing a number that ties.
    recon_frames = []
    if reconciliation is not None:
        recon_frames = [f for f in (reconciliation if isinstance(reconciliation, (list, tuple))
                                    else [reconciliation]) if f is not None]
    for stmt, frame in zip(("BS", "IS"), recon_frames):
        try:
            if getattr(frame, "empty", True) or "Match" not in frame.columns:
                continue
            label_col = frame.columns[0]
            for needle, severity, wording in (
                ("❌ Diff", "high",
                 "the summary and the supporting tab disagree"),
                ("⚠️ Interim Diff", "medium",
                 "the latest (partial) period differs, though the prior period ties"),
                ("⚠️ Not Found", "medium",
                 "no supporting tab was found for this summary row"),
            ):
                rows = frame[frame["Match"].astype(str).str.startswith(needle, na=False)]
                if rows.empty:
                    continue
                names = [str(v) for v in rows[label_col].tolist()]
                add(f"{stmt}: {len(names)} row(s) where {wording}: {', '.join(names[:6])}"
                    + (" ..." if len(names) > 6 else ""),
                    "reconciliation", [f"recon:{stmt}:{n}" for n in names], severity)
                if needle == "❌ Diff":
                    questions.append(
                        f"{stmt} — the summary and the supporting tab disagree on "
                        f"{', '.join(names[:4])}; which figure is the one to report?"
                    )
        except Exception as exc:
            logger.debug("Insight summary: reconciliation scan failed for %s: %s", stmt, exc)

    # -- 7. Mapping confidence -----------------------------------------
    resolved = ((resolution or {}).get("resolved") or {}) if isinstance(resolution, dict) else {}
    weak = []
    for mapping_key, chosen in resolved.items():
        score = _insight_num((chosen or {}).get("score"))
        if score is None or score >= _RESOLUTION_FLOOR + _RESOLUTION_WARN_BAND:
            continue
        weak.append((str(mapping_key), score, str((chosen or {}).get("sheet_name") or "")))
    for mapping_key, score, sheet in sorted(weak):
        add(f"{mapping_key} was matched to sheet '{sheet}' with a score of {score:.1f}, within "
            f"{_RESOLUTION_WARN_BAND:.0f} of the {_RESOLUTION_FLOOR:.0f} acceptance floor",
            "resolution.score", [f"resolution:{mapping_key}"], "medium")
        questions.append(f"Confirm that tab '{sheet}' is the supporting schedule for {mapping_key}.")
    unresolved = [_sheet_name_of(s) for s in
                  ((resolution or {}).get("unresolved_sheets") or [])] if isinstance(resolution, dict) else []
    if unresolved:
        add(f"{len(unresolved)} sheet(s) matched no account and are absent from the deck: "
            f"{', '.join(unresolved[:6])}" + (" ..." if len(unresolved) > 6 else ""),
            "resolution.unresolved_sheets",
            [f"unresolved:{s}" for s in unresolved], "medium")

    # -- 8. Accounts that will never reach a slide ----------------------
    if mappings:
        try:
            diagnostics = build_account_mapping_diagnostics(account_names, mappings)
            orphans = diagnostics[diagnostics["classification"] == "other"]
            if not orphans.empty:
                names = [str(v) for v in orphans["account_name"].tolist()]
                add(f"{len(names)} account(s) have no BS/IS mapping, so their commentary is "
                    f"generated, paid for and dropped before the deck: {', '.join(names[:6])}"
                    + (" ..." if len(names) > 6 else ""),
                    "build_account_mapping_diagnostics",
                    [f"mapping:{n}" for n in names], "high")
        except Exception as exc:
            logger.debug("Insight summary: mapping diagnostics failed: %s", exc)

    # -- 9. Detail tables that do not tie -------------------------------
    untied, unchecked = [], 0
    for key in sorted(dfs):
        table = getattr(dfs.get(key), "attrs", {}).get("presentation_detail_table")
        if not isinstance(table, dict):
            continue
        status = str(table.get("tie_status") or "")
        if not status or status == "not checked":
            unchecked += 1
        elif "differs on 0" not in status:
            untied.append((key, status))
    for key, status in untied:
        add(f"{key}: its breakdown table {status} against the account total",
            "presentation_detail_table.tie_status", [f"tie:{key}"], "medium")
    if unchecked:
        add(f"{unchecked} breakdown table(s) were never tie-tested (synthesized from the "
            "sheet rather than detected, which runs no tie test at all)",
            "presentation_detail_table.tie_status", ["tie:unchecked"], "low")

    # -- assembly -------------------------------------------------------
    order = {"critical": 0, "high": 1, "medium": 2, "low": 3}
    issues.sort(key=lambda i: order.get(i["severity"], 9))

    if not issues:
        summary = (f"{len(account_names)} account(s) with nothing flagged: no failed LLM "
                   "calls, no unsupported clauses, no reconciliation breaks and no low-confidence "
                   "mappings. Review the text itself; this layer only sees what it can compute.")
        instruction = "No deterministic defect to act on. Read the commentary on its own merits."
    else:
        head = issues[0]
        summary = (
            f"{len(account_names)} account(s); {len(issues)} deterministic finding(s) "
            + ", ".join(f"{n} {sev}" for sev, n in sorted(counts.items(), key=lambda kv: order.get(kv[0], 9)))
            + f". Most severe: {head['issue']}"
        )
        instruction = _insight_instruction(issues)

    return {
        "summary": summary,
        "visible_issues": issues,
        "client_questions": questions,
        "external_research": {
            # v1 has no deterministic trigger for this. Every finding above is
            # answerable from the databook or by asking management, and the
            # plan forbids an unverified industry claim reaching a prompt --
            # so this stays false until N4 supplies a real trigger.
            "needed": False,
            "reason": "No finding here requires information from outside the databook; "
                      "every open item is a question for management.",
        },
        "commentary_instruction": instruction,
        # Nothing in v1 may hypothesise: no LLM call, and N2's verified
        # relationships (the only sanctioned source of a cross-account guess)
        # have not landed. Kept in the shape so the field's meaning is fixed
        # before anything fills it.
        "unverified_hypotheses": [],
    }


def _insight_instruction(issues: List[Dict[str, Any]]) -> str:
    """One line telling the reviewer what to do first, keyed off the top issue's
    basis rather than its wording, so it survives a rephrasing."""
    basis = str((issues[0] or {}).get("basis") or "")
    if basis.startswith("__run_health__"):
        return ("Do not send this deck. Fix the provider/model configuration and re-run — "
                "some or all of the commentary is not model output.")
    if basis == "clause_reviews":
        return ("Read the flagged clauses against the tab before sending: the verifier could "
                "not find those figures in this account's own numbers.")
    if basis == "reconciliation":
        return ("Settle the reconciliation breaks first — commentary written on a figure the "
                "summary and the tab disagree about will have to be rewritten anyway.")
    if basis.startswith("resolution"):
        return ("Confirm the low-confidence tab-to-account matches before reading the "
                "commentary; a wrong match makes every figure in that bullet wrong.")
    if basis == "build_account_mapping_diagnostics":
        return ("Add the unmapped accounts to mappings.yml or accept that their commentary is "
                "discarded — it is being generated and paid for either way.")
    if basis == "claim_contract":
        return ("The prompt computed a figure or an opening the text did not use. Check whether "
                "the instruction or the detector is at fault before changing either.")
    return "Work the findings below in severity order."


def batch_run_ai_for_entity(
    *,
    extracted: Dict[str, Any],
    model_type: str = "local",
    model_name: Optional[str] = None,
    use_multithreading: bool = True,
    max_workers: Optional[int] = None,
    user_comments: Optional[Dict[str, str]] = None,
    template_path: Optional[str] = None,
    output_dir: str = "fdd_utils/output",
    progress_callback: Optional[Callable[..., None]] = None,
) -> Dict[str, Any]:
    """Phase 2 of the batch entity pipeline (AI generation + PPTX export --
    slow). Takes the successful result dict batch_extract_entity_data()
    returned (via its "_internal" bundle) and picks up where extraction
    left off, without re-deriving anything.

    Returns the same shape batch_process_entity's single-call version
    always did: {"status", "output_path", "bs_count", "is_count",
    "accounts_processed", "state": {full session_state-shaped bundle
    including ai_results/pptx_download_data, for swapping into
    st.session_state and reusing render_processed_view unchanged}} on
    success, {"status": "failed", "entity_name", "error"} on failure.
    """
    from ..ai import run_ai_pipeline_with_progress
    from ..pptx import export_pptx_from_structured_data_combined

    internal = extracted["_internal"]
    entity_name = internal["entity_name"]
    dfs = internal["dfs"]
    state = internal["raw_state"]
    reconciliation = internal["reconciliation"]
    resolution = internal["resolution"]
    mappings = internal["mappings"]
    matched_mapping_keys = internal["matched_mapping_keys"]
    effective_language = internal["effective_language"]
    temp_path = internal["temp_path"]
    selected_sheet = internal["selected_sheet"]
    financials_from = internal["financials_from"]
    financials_sheet = internal["financials_sheet"]
    mapping_overrides = internal["mapping_overrides"]

    result: Dict[str, Any] = {"entity_name": entity_name, "status": "ok"}

    ai_results = run_ai_pipeline_with_progress(
        mapping_keys=matched_mapping_keys,
        dfs=dfs,
        model_type=model_type,
        model_name=model_name,
        language=effective_language,
        use_multithreading=use_multithreading,
        max_workers=max_workers,
        progress_callback=progress_callback,
        user_comments=user_comments or {},
    )

    # Executive summary (coSummaryShape) generation -- mirrors what
    # render_ai_generation_section does for the single-file flow right
    # after its own per-account AI pass, which this batch path never had
    # an equivalent of. Without it, export_pptx_from_structured_data_
    # combined's own in-export summary call is SKIPPED ENTIRELY (a
    # deliberate choice there, not a bug -- an in-export LLM call was
    # reported to hang 10+ minutes when the API was flaky), leaving
    # coSummaryShape genuinely blank on every entity's first BS/IS slide.
    # Confirmed via a real batch export's --dump-text output: a literal
    # empty coSummaryShape text frame on both statements.
    section_summaries = build_section_summaries(
        ai_results=ai_results,
        mappings=mappings,
        is_chinese_db=(effective_language == "Chn"),
        model_type=model_type,
        model_name=model_name,
        label=str(entity_name),
    )

    structured_payloads = build_pptx_structured_payloads(
        ai_results=ai_results,
        mappings=mappings,
        bs_is_results=state.get("bs_is_results"),
        dfs=dfs,
    )
    bs_data = structured_payloads.get("BS", [])
    is_data = structured_payloads.get("IS", [])
    if not bs_data and not is_data:
        result["status"] = "failed"
        result["error"] = "No content generated for PPTX (empty BS and IS payloads)."
        return result

    resolved_template_path = template_path
    if not resolved_template_path:
        for candidate in ["fdd_utils/template.pptx", "template.pptx"]:
            if os.path.exists(candidate):
                resolved_template_path = candidate
                break
    if not resolved_template_path:
        result["status"] = "failed"
        result["error"] = "PowerPoint template not found (fdd_utils/template.pptx)."
        return result

    os.makedirs(output_dir, exist_ok=True)
    timestamp = dt_module.datetime.now().strftime("%Y%m%d_%H%M%S")
    sanitized_entity = re.sub(r"[^\w\-_]", "_", str(entity_name)).strip("_") or "Entity"
    output_path = os.path.join(output_dir, f"{sanitized_entity}_{timestamp}.pptx")

    # embed_financial_tables reads (temp_path, selected_sheet) for its
    # currency-unit-label detection and as a fresh-extraction fallback --
    # when financials came from a roll-up workbook, that source (not this
    # entity's own file/blank sheet) is the one that actually holds them.
    # Mirrors process_workbook_data's own precedence (financials_from or
    # temp_path).
    financials_workbook_path = financials_from or temp_path
    financials_sheet_name = financials_sheet if financials_from else selected_sheet

    export_pptx_from_structured_data_combined(
        resolved_template_path,
        bs_data,
        is_data,
        output_path,
        entity_name,
        language="chinese" if effective_language == "Chn" else "english",
        temp_path=financials_workbook_path,
        selected_sheet=financials_sheet_name,
        is_chinese_databook=(effective_language == "Chn"),
        bs_is_results=state.get("bs_is_results"),
        model_type=model_type,
        model_name=model_name,
        skip_summary_ai=False,
        pre_generated_summaries=section_summaries or None,
        mappings=mappings,
    )

    result["output_path"] = output_path
    result["bs_count"] = len(bs_data)
    result["is_count"] = len(is_data)
    result["accounts_processed"] = len(matched_mapping_keys)

    with open(output_path, "rb") as handle:
        pptx_bytes = handle.read()

    # Full session_state-shaped bundle so a caller (the batch UI) can swap
    # this entity's results into st.session_state and reuse the single-file
    # render_processed_view/generate_pptx_presentation UI UNCHANGED, instead
    # of only ever seeing this thin status dict.
    result["state"] = {
        "dfs": state.get("dfs"),
        "display_dfs": state.get("display_dfs"),
        "dfs_variants": state.get("dfs_variants"),
        "display_df_variants": state.get("display_df_variants"),
        "workbook_list": state.get("workbook_list"),
        "display_workbook_list": state.get("display_workbook_list"),
        "language": effective_language,
        "detected_language": effective_language,
        "bs_is_results": state.get("bs_is_results"),
        "reconciliation": reconciliation,
        "resolution": resolution,
        "project_name": state.get("project_name"),
        "entity_name": entity_name,
        # The financials source (not necessarily this entity's own file --
        # see financials_workbook_path/financials_sheet_name above), so a
        # later "Regenerate PPTX" click from within the reused single-file
        # UI still finds the right sheet for the embedded table instead of
        # re-hitting the same blank-selected_sheet bug this export call
        # just worked around.
        "temp_path": financials_workbook_path,
        "selected_sheet": financials_sheet_name,
        "mapping_overrides": mapping_overrides,
        "ai_results": ai_results,
        # So a later "Regenerate PPTX" click from within the reused
        # single-file UI (generate_pptx_presentation, which reads
        # session_state.section_summaries) reuses these instead of
        # falling back to the in-export summary skip.
        "section_summaries": section_summaries,
        "model_type": model_type,
        "model_name": model_name,
        "use_multithreading": use_multithreading,
        "pptx_ready": True,
        "pptx_download_data": pptx_bytes,
        "pptx_download_filename": os.path.basename(output_path),
        "pptx_download_mime": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    }
    return result


def batch_process_entity(
    *,
    temp_path: str,
    entity_name: str,
    selected_sheet: Optional[str] = None,
    financials_from: Optional[str] = None,
    financials_sheet: Optional[str] = None,
    mapping_overrides: Optional[Dict[str, str]] = None,
    model_type: str = "local",
    model_name: Optional[str] = None,
    language: Optional[str] = None,
    use_multithreading: bool = True,
    max_workers: Optional[int] = None,
    user_comments: Optional[Dict[str, str]] = None,
    template_path: Optional[str] = None,
    output_dir: str = "fdd_utils/output",
    progress_callback: Optional[Callable[..., None]] = None,
    on_data_ready: Optional[Callable[[Dict[str, Any]], None]] = None,
) -> Dict[str, Any]:
    """Headless, session_state-free equivalent of the single-file
    process -> reconcile -> AI -> export flow, for driving one entity in
    a single blocking call -- a thin composition of
    batch_extract_entity_data() then batch_run_ai_for_entity(), kept for
    callers that want one-shot headless behavior (e.g. inspect_databook.py
    -style scripts). Mirrors inspect_databook.py's inspect_one() pattern.

    A checkpoint-based batch UI that wants the browser to actually paint
    an intermediate "data ready, AI still pending" state should call the
    two phases separately across two st.rerun()s instead -- see
    fdd_app.py's render_batch_processing_section, and the phase functions'
    own docstrings for why a callback fired midway through this single
    call can't achieve that on its own.

    on_data_ready, if given, fires once (right after data extraction +
    reconciliation complete, before AI generation starts) with the
    extraction phase's "data_summary" dict.
    """
    extracted = batch_extract_entity_data(
        temp_path=temp_path,
        entity_name=entity_name,
        selected_sheet=selected_sheet,
        financials_from=financials_from,
        financials_sheet=financials_sheet,
        mapping_overrides=mapping_overrides,
        language=language,
    )
    if extracted.get("status") != "ok":
        return extracted

    if on_data_ready:
        try:
            on_data_ready(extracted["data_summary"])
        except Exception:
            pass  # a UI-side display glitch should never abort the pipeline

    return batch_run_ai_for_entity(
        extracted=extracted,
        model_type=model_type,
        model_name=model_name,
        use_multithreading=use_multithreading,
        max_workers=max_workers,
        user_comments=user_comments,
        template_path=template_path,
        output_dir=output_dir,
        progress_callback=progress_callback,
    )
# --- end ui/pptx_export.py ---


# --- Bridge Lab (experimental, isolated testing page) ---
# Lets the project team upload a workbook, pick one tab, and auto-detect a
# 'Base'/'Change' bridge helper block (the same convention as the real
# 成都-量价桥图 tab) to generate a NATIVE Excel waterfall chart -- not PPTX,
# because the team's actual downstream tool (UpSlide) links PowerPoint
# charts FROM Excel chart objects; it has no concept of a python-pptx chart,
# so a python-pptx output would be a dead end for their real workflow.
# Entirely separate code path from the main upload/process/AI/PPTX flow --
# gated behind its own session_state flag so it can't interfere with it.
import io as _bridge_lab_io

from openpyxl import Workbook as _bridge_lab_Workbook
from openpyxl import load_workbook as _bridge_lab_load_workbook

from ..bridge_chart_prototype import build_excel_waterfall_chart, find_bridge_blocks
from ..generate_bridge_waterfall_batch import build_bridges_for_ab_tab


def render_bridge_lab_toggle() -> None:
    with st.sidebar:
        st.markdown("---")
        if st.session_state.get("show_bridge_lab"):
            if st.button("← 返回主流程", use_container_width=True, key="bridge_lab_back_btn"):
                st.session_state["show_bridge_lab"] = False
                st.rerun()
        else:
            if st.button("🧪 橋圖測試 (Bridge Lab)", use_container_width=True, key="bridge_lab_enter_btn"):
                st.session_state["show_bridge_lab"] = True
                st.rerun()


def _bridge_lab_show_block(index: int, block, note: str = "") -> bool:
    """Renders one detected bridge block as a table plus a chart, and returns
    whether it's usable (check passed, or unverified -- never a hard mismatch).

    Previously printed one st.write per item, which for 11 items across
    several transitions filled the page with lines nobody reads. A table and
    the chart itself are what actually get checked, so that is all this shows;
    the per-item source notes still travel into the downloaded workbook."""
    header = f"區塊 {index + 1}：{block.items[0].label} → {block.items[-1].label}"
    with st.expander(header, expanded=True):
        if note:
            st.caption(note)

        rows = []
        running = 0.0
        for it in block.items:
            if it.kind == "total":
                running = it.value
                rows.append({"項目": it.label, "類型": "合計", "金額 (千元)": round(it.value, 1),
                             "累計": round(running, 1)})
            else:
                running += it.value
                rows.append({"項目": it.label, "類型": "變動", "金額 (千元)": round(it.value, 1),
                             "累計": round(running, 1)})
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

        try:
            _bridge_lab_render_preview_chart(block, header)
        except Exception as exc:  # a preview failure must not block the download
            st.caption(f"（圖表預覽無法產生：{exc}）")

        if block.check_ok is True:
            st.success("✅ 核對一致")
            return True
        if block.check_ok is False:
            st.error("❌ 核對不一致 -- 不會為此區塊生成圖表，請檢查來源表格")
            return False
        st.warning("⚠️ 無法核對一致性，仍會生成圖表，請自行核對數字")
        return True


def _bridge_lab_render_preview_chart(block, title: str) -> None:
    """On-screen waterfall preview using the same invisible-base-series maths
    as the downloaded Excel chart, so what is checked here is what ships."""
    from ..bridge_chart_prototype import _compute_waterfall_series

    categories, base_vals, total_vals, inc_vals, dec_vals = _compute_waterfall_series(block)
    frame = pd.DataFrame(
        {
            "（基準）": base_vals,
            "合計": total_vals,
            "增加": inc_vals,
            "減少": dec_vals,
        },
        index=categories,
    )
    st.bar_chart(frame, stack=True, height=340,
                 color=("#00000000", "#00338D", "#6D2077", "#00A3A1"))


def render_bridge_lab() -> None:
    st.title("🧪 橋圖測試 (Bridge Chart Lab)")
    st.caption(
        "實驗性功能，與主流程完全獨立。支援兩類 tab：(1) 已建好 Base/Change 輔助區塊的橋圖表，"
        "(2) AB- 原始數據表（自動計算價/量/天數因子分解）。選擇 tab 後自動偵測、生成原生 Excel "
        "疊加圖，供下載後透過 UpSlide 帶入 PPT。"
    )

    uploaded = st.file_uploader("上傳 Excel 檔案 (.xlsx)", type=["xlsx"], key="bridge_lab_upload")
    if not uploaded:
        st.info("請先上傳一個 .xlsx 檔案。")
        return

    file_bytes = uploaded.getvalue()
    try:
        wb_values = _bridge_lab_load_workbook(_bridge_lab_io.BytesIO(file_bytes), data_only=True)
    except Exception as exc:
        st.error(f"無法讀取此 Excel 檔案：{exc}")
        return

    sheet_names = wb_values.sheetnames
    selected_sheet = st.selectbox("選擇 tab", sheet_names, key="bridge_lab_sheet")

    if not st.button("偵測並生成", type="primary", key="bridge_lab_detect_btn"):
        return

    ws = wb_values[selected_sheet]

    # Route by tab type. A pre-built Base/Change helper block (like a
    # <entity>-量价桥图 tab) is read directly; otherwise try to treat it as a
    # raw AB-* data tab and COMPUTE the factor decomposition. Both paths
    # converge on a list of BridgeBlock objects rendered identically.
    renderable = []  # list of BridgeBlock that passed (or lack) their check
    prebuilt = find_bridge_blocks(ws)
    if prebuilt:
        st.success(f"偵測到 {len(prebuilt)} 個預建 Base/Change 橋圖區塊。")
        for i, block in enumerate(prebuilt):
            if _bridge_lab_show_block(i, block):
                renderable.append(block)
    else:
        ab_blocks, results = build_bridges_for_ab_tab(ws, selected_sheet)
        if ab_blocks is None:
            st.warning(
                f"「{selected_sheet}」既不是 Base/Change 結構的橋圖表，也不是可識別的 AB- 原始數據表"
                "（找不到 Year/Days 標籤列或分期區塊）。請確認選對了 tab。"
            )
            return
        if not results:
            st.warning(
                f"「{selected_sheet}」偵測為 AB- 原始數據表，但算不出任何年度轉換（可能只有單一年度資料）。"
            )
            return
        st.success(
            f"偵測為 AB- 原始數據表，已計算出 {len(results)} 個年度轉換橋圖"
            "（價/量/天數因子分解，末期採 LTM 滾動12個月口徑）。"
        )
        for i, res in enumerate(results):
            note = "註：末期為不完整年度，已改用 LTM 滾動12個月窗口比較" if res.is_ltm else ""
            if _bridge_lab_show_block(i, res.bridge, note=note):
                renderable.append(res.bridge)

    if not renderable:
        st.error("沒有通過核對、可生成圖表的區塊。")
        return

    # STANDALONE output workbook -- we build a brand-new Workbook and NEVER
    # re-save the user's upload, so their original file's formulas, cached
    # values, and existing native charts are physically untouched (openpyxl
    # round-tripping a workbook silently drops every formula's cached result,
    # which is what made an earlier version appear to "change" old tabs).
    out_wb = _bridge_lab_Workbook()
    out_wb.remove(out_wb.active)
    out_ws = out_wb.create_sheet("Bridge_Output")
    next_row = 1
    for block in renderable:
        title = f"{block.items[0].label} → {block.items[-1].label}"
        next_row = build_excel_waterfall_chart(out_ws, block, title, start_row=next_row)

    out_buffer = _bridge_lab_io.BytesIO()
    out_wb.save(out_buffer)
    st.download_button(
        "下載含橋圖的 Excel",
        data=out_buffer.getvalue(),
        file_name="bridge_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="bridge_lab_download_btn",
    )
    st.caption(
        "此為獨立的新檔案，只含生成的「Bridge_Output」分頁（結構化數據表＋橋圖）；"
        "你上傳的原始檔案完全沒有被修改。"
    )
