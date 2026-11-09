"""Export a deck from validated candidates, then look at the slides they touched.

Reuses the production path end to end -- build_pptx_structured_payloads, the
combined exporter with its packing and splitting, inspect_pptx -- and adds only
what an edit needs: the base commentary of the accounts nobody touched, the
delivery decision over the merged results, and the list of slides the edited
accounts landed on (a continuation slide has no account label of its own, so
slides are found by the text, not the heading).
"""
from __future__ import annotations

import copy
import importlib.util
import os
import re
import time
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

from ..ai import claims as cl
from ..delivery import Finding, banner, decide_delivery, labelled_path, relabel_deck
from .session import SourceSession

REPO_ROOT = Path(__file__).resolve().parents[2]

#: layout findings that make a touched slide a problem (inspect_databook._REEXPORT_BLOCKING)
BLOCKING_LAYOUT = (("overflow", "OVERFLOW RISK"), ("placeholder_leak", "PLACEHOLDER LEAKED"),
                   ("table_overlap", "table overlaps"), ("lr_collision", "L/R collision"))


def _load_inspect_pptx():
    try:
        import inspect_pptx                                    # repo root on sys.path (CLI, app)
        return inspect_pptx
    except ImportError:
        spec = importlib.util.spec_from_file_location("inspect_pptx", REPO_ROOT / "inspect_pptx.py")
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module


def merged_results(session: SourceSession, accounts: List[str], first_generation: bool = False) -> Dict[str, Any]:
    """The base commentary with each chosen account replaced by its validated candidate; with first_generation, the candidates
    alone (no base text stands in for an account the agent did not finish)."""
    results = {} if first_generation else copy.deepcopy(session.base_results or {})
    for account in accounts:
        candidate = session.candidates[account]
        results[account] = cl.to_legacy_result(account, candidate.claims, candidate.outcome, flags=candidate.flags)
    if "__run_health__" not in results:
        results["__run_health__"] = {"zero_successful_calls": False, "accounts_on_fallback_bullet": [], "accounts_on_passthrough": [],
                                     "accounts_on_error_text": [], "accounts_without_prompt": [], "agent_path": True}
    return results


def slides_with_text(deck_path: str, fragments: List[str]) -> List[int]:
    """1-based slide numbers whose text holds any fragment, widened to the slides in
    between (an account that continues across pages has text on the pages between)."""
    from pptx import Presentation
    hits: List[int] = []
    for number, slide in enumerate(Presentation(deck_path).slides, 1):
        text = re.sub(r"\s+", "", "\n".join(shape.text_frame.text for shape in slide.shapes if getattr(shape, "has_text_frame", False)))
        if any(f and re.sub(r"\s+", "", f) in text for f in fragments):
            hits.append(number)
    return list(range(min(hits), max(hits) + 1)) if hits else []


def text_fragments(text: str, width: int = 14) -> List[str]:
    """Short pieces of a paragraph, spread through it, to find it again after packing cut it."""
    flat = re.sub(r"\s+", "", text)
    return [flat[i:i + width] for i in range(0, max(len(flat) - width, 1), width * 2)][:40]


def review_findings(session: SourceSession, accounts: Iterable[str]) -> List[Finding]:
    """The meaning review's state on each exported candidate (controller.py): an objection that stands blocks the deck; a review
    that was due and did not happen leaves it for a person. Accepting the draft does not clear either."""
    out: List[Finding] = []
    for account in accounts:
        candidate = session.candidates.get(account)
        review = (candidate.review if candidate is not None else {}) or {}
        for o in review.get("overruled") or []:
            out.append(Finding("VERIFIER_OVERRULED", "review", f"claim {o.get('claim')}: the lead model kept it against the verifier "
                               f"({', '.join(o.get('verifier') or [])}): {str(o.get('why') or '')[:160]}", account))
        if review.get("status") == "objected":
            why = "; ".join(o.get("why", "") for o in review.get("objections") or []) or "; ".join(m.get("why", "") for m in review.get("missing") or [])
            out.append(Finding("SEMANTIC_OBJECTION", "blocked", f"the reviewer's objection to this paragraph stands: {why[:200]}", account))
        elif review.get("status") == "unavailable":
            out.append(Finding("REVIEW_PENDING", "review", f"a review of this paragraph's meaning was due and did not happen: {str(review.get('why') or '')[:160]}", account))
        flags = [f for f in ((candidate.flags if candidate is not None else []) or [])
                 if f.get("kind") in ("deferred_note", "note_amount", "cause_not_proven", "note_wording_unchecked", "detail_caveat",
                                      "lead_left_to_person", "open_fix")]
        if flags:
            kinds: Dict[str, int] = {}
            for f in flags:
                kinds[f["kind"]] = kinds.get(f["kind"], 0) + 1
            out.append(Finding("HUMAN_CHECK", "review", "check against the source: " + ", ".join(f"{n} {k.replace('_', ' ')}" for k, n in kinds.items()),
                               account))
    return out


def session_delivery(session: SourceSession, results: Dict[str, Any], edited: Iterable[str] = (), first_generation: bool = False):
    """The delivery decision for results made against this session's sources: the run's own
    findings plus three it cannot see -- a Financials sheet that ties to almost nothing, a base run nothing ties to
    this workbook, and executive summaries written before the edit."""
    from .tools import financials_pairing
    pairing = financials_pairing(session)
    extra = [Finding("FINANCIALS_PAIRING_SUSPECT", "blocked",
                     f"only {pairing['agreeing_latest']} of {pairing['compared']} accounts tie to the Financials sheet on their latest period; "
                     "the sheet may belong to another entity or period")] if pairing["suspect"] else []
    edited = set(edited)
    extra += review_findings(session, [a for a in edited if a in results])
    if first_generation:
        # nothing from a base run is in this deck, so nothing about one applies; the summary band is the written overview
        # (each sentence verified against its account), or else the exporter's splice
        if not session.overview:
            extra.append(Finding("SUMMARY_SPLICED", "review", "the executive summary band is the exporter's splice of each account's first "
                                 "sentence; no summary was written for this first generation"))
        for account in session.contracts:
            if not str(account).startswith("__") and session.in_scope(account) and account not in results:
                status = ((session.candidates.get(account).review if session.candidates.get(account) else {}) or {}).get("status")
                if status == "objected":
                    extra.append(Finding("SEMANTIC_OBJECTION", "blocked", "not in the deck: the reviewer's objection to its draft stands", account))
                elif status == "unavailable":
                    extra.append(Finding("REVIEW_PENDING", "blocked", "not in the deck: its draft's review was due and did not happen", account))
                else:
                    extra.append(Finding("ACCOUNT_NOT_DRAFTED", "blocked", "in scope, but the agent has no accepted paragraph for it", account))
        return decide_delivery(selection=session.selection, ai_results=results, bs_is_results=session.bs_is_results, extra_findings=extra)
    check = session.base_run_check or {}
    if check.get("status") == "mismatch":
        for reason in (check.get("reasons") or [])[:4]:
            extra.append(Finding("BASE_RUN_MISMATCH", "blocked", f"the base run does not describe this workbook ({reason}); the commentary "
                                 "of accounts you did not edit is out of date"))
    if check.get("status") in ("mismatch", "edited_since"):
        for account in check["accounts"].get("changed") or []:
            if account not in edited:
                extra.append(Finding("BASE_ACCOUNT_CHANGED", "blocked", "the workbook's data for this account differs from what its commentary was written from", account))
    if check.get("status") == "unverified":
        acc = check["accounts"]
        extra.append(Finding("BASE_RUN_UNVERIFIED", "review",
                             f"nothing proves the base run was written from this workbook (no identity file; {acc['matched']} of {acc['in_run']} accounts have "
                             "a matching data fingerprint): the commentary of accounts you did not edit is carried over as it is"))
    if session.base_results:
        summaries = "the saved executive summaries" if session.base_summaries else "the exporter's splice of each account's first sentence (the base run saved no summary)"
        extra.append(Finding("SUMMARY_NOT_REGENERATED", "info" if session.base_summaries else "review",
                             f"the executive summary band is {summaries}, written before this edit: read it against the edited paragraph"))
    return decide_delivery(selection=session.selection, ai_results=results, bs_is_results=session.bs_is_results, extra_findings=extra)


def layout_problems(deck_path: str, slides: List[int]) -> Dict[str, Any]:
    """inspect_pptx on a deck: the blocking findings on the given slides, and how many warnings sit elsewhere."""
    inspector = _load_inspect_pptx()
    layout = inspector.inspect_pptx(deck_path, inspector._load_config(None), quiet=True)
    problems: List[str] = []
    for detail in layout.get("warning_details") or []:
        m = re.search(r"Slide (\d+)", str(detail))
        if m and int(m.group(1)) in slides and any(needle.lower() in str(detail).lower() for _k, needle in BLOCKING_LAYOUT):
            problems.append(str(detail))
    return {"problems": problems, "warnings_elsewhere": max(len(layout.get("warning_details") or []) - len(problems), 0),
            "content_slides": layout.get("content_slides"), "total_slides": layout.get("total_slides")}


def export_session_deck(session: SourceSession, accounts: List[str], out_path: str = "", first_generation: bool = False) -> Dict[str, Any]:
    from ..pptx import build_pptx_structured_payloads, export_pptx_from_structured_data_combined
    from ..pptx.helpers import _rendered_bullet_label
    from .tools import ToolError

    template = REPO_ROOT / "fdd_utils" / "template.pptx"
    if not template.exists():
        raise ToolError("NO_TEMPLATE", f"template not found: {template}")
    results = merged_results(session, accounts, first_generation=first_generation)
    delivery = session_delivery(session, results, edited=accounts, first_generation=first_generation)
    payloads = build_pptx_structured_payloads(results, session.mappings, bs_is_results=session.bs_is_results, dfs=session.dfs)
    stem = re.sub(r"[^\w\-]", "_", session.entity_name or Path(session.workbook_path).stem or "deck")
    target = out_path or str(REPO_ROOT / "fdd_utils" / "output" / f"{stem}_agent_{time.strftime('%Y%m%d_%H%M%S')}.pptx")
    Path(target).parent.mkdir(parents=True, exist_ok=True)
    target = labelled_path(target, delivery)
    is_chinese = session.language in ("Chi", "Chn", "chinese", "Chinese")
    export_pptx_from_structured_data_combined(
        str(template), payloads["BS"], payloads["IS"], target,
        project_name=session.entity_name or Path(session.workbook_path).stem or None,
        language="chinese" if is_chinese else "english",
        temp_path=session.financials_from or session.workbook_path, selected_sheet=session.financials_sheet,
        is_chinese_databook=is_chinese, bs_is_results=session.bs_is_results, model_type=None,
        skip_summary_ai=True, pre_generated_summaries=((session.overview or None) if first_generation else session.base_summaries or None),
        mappings=session.mappings,
        financials_from=session.financials_from, delivery=delivery)

    touched: Dict[str, List[int]] = {}
    for account in accounts:
        fragments = text_fragments(cl.render_claims(session.candidates[account].claims))
        labels = [_rendered_bullet_label(item, is_chinese) for statement in payloads.values() for item in statement
                  if str(item.get("account_name") or item.get("mapping_key") or "") == account]
        touched[account] = slides_with_text(target, fragments)
        if not touched[account]:
            delivery.add("ACCOUNT_NOT_IN_DECK", "blocked", "the edited commentary was not found on any slide", account)

    touched_slides = sorted({n for slides in touched.values() for n in slides})
    checked = layout_problems(target, touched_slides)
    problems, layout = checked["problems"], checked
    for problem in problems:
        delivery.add("LAYOUT_PROBLEM", "review", problem[:200])
    # Which paragraph of which slide says what, from which facts: the manifest a returned deck is read against.
    from ..pptx import lineage
    public, internal = lineage.build_manifest(
        target, [item for statement in payloads.values() for item in statement], is_chinese=is_chinese,
        source_version=session.source_version, contracts=session.contracts,
        claims_by_account={a: session.candidates[a].claims for a in accounts}, binding=session.binding.to_dict())
    lineage.stamp_deck(target, public, internal)
    # What the post-export checks found can change the label the file was named with.
    target = relabel_deck(target, delivery)
    record = {"path": target, "delivery": delivery.to_dict(), "touched_slides": touched, "touched_slide_numbers": touched_slides,
              "layout": {"content_slides": layout.get("content_slides"), "total_slides": layout.get("total_slides"),
                         "problems_on_touched_slides": problems, "warnings_elsewhere": layout["warnings_elsewhere"]},
              "manifest": {"elements": len(public["elements"]), "deck_id": public["deck_id"]},
              "source_version": session.source_version}
    # "ok" is what a later gate may rely on: the deck was made, nothing blocks it, no touched slide has a
    # layout problem. The digests bind the record to the exact candidates it exported.
    record["ok"] = not problems and delivery.status != "blocked"
    record["candidates"] = {a: session.candidates[a].claims_digest for a in accounts}
    session.exports.append(record)
    if problems:
        raise ToolError("LAYOUT_PROBLEM", f"{len(problems)} layout problem(s) on the touched slides {touched_slides}", record)
    if delivery.status == "blocked":
        raise ToolError("DELIVERY_BLOCKED", "the deck was exported as a draft and is not deliverable: " + "; ".join(delivery.reasons(limit=3)), record)
    return record
