"""Databook inspection tool — run this against each real client databook
(Windows-side "inputs" folder) BEFORE trusting the pipeline's output on it.

Built to answer these specific questions from a real-databook QA pass:
  1. Are financial + breakdown tabs actually being read correctly?
  2. Does reconciliation match up — and for anything that doesn't, which
     tab names actually exist so a mapping can be added?
  3. Does the databook have indented / total-then-breakdown row structures
     that a human eye catches but extraction might misread?
  4. Are '000-style unit markers (CNY'000 / 人民币千元) being detected
     correctly per tab, or could a tab silently fall back to a 1x
     multiplier when it should be 1000x?

Everything above is DETERMINISTIC — no AI calls, fast, safe to run against
any real client file without touching API budgets or the network.

Usage:
    python inspect_databook.py "inputs/SomeClient.databook.xlsx"
    python inspect_databook.py "inputs/"                    # scans every .xlsx in the folder,
                                                              # prints a final aggregate summary table
    python inspect_databook.py "inputs/SomeClient.databook.xlsx" --sheet Financials
    python inspect_databook.py "inputs/" --run-ai --model workbench   # adds AI-dependent checks (see below)
    python inspect_databook.py "inputs/" --run-ai --model workbench --limit 5
        # fast smoke test: only the first 5 mapped accounts per file go to AI, across
        # every file in the folder — use this to sanity-check the whole batch before
        # committing to a full (potentially hours-long) unlimited run

AI-dependent checks (only with --run-ai; needs a configured provider in
fdd_utils/config.yml, costs real tokens/time — run once per databook when
ready to review quality, not on every edit). Only accounts that passed
reconciliation with an included Match status go to AI — exactly mirroring
production (fdd_utils/ui.py:derive_reconciliation_matched_keys) — so this
never wastes tokens on tabs the real pipeline would never send, and the
timing numbers are directly comparable across databooks/models. A tqdm
progress bar shows live stage/account progress during the run:
  5. Runs the full AI pipeline once and reports wall-clock time and
     per-agent retry counts (compare qwen local vs GPT-5.5 workbench by
     re-running with --model).
  6. Numeric grounding sweep: extracts every number mentioned in generated
     commentary and cross-checks it against the ground-truth dfs value at
     1x, 1000x, 0.001x, 10000x, and 0.0001x. Flags any number that ONLY
     matches at a non-1x scale — this is the generalized version of the
     1000x-scaling bug fixed this session (fdd_utils/pptx.py
     embed_financial_tables mutating bs_is_results in place). If a
     different databook layout reintroduces a scaling bug anywhere in the
     pipeline, this check catches it automatically instead of relying on
     a human noticing a wrong number in a slide.
  7. Chinese unit-label sanity check: for Chinese output, flags any
     万元/亿元 amount that doesn't reconcile to the underlying actual-CNY
     ground truth within tolerance — the 千元→萬元 mislabeling class of bug.

Output is plain text to stdout — copy/paste it back for review; nothing is
written back to the workbook.
"""
from __future__ import annotations

import argparse
import re
import sys
import threading
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
from tqdm import tqdm

from fdd_utils.workbook import (
    _unit_markers,
    extract_data_from_excel,
    extract_balance_sheet_and_income_statement,
    find_account_in_dfs,
    get_effective_mappings,
    load_mappings,
    reconcile_financial_statements,
)
from fdd_utils.ui import derive_reconciliation_matched_keys
from fdd_utils.financial_common import get_pipeline_result_text

pd.set_option("display.width", 200)


def _hr(title: str = "") -> None:
    print("\n" + "=" * 78)
    if title:
        print(f"  {title}")
        print("=" * 78)


# ---------------------------------------------------------------------------
# 1. Tab read summary
# ---------------------------------------------------------------------------

def check_tab_read_summary(databook_path: str, entity_name: str = "") -> Dict[str, pd.DataFrame]:
    _hr("1. TAB READ SUMMARY")
    xl = pd.ExcelFile(databook_path)
    print(f"Sheets found in workbook: {xl.sheet_names}")

    dfs, workbook_list, overall_result_type, language, resolution = extract_data_from_excel(
        databook_path=databook_path, entity_name=entity_name, mode="All",
        return_resolution=True,
    )
    print(f"\nDetected language: {language}")
    print(f"Tabs successfully parsed into dfs: {len(dfs)} of {len(xl.sheet_names)} sheets")
    dfs_keys_normalized = {str(k).strip() for k in dfs.keys()}
    unparsed = [
        s for s in xl.sheet_names
        if s.strip() not in dfs_keys_normalized and s.strip().lower() != "financials"
    ]
    if unparsed:
        print(f"⚠️  Sheets NOT parsed into any account tab (check if these should map): {unparsed}")

    for key, df in dfs.items():
        if df is None or df.empty:
            print(f"  ⚠️  {key}: EMPTY dataframe")
            continue
        date_cols = [c for c in df.columns if not str(c).endswith("_formatted") and not str(c).startswith("__")]
        print(f"  {key}: {len(df)} rows, columns={date_cols[:8]}{'...' if len(date_cols) > 8 else ''}")
    return dfs


# ---------------------------------------------------------------------------
# 2. Unit-marker sanity check (root cause class of the 1000x bug)
# ---------------------------------------------------------------------------

def check_unit_markers(databook_path: str, dfs: Dict[str, pd.DataFrame]) -> None:
    _hr("2. UNIT-MARKER SANITY CHECK (CNY'000 / 人民币千元 detection per tab)")
    print(
        "Each tab is scanned independently for a thousands-unit marker in its\n"
        "first 8 rows. If a tab's marker ISN'T found here, its multiplier\n"
        "silently falls back to 1x instead of 1000x — this is the failure mode\n"
        "that produces numbers 1000x too small for THAT tab specifically.\n"
        "Only tabs that actually parsed into an account (see section 1) are\n"
        "flagged as missing — navigation/cover/TB/pivot tabs are expected to\n"
        "have no unit marker and are listed but not counted as a problem.\n"
    )
    relevant_tabs = {str(k).strip() for k in dfs.keys()}
    xl = pd.ExcelFile(databook_path)
    any_missing = False
    missing_relevant: List[str] = []
    for sheet in xl.sheet_names:
        try:
            df = xl.parse(sheet, header=None, nrows=12)
        except Exception as exc:
            print(f"  ⚠️  {sheet}: could not read ({exc})")
            continue
        markers_8 = _unit_markers(df.head(8))
        markers_12 = _unit_markers(df)
        is_relevant = sheet.strip() in relevant_tabs or sheet.strip().lower() == "financials"
        if markers_8:
            status = "✅"
        elif markers_12:
            status = "⚠️ found beyond row 8"
        elif is_relevant:
            status = "❌ NOT FOUND"
            any_missing = True
            missing_relevant.append(sheet)
        else:
            status = "·  (not a parsed account tab, skipped)"
        print(f"  {status}  {sheet}: markers(first 8 rows)={markers_8}  markers(first 12 rows)={markers_12}")
    if any_missing:
        print(
            f"\n⚠️  {len(missing_relevant)} PARSED account tab(s) have no unit marker in the\n"
            f"   scanned window: {missing_relevant}\n"
            "   If the databook actually reports in thousands, those tabs' values will\n"
            "   be under-scaled by 1000x relative to tabs where the marker IS found.\n"
            "   Fix: either the tab needs its own 'CNY'000' header repeated, or\n"
            "   _unit_markers()'s max_rows=8 scan window needs widening for this\n"
            "   databook's layout (paste this output back for a targeted fix)."
        )
    else:
        print("\n✅ All parsed account tabs declare a unit marker within the scan window.")


# ---------------------------------------------------------------------------
# 3. Row-structure sanity check (indentation / total-then-breakdown)
# ---------------------------------------------------------------------------

def check_row_structures(dfs: Dict[str, pd.DataFrame]) -> None:
    _hr("3. ROW-STRUCTURE SANITY CHECK (indentation / total-then-breakdown)")
    print(
        "Reports how each tab's rows were auto-classified (breakdown / subtotal\n"
        "/ total / plain). Skim each tab's list against the real Excel layout —\n"
        "if a row that's visually a breakdown item got classified as a total (or\n"
        "vice versa), that's exactly the class of layout the auto-detector can\n"
        "misread. Only tabs with a NON-trivial mix of types are printed in full;\n"
        "tabs with all-plain rows are summarized in one line.\n"
    )
    for key, df in dfs.items():
        if df is None or df.empty:
            continue
        row_types = df.attrs.get("row_types_by_description") or {}
        desc_col = df.columns[0]
        descriptions = [str(v) for v in df[desc_col].tolist()]
        if not row_types or all(row_types.get(d, "plain") == "plain" for d in descriptions):
            print(f"  {key}: {len(descriptions)} rows, all classified 'plain' (no total/subtotal/breakdown detected)")
            continue
        print(f"  {key}:")
        for desc in descriptions:
            rtype = row_types.get(desc, "plain")
            marker = {"total": "Σ TOTAL", "subtotal": "Σ subtotal", "breakdown": "  breakdown"}.get(rtype, "  plain")
            print(f"      [{marker:12s}] {desc}")


# ---------------------------------------------------------------------------
# 4. Reconciliation summary with actionable tab-name listing
# ---------------------------------------------------------------------------

def check_reconciliation(
    databook_path: str, sheet_name: str, dfs: Dict[str, pd.DataFrame], entity_name: str = ""
) -> Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame]]:
    _hr(f"4. RECONCILIATION SUMMARY — sheet: {sheet_name}")
    try:
        bs_is_results = extract_balance_sheet_and_income_statement(
            workbook_path=databook_path, sheet_name=sheet_name, debug=False,
        )
    except Exception as exc:
        print(f"❌ Could not extract Financials sheet '{sheet_name}': {exc}")
        return None, None

    mappings = get_effective_mappings(load_mappings(), None)
    bs_recon, is_recon = reconcile_financial_statements(
        bs_is_results=bs_is_results, dfs=dfs, mappings=mappings,
        tolerance=1.0, materiality_threshold=0.005, debug=False,
    )

    all_tab_names = sorted(dfs.keys())
    for label, recon in (("Balance Sheet", bs_recon), ("Income Statement", is_recon)):
        if recon is None or recon.empty:
            print(f"\n{label}: no rows")
            continue
        print(f"\n{label}: {len(recon)} lines")
        print(recon["Match"].value_counts().to_string())
        problems = recon[recon["Match"].astype(str).str.contains("Not Found|Diff", na=False)]
        if not problems.empty:
            print(f"\n  Unmatched/diff lines for {label} (account -> mapping status):")
            for _, r in problems.iterrows():
                print(f"    - {r['Financials_Account']!r}: {r['Mapping_Status']} | {r['Mapping_Note']}")
            print(f"\n  All tab names actually present in this workbook (for adding to mappings.yml):")
            print(f"    {all_tab_names}")

    return bs_recon, is_recon


# ---------------------------------------------------------------------------
# 5-7. AI-dependent checks
# ---------------------------------------------------------------------------

_NUMBER_RE = re.compile(
    # Currency prefix is REQUIRED, not optional — the project's enforced
    # number format (prompts.yml) always prefixes real amounts with
    # CNY/人民币. Without this, a bare number match also catches date years
    # ("31 January 2026" -> 2026), footnote refs ("Note 1"), and other
    # non-monetary digits, which produced false "1000x too small" flags
    # (e.g. 2026 * 1000 landing within tolerance of an unrelated real balance).
    r"(?:CNY|RMB|USD|HKD|US\$|\$|人民币|人民幣)\s*"
    r"(\d[\d,]*(?:\.\d+)?)\s*"
    r"(million|mn|万|萬|亿|億|K)?",
    re.IGNORECASE,
)

_SCALE_MAP = {
    None: 1, "": 1, "k": 1_000,
    "million": 1_000_000, "mn": 1_000_000,
    "万": 10_000, "萬": 10_000,
    "亿": 100_000_000, "億": 100_000_000,
}


def _extract_numbers_with_scale(text: str) -> List[Tuple[float, int, int]]:
    """Returns (value, match_start, match_end) so callers can pull the
    surrounding sentence for a flagged number — the isolated number alone
    ("40,378,000.00") isn't enough to tell whether it's a real hallucinated
    scale bug or something benign; the actual written clause is."""
    values: List[Tuple[float, int, int]] = []
    for match in _NUMBER_RE.finditer(text or ""):
        raw, suffix = match.group(1), (match.group(2) or "").lower()
        if not raw or raw in (",", "."):
            continue
        try:
            base = float(raw.replace(",", ""))
        except ValueError:
            continue
        scale = _SCALE_MAP.get(suffix, 1)
        values.append((base * scale, match.start(), match.end()))
    return values


def _context_snippet(text: str, start: int, end: int, radius: int = 50) -> str:
    lo = max(0, start - radius)
    hi = min(len(text), end + radius)
    prefix = "…" if lo > 0 else ""
    suffix = "…" if hi < len(text) else ""
    return f"{prefix}{text[lo:hi].strip()}{suffix}"


def _ground_truth_values(dfs: Dict[str, pd.DataFrame]) -> List[float]:
    truth = []
    for df in dfs.values():
        if df is None or df.empty:
            continue
        for col in df.columns:
            if str(col).endswith("_formatted") or str(col).startswith("__"):
                continue
            if pd.api.types.is_numeric_dtype(df[col]):
                truth.extend(float(v) for v in df[col].dropna().tolist() if v != 0)
    return truth


def check_numeric_grounding(mapping_key: str, generated_text: str, dfs: Dict[str, pd.DataFrame]) -> List[str]:
    """Returns a list of warning strings for numbers that only match ground
    truth at a non-1x scale factor (the generalized 1000x-bug detector)."""
    truth_values = _ground_truth_values(dfs)
    if not truth_values:
        return []
    warnings = []
    for value, start, end in _extract_numbers_with_scale(generated_text):
        if value == 0:
            continue
        matches_1x = any(abs(value - t) / max(abs(t), 1) < 0.06 for t in truth_values)
        if matches_1x:
            continue
        for factor, label in ((1000, "1000x too small"), (0.001, "1000x too large"),
                               (10000, "10000x too small"), (0.0001, "10000x too large")):
            scaled = value * factor
            if any(abs(scaled - t) / max(abs(t), 1) < 0.06 for t in truth_values):
                snippet = _context_snippet(generated_text, start, end)
                warnings.append(
                    f"  🔴 [{mapping_key}] number {value:,.2f} in generated text only matches "
                    f"ground truth when scaled — looks {label} (matches {scaled:,.2f})\n"
                    f"      context: \"{snippet}\""
                )
                break
    return warnings


def run_ai_checks(
    databook_path: str, sheet_name: str, dfs: Dict[str, pd.DataFrame],
    entity_name: str, model_type: str, model_name: Optional[str], language: str,
    bs_recon: Optional[pd.DataFrame], is_recon: Optional[pd.DataFrame],
    limit: Optional[int] = None, workers: Optional[int] = None,
) -> Dict[str, Any]:
    _hr("5-7. AI-DEPENDENT CHECKS (running full pipeline once — this costs real tokens/time)")
    from fdd_utils.ai import run_ai_pipeline_with_progress, SUBAGENT_SEQUENCE

    # Mirror production exactly: only accounts that passed reconciliation with
    # an included Match status go to AI (fdd_utils/ui.py:derive_reconciliation_matched_keys).
    # BS excludes ❌ Diff (needs human review first); IS includes it (IS recon
    # is inherently noisier). Running every dfs tab through the LLM — including
    # ones the real pipeline would never send — wastes tokens/time and makes
    # the qwen-vs-GPT-5.5 timing comparison meaningless.
    mapping_keys = derive_reconciliation_matched_keys((bs_recon, is_recon), dfs.keys(), None)
    total_mapped = len(mapping_keys)
    if not mapping_keys:
        print(
            "❌ No accounts passed reconciliation with an included Match status "
            "(✅ Match / ⚠️ Match / ✅ Immaterial, or ❌ Diff for IS) — nothing "
            "would be sent to AI in production either. Skipping AI-dependent checks."
        )
        return {"skipped": True, "reason": "no mapped accounts"}
    if limit and limit > 0 and len(mapping_keys) > limit:
        print(f"--limit {limit}: sampling {limit} of {total_mapped} mapped accounts "
              f"(fast smoke-test mode, not a full run).")
        mapping_keys = mapping_keys[:limit]
    # Mirror run_agent_stage's ACTUAL resolution chain so this print is
    # trustworthy — it previously ignored config.yml entirely and always
    # printed "built-in default" even when <provider>.max_workers was set,
    # which was actively misleading (looked like the config value wasn't
    # being picked up when it may well have been).
    if workers:
        effective_workers, workers_source = workers, "--workers override"
    else:
        from fdd_utils.ai import load_yaml_config, get_default_config_path
        try:
            _provider_cfg = load_yaml_config(get_default_config_path()).get(model_type, {}) or {}
        except Exception:
            _provider_cfg = {}
        _configured = _provider_cfg.get("max_workers")
        if _configured:
            effective_workers, workers_source = int(_configured), f"{model_type}.max_workers in config.yml"
        else:
            effective_workers = 4 if model_type == "local" else 2
            workers_source = "built-in default — no <provider>.max_workers set in config.yml"
    print(f"Running pipeline for {len(mapping_keys)} MAPPED accounts (of {len(dfs)} total tabs), "
          f"model_type={model_type}, model_name={model_name}, workers={effective_workers} ({workers_source})...")

    # SUBAGENT_SEQUENCE is the ACTUAL active pipeline (Generator, Auditor,
    # Validator — Refiner is dormant, see ai.py:SUBAGENT_ALIASES comment).
    # Hardcoding 4 here made the bar always stall at 75% (60/80) since only
    # 3 stages ever fire a progress callback.
    stage_count = len(SUBAGENT_SEQUENCE)
    total_steps = stage_count * len(mapping_keys)
    pbar = tqdm(total=total_steps, desc="AI pipeline", unit="step")
    seen_step = {"n": 0}
    progress_lock = threading.Lock()

    # Per-stage wall-clock breakdown — lets you SEE whether a per-agent
    # reasoning_effort override (e.g. Auditor set to "low") actually reduced
    # that stage's time, not just eyeball the total. A stage is "done" the
    # moment its last account reports completed == total_eligible; the gap
    # since the previous stage's completion is that stage's duration.
    stage_timing: Dict[str, float] = {}
    stage_clock = {"last_boundary": time.time(), "seen_labels": set()}

    def _tqdm_progress(agent_num, agent_label, completed, total_eligible, overall_step, mapping_key):
        with progress_lock:
            pbar.set_postfix_str(f"{agent_label}: {mapping_key}"[:60])
            delta = overall_step - seen_step["n"]
            if delta > 0:
                pbar.update(delta)
                seen_step["n"] = overall_step
            if agent_label not in stage_clock["seen_labels"] and completed == total_eligible and total_eligible > 0:
                now = time.time()
                stage_timing[agent_label] = now - stage_clock["last_boundary"]
                stage_clock["last_boundary"] = now
                stage_clock["seen_labels"].add(agent_label)

    # A single account's API call can take 20-60s+ with nothing to report in
    # between (no discrete progress event fires mid-call), which makes the
    # bar look frozen even though it's working. Refresh the display every
    # second regardless — this ticks tqdm's elapsed-time counter so it's
    # visibly alive, without needing an actual progress-count change.
    stop_refresh = threading.Event()

    def _tick_refresh():
        while not stop_refresh.wait(1.0):
            with progress_lock:
                pbar.refresh()

    refresh_thread = threading.Thread(target=_tick_refresh, daemon=True)
    refresh_thread.start()

    start = time.time()
    stage_clock["last_boundary"] = start
    try:
        results = run_ai_pipeline_with_progress(
            mapping_keys=mapping_keys, dfs=dfs, model_type=model_type, model_name=model_name,
            language=language, use_multithreading=True, max_workers=workers,
            progress_callback=_tqdm_progress,
        )
    finally:
        stop_refresh.set()
        refresh_thread.join(timeout=2.0)
        pbar.close()
    elapsed = time.time() - start
    print(f"\n5. TIMING: full pipeline for {len(mapping_keys)} mapped accounts took {elapsed:.1f}s "
          f"({elapsed / max(len(mapping_keys), 1):.1f}s/account) on {model_type}"
          f"{'/' + model_name if model_name else ''}")
    if stage_timing:
        print("   Per-stage breakdown (wall-clock, all accounts in that stage together):")
        for stage_label, stage_seconds in stage_timing.items():
            per_account = stage_seconds / max(len(mapping_keys), 1)
            print(f"     {stage_label:12s}: {stage_seconds:7.1f}s total, {per_account:6.1f}s/account")

    _hr("6-7. NUMERIC GROUNDING + UNIT-LABEL SWEEP")
    all_warnings: List[str] = []
    checked_count = 0
    empty_accounts: List[str] = []
    for key, content in (results or {}).items():
        # "final" is the pipeline's actual output field (fdd_utils/ai.py
        # get_pipeline_result_text) — NOT "final_content", which only exists
        # inside the raw agent_4_validation sub-dict, not at this top level.
        text = get_pipeline_result_text(content)
        if not text:
            empty_accounts.append(key)
            continue
        checked_count += 1
        all_warnings.extend(check_numeric_grounding(key, text, {key: dfs.get(key)} if key in dfs else dfs))
    print(f"Checked {checked_count} of {len(results or {})} account(s) with non-empty output.")
    if empty_accounts:
        print(
            f"⚠️  {len(empty_accounts)} account(s) came back with EMPTY final output — this is a\n"
            f"   real missing-commentary problem, not a test artifact (these would render as\n"
            f"   blank/missing bullets in the actual PPTX). Likely cause for a reasoning model:\n"
            f"   the account's max_tokens budget was consumed entirely by hidden reasoning\n"
            f"   tokens with nothing left for the visible answer (see agents.*.max_tokens and\n"
            f"   workbench.reasoning_effort in config.yml).\n"
            f"   Empty accounts: {empty_accounts}"
        )
    if all_warnings:
        print(f"Found {len(all_warnings)} suspicious number(s):")
        for w in all_warnings:
            print(w)
    else:
        print("✅ No numbers found that only match ground truth at a non-1x scale factor.")

    return {
        "skipped": False,
        "total_mapped": total_mapped,
        "ran": len(mapping_keys),
        "elapsed": elapsed,
        "empty_accounts": empty_accounts,
        "grounding_warnings": len(all_warnings),
        "stage_timing": stage_timing,
    }


# ---------------------------------------------------------------------------

def _resolve_financials_sheets(xl: pd.ExcelFile) -> List[str]:
    """Exact 'Financials' match wins. Otherwise look for a multi-entity
    pattern like 'Financials - NB' / 'Financials - HN' (no single combined
    Financials sheet — common when one workbook covers several properties).
    Falling back to sheet_names[0] (e.g. a 'Briefing'/'Cover' tab) silently
    reconciles against the wrong sheet and reports an empty/misleading
    result, so that fallback is deliberately NOT used here.
    """
    exact = [s for s in xl.sheet_names if s.strip().lower() == "financials"]
    if exact:
        return exact
    prefixed = [s for s in xl.sheet_names if re.match(r"^financials\s*-", s.strip(), re.IGNORECASE)]
    return prefixed


def inspect_one(path: str, sheet: Optional[str], entity_name: str, run_ai: bool,
                 model_type: str, model_name: Optional[str], limit: Optional[int] = None,
                 workers: Optional[int] = None) -> Dict[str, Any]:
    _hr(f"INSPECTING: {path}")
    summary: Dict[str, Any] = {"file": Path(path).name, "status": "ok"}
    dfs = check_tab_read_summary(path, entity_name=entity_name)
    summary["tabs_parsed"] = len(dfs)
    check_unit_markers(path, dfs)
    check_row_structures(dfs)

    xl = pd.ExcelFile(path)
    summary["total_sheets"] = len(xl.sheet_names)
    if sheet:
        sheet_names = [sheet]
    else:
        sheet_names = _resolve_financials_sheets(xl)
        if not sheet_names:
            print(
                f"\n❌ No sheet named exactly 'Financials' or matching 'Financials - <entity>' "
                f"found in {path}. Skipping reconciliation — pass --sheet explicitly.\n"
                f"   Sheet names in this workbook: {xl.sheet_names}"
            )
            summary["status"] = "no Financials sheet"
            return summary
        if len(sheet_names) > 1:
            print(
                f"\nℹ️  Multi-entity workbook: found {len(sheet_names)} Financials sheets "
                f"{sheet_names} — running reconciliation against each."
            )

    bs_recon_parts: List[pd.DataFrame] = []
    is_recon_parts: List[pd.DataFrame] = []
    for sheet_name in sheet_names:
        bs_recon, is_recon = check_reconciliation(path, sheet_name, dfs, entity_name=entity_name)
        if bs_recon is not None and not bs_recon.empty:
            bs_recon_parts.append(bs_recon)
        if is_recon is not None and not is_recon.empty:
            is_recon_parts.append(is_recon)

    combined_bs_recon = pd.concat(bs_recon_parts, ignore_index=True) if bs_recon_parts else None
    combined_is_recon = pd.concat(is_recon_parts, ignore_index=True) if is_recon_parts else None
    if combined_bs_recon is not None:
        summary["bs_match"] = combined_bs_recon["Match"].value_counts().to_dict()
    if combined_is_recon is not None:
        summary["is_match"] = combined_is_recon["Match"].value_counts().to_dict()

    if run_ai:
        language = "Eng"
        try:
            from fdd_utils.workbook import process_workbook_data
            state = process_workbook_data(temp_path=path, entity_name=entity_name,
                                           selected_sheet=sheet_names[0], debug=False)
            language = state.get("language", "Eng")
        except Exception:
            pass
        ai_summary = run_ai_checks(
            path, sheet_names[0], dfs, entity_name, model_type, model_name, language,
            combined_bs_recon, combined_is_recon, limit=limit, workers=workers,
        )
        summary["ai"] = ai_summary
    return summary


def _print_final_summary(summaries: List[Dict[str, Any]], run_ai: bool) -> None:
    _hr("FINAL SUMMARY — all files")
    rows = []
    for s in summaries:
        if s.get("status") != "ok":
            rows.append({
                "file": s.get("file", "?"), "tabs": "-", "BS": "-", "IS": "-",
                "AI": s.get("status", "error"),
            })
            continue
        bs_match = s.get("bs_match") or {}
        is_match = s.get("is_match") or {}
        bs_str = ", ".join(f"{k}={v}" for k, v in bs_match.items()) or "-"
        is_str = ", ".join(f"{k}={v}" for k, v in is_match.items()) or "-"
        ai_str = "-"
        if run_ai:
            ai = s.get("ai") or {}
            if ai.get("skipped"):
                ai_str = f"skipped ({ai.get('reason', '?')})"
            elif ai:
                empty_n = len(ai.get("empty_accounts") or [])
                ai_str = (
                    f"{ai.get('ran', 0)}/{ai.get('total_mapped', 0)} ran, "
                    f"{ai.get('elapsed', 0):.0f}s, "
                    f"{empty_n} empty, {ai.get('grounding_warnings', 0)} 🔴"
                )
        stage_timing = ((s.get("ai") or {}).get("stage_timing")) or {}
        stage_str = (
            ", ".join(f"{label}={secs / max((s.get('ai') or {}).get('ran', 1), 1):.1f}s/acct"
                      for label, secs in stage_timing.items())
            if stage_timing else "-"
        )
        rows.append({
            "file": s.get("file", "?"),
            "tabs": f"{s.get('tabs_parsed', '?')}/{s.get('total_sheets', '?')}",
            "BS": bs_str, "IS": is_str, "AI": ai_str, "stage s/acct": stage_str,
        })
    df = pd.DataFrame(rows)
    print(df.to_string(index=False))
    if run_ai:
        total_empty = sum(len((s.get("ai") or {}).get("empty_accounts") or []) for s in summaries)
        total_warnings = sum((s.get("ai") or {}).get("grounding_warnings", 0) for s in summaries)
        print(f"\nAcross all files: {total_empty} empty account(s) total, {total_warnings} grounding warning(s) total.")
        if total_empty == 0 and total_warnings == 0:
            print("✅ Nothing flagged across the whole batch.")


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("path", help="databook .xlsx file, or a folder to scan every .xlsx in it")
    ap.add_argument("--sheet", default=None, help="Financials sheet name (default: auto-detect 'Financials')")
    ap.add_argument("--entity", default="", help="entity name filter, if the workbook has multiple entities")
    ap.add_argument("--run-ai", action="store_true", help="also run checks 5-7 (real AI calls, costs time/tokens)")
    ap.add_argument("--model", default="local", help="model_type for --run-ai: local | workbench | deepseek | openai")
    ap.add_argument("--model-name", default=None, help="specific model id within the provider (e.g. GPT-5.5's id)")
    ap.add_argument("--limit", type=int, default=None,
                    help="with --run-ai, cap how many mapped accounts per file go to AI "
                         "(fast smoke-test across many files; omit for a full run)")
    ap.add_argument("--workers", type=int, default=None,
                    help="with --run-ai, concurrent worker threads per pipeline stage. "
                         "Built-in default is 4 for local, 2 for everything else (rate-limit "
                         "caution) — override here to test a higher value, e.g. --workers 4 "
                         "on workbench if you've already validated the gateway handles it.")
    args = ap.parse_args()

    target = Path(args.path)
    if target.is_dir():
        files = sorted(target.glob("*.xlsx"))
        files = [f for f in files if not f.name.startswith("~$")]
        if not files:
            print(f"No .xlsx files found in {target}")
            return 1
        print(f"Found {len(files)} databook(s) in {target}: {[f.name for f in files]}")
        if args.run_ai and not args.limit:
            print(
                f"⚠️  --run-ai with no --limit on {len(files)} files will run EVERY mapped "
                f"account in EVERY file — based on inputs/昆山.xlsx (20 accounts, ~15-20 min\n"
                f"   on workbench), a full batch across all files could take a long time. "
                f"Consider --limit 5 for a first pass across everything, then a full,\n"
                f"   unlimited run on just the file(s) that need a closer look."
            )
    else:
        files = [target]

    summaries: List[Dict[str, Any]] = []
    for f in files:
        try:
            summary = inspect_one(str(f), args.sheet, args.entity, args.run_ai, args.model,
                                   args.model_name, limit=args.limit, workers=args.workers)
            summaries.append(summary)
        except Exception as exc:
            print(f"\n❌ FAILED inspecting {f}: {type(exc).__name__}: {exc}")
            import traceback
            traceback.print_exc()
            summaries.append({"file": f.name, "status": f"FAILED: {exc}"})

    if len(files) > 1:
        _print_final_summary(summaries, args.run_ai)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
