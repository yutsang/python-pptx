#!/usr/bin/env python3
"""List every N2 relationship candidate and its test result, for one databook.

Free and offline: no LLM call, no network, seconds per file. This is the
acceptance harness for milestone N2 -- it prints every candidate edge with the
numbers behind it, so a pass or a rejection can be traced back to cells in the
workbook by hand, and it re-renders every prompt with and without the fact
table to show what actually reaches the model and what it costs.

    python ad-hoc/workbench/report_cross_account_facts.py <databook.xlsx>
        [--sheet Financials] [--language Chi|Eng] [--quiet]

Three things it is there to prove, in this order:

  1. every candidate relationship, with its test and its evidence;
  2. that no REJECTED relationship's wording reaches any rendered prompt --
     the rule the whole milestone rests on;
  3. that adding the fact table breaks no prompt substitution. _safe_format
     swallows a KeyError and returns the template UNSUBSTITUTED, so a
     placeholder without a matching format_params key silently kills every
     substitution in that prompt. The before/after placeholder counts here are
     the check that catches it.

Never hardcode a databook path in this file: the repo is public and databooks
are client data.
"""

from __future__ import annotations

import argparse
import logging
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from fdd_utils.ai.facts import (  # noqa: E402
    build_cross_account_facts,
    build_digest_graph_links,
    build_financials_by_key,
    cross_account_links_for,
    digest_tieout_diagnostics,
    graph_summary,
    merge_graph_links,
)
from fdd_utils.ai.evidence import linked_fact_records  # noqa: E402
from fdd_utils.ai.prompts import PromptEngine  # noqa: E402
from fdd_utils.workbook import (  # noqa: E402
    extract_balance_sheet_and_income_statement,
    extract_data_from_excel,
)

_RATIO_KINDS = ("receivable_days", "payable_days", "inventory_days",
                "advance_days", "prepayment_days", "expense_to_revenue")
_PLACEHOLDER = re.compile(r"\{[a-z_]+\}")


def _edge_line(edge: dict) -> str:
    mark = "PASS" if edge["passed"] else "FAIL"
    ev = edge.get("evidence") or {}
    if edge.get("kind") == "supporting_schedule_tieout":
        extra = (f"  row={ev.get('row')} label={ev.get('row_label')!r} "
                 f"periods={edge.get('periods')} multiplier={ev.get('multiplier'):g} "
                 f"ambiguous={ev.get('ambiguous')}")
    elif "values" in ev:
        extra = (f"  values={ev['values']} ceiling={ev['band'][1]:g}"
                 f" outside={ev['outside_band']} immaterial={ev.get('immaterial')}")
    elif edge.get("kind") == "shared_counterparty":
        extra = (f"  label={ev.get('source_label')!r} "
                 f"normalised={ev.get('normalised_label')!r}")
    elif edge.get("kind") == "remark_reference":
        extra = (f"  matched={ev.get('matched_name')!r} origin={ev.get('origin')}"
                 f" excerpt={ev.get('excerpt')!r}")
    else:
        agreed = ev.get("tied_periods") or ev.get("agreeing_periods") or []
        extra = f"  agree={agreed} differ={ev.get('differing_periods')}"
    return (f"    [{mark}] {edge['source']} -> {edge['target']}"
            f"  id={edge.get('edge_id')} tier={edge.get('tier')}\n"
            f"           test: {edge['test']}\n         {extra}")


def _decoy_facts(facts: dict) -> dict:
    """Rotate each period's account totals; labels and graph stay untouched."""
    decoy = dict(facts)
    series = {name: dict(values) for name, values in (facts.get("series") or {}).items()}
    periods = sorted({period for values in series.values() for period in values})
    for period_idx, period in enumerate(periods):
        names = sorted(name for name, values in series.items() if period in values)
        if len(names) < 2:
            continue
        original = [series[name][period] for name in names]
        for idx, name in enumerate(names):
            # A different offset per period prevents moving one intact series
            # wholesale to another account, which would not be a useful decoy.
            series[name][period] = original[(idx + period_idx + 1) % len(original)]
    decoy["series"] = series
    return decoy


def _print_unmapped_diagnostics(digest: dict, facts: dict) -> None:
    rows = digest_tieout_diagnostics(digest, facts)
    reasons = {}
    for row in rows:
        reasons[row["reason"]] = reasons.get(row["reason"], 0) + 1
    print(f"=== unmapped supporting-tab tie-out diagnostics: {len(rows)} sheets ===")
    print(f"    reasons={dict(sorted(reasons.items()))}")
    print("    period_columns = dates found across columns; date_rows = dates found down rows")
    rows.sort(key=lambda row: (
        not bool(row["named_accounts"]),
        not bool(row["total_labels"]),
        not bool(row["period_columns"] or row["date_rows"]),
        -row["numeric_nonzero"],
        row["sheet"],
    ))
    for row in rows:
        best = row.get("best_near_match")
        best_text = (
            f"{best['account']} max_diff={best['max_difference_pct']:.1f}% "
            f"periods={len(best['periods'])}"
            if best else "-"
        )
        print(
            f"    sheet={row['sheet']!r}  reason={row['reason']}  "
            f"nz={row['numeric_nonzero']} blocks={row['blocks']}\n"
            f"        periods={row['period_columns']} date_rows={row['date_rows']} "
            f"totals={row['total_labels']} units={row['unit_markers']} "
            f"multiplier={row['multiplier']:g}\n"
            f"        named_accounts={row['named_accounts']} best={best_text}"
        )


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("databook")
    parser.add_argument("--sheet", default="Financials",
                        help="Financials sheet name, for the tab-to-Financials edges")
    parser.add_argument("--language", default=None, choices=["Chi", "Eng"],
                        help="override the detected report language when rendering prompts")
    parser.add_argument("--quiet", action="store_true", help="suppress prompt-engine warnings")
    parser.add_argument("--unmapped-only", action="store_true",
                        help="print why unmapped tabs did not produce tie-outs, then stop")
    args = parser.parse_args()

    if args.quiet:
        logging.disable(logging.WARNING)

    dfs, _wl, _rt, language, _res = extract_data_from_excel(
        databook_path=args.databook, entity_name="", mode="All", return_resolution=True)
    language = args.language or language
    engine = PromptEngine()
    digest = (_res or {}).get("workbook_digest")

    if args.unmapped_only:
        diagnostic_facts = build_cross_account_facts(
            dfs, type_lookup=lambda key: engine.get_mapping_component(key, component="type"))
        _print_unmapped_diagnostics(digest, diagnostic_facts)
        return 0

    financials = {}
    try:
        bs_is = extract_balance_sheet_and_income_statement(
            workbook_path=args.databook, sheet_name=args.sheet, debug=False)
        financials = build_financials_by_key(bs_is, dfs)
    except Exception as exc:
        print(f"(no Financials sheet {args.sheet!r}: {exc}; tab-to-Financials edges skipped)")

    base_facts = build_cross_account_facts(
        dfs, financials=financials,
        type_lookup=lambda key: engine.get_mapping_component(key, component="type"))
    digest_links = build_digest_graph_links(digest, base_facts)
    facts = merge_graph_links(base_facts, digest_links)

    print(f"=== {Path(args.databook).name}  language={language}  accounts={len(dfs)}  "
          f"Financials rows matched={len(financials)}")
    sources = {}
    for name, source in (facts.get("sources") or {}).items():
        sources[source] = sources.get(source, 0) + 1
    print(f"    series read from: {sources}")
    stubbed = [n for n in facts["periods"]
               if len(facts["full_periods"][n]) < len(facts["periods"][n])]
    print(f"    accounts whose tail period is a stub (kept in the table, "
          f"excluded from ratios): {len(stubbed)}")
    graph = graph_summary(facts)
    print(f"    typed graph: total={graph['total']} passed={graph['passed']} "
          f"failed={graph['failed']} malformed={graph['malformed']}")
    support = [e for e in digest_links if e.get("passed")]
    linked_sheets = {str((e.get("evidence") or {}).get("sheet")) for e in support}
    decoy_links = build_digest_graph_links(digest, _decoy_facts(facts))
    decoy_passed = sum(1 for e in decoy_links if e.get("passed"))
    false_rate = (100.0 * decoy_passed / len(support)) if support else 0.0
    print(f"    unmapped supporting tabs linked by verified total: "
          f"{len(linked_sheets)} sheets, {len(support)} edges")
    print(f"    shuffled-value decoy: {decoy_passed} surviving tie-outs; "
          f"false-link rate={false_rate:.1f}%")
    linked = {name: linked_fact_records(name, facts) for name in facts.get("series") or {}}
    linked = {name: records for name, records in linked.items() if records}
    print(f"    linked evidence closure: {len(linked)} accounts, "
          f"{sum(len(records) for records in linked.values())} facts")

    print("\n--- 1. every candidate relationship ---")
    by_kind: dict = {}
    for edge in facts["links"]:
        by_kind.setdefault(edge["kind"], []).append(edge)
    for kind in sorted(by_kind):
        edges = by_kind[kind]
        passed = sum(1 for e in edges if e["passed"])
        print(f"\n  {kind}: {len(edges)} candidates, {passed} PASS, {len(edges) - passed} FAIL")
        for edge in edges:
            print(_edge_line(edge))

    print("\n--- 2. what may be quoted (passed ratio edges, this account as source) ---")
    quotable = 0
    for name in sorted(facts["series"]):
        edges = [e for e in cross_account_links_for(facts, name, source_only=True)
                 if e["kind"] in _RATIO_KINDS]
        if edges:
            quotable += 1
            print(f"    {name}: {edges[0]['kind']} -> {edges[0]['target']} "
                  f"{(edges[0]['evidence'] or {}).get('values')}")
    print(f"    {quotable} of {len(dfs)} accounts have a quotable cross-account fact")

    print("\n--- 3. rendered prompts: leaks, substitution, size ---")
    before = after = 0
    digest_only_delta = 0
    gained = 0
    leaks = []
    bad_before = []
    bad_after = []
    rejected = [e for e in facts["links"] if not e["passed"] and e["kind"] in _RATIO_KINDS]
    for agent in ("subagent_1", "subagent_2"):
        for key, df in dfs.items():
            extra = {}
            if agent == "subagent_2":
                extra = {"account": engine.get_mapping_component(key, component="type") or key,
                         "output": "PRIOR OUTPUT PLACEHOLDER"}
            s0, u0 = engine.render_prompt(agent_name=agent, language=language, mapping_key=key,
                                          df=df, data_format="markdown", **extra)
            s1, u1 = engine.render_prompt(agent_name=agent, language=language, mapping_key=key,
                                          df=df, data_format="markdown",
                                          cross_account_facts=facts, **extra)
            sb, ub = engine.render_prompt(agent_name=agent, language=language, mapping_key=key,
                                          df=df, data_format="markdown",
                                          cross_account_facts=base_facts, **extra)
            before += len(s0) + len(u0)
            after += len(s1) + len(u1)
            digest_only_delta += len(s1) + len(u1) - len(sb) - len(ub)
            if len(s1) + len(u1) != len(s0) + len(u0):
                gained += 1
            bad_before += [(agent, key, m) for m in _PLACEHOLDER.findall(s0 + u0)]
            bad_after += [(agent, key, m) for m in _PLACEHOLDER.findall(s1 + u1)]
            rendered = s1 + u1
            for edge in rejected:
                values = [f"{v:.1f}" if abs(v) < 10 else f"{v:.0f}"
                          for v in ((edge.get("evidence") or {}).get("values") or {}).values()]
                if len(values) < 2:
                    continue
                probes = ["、".join(values), ", ".join(values[:-1]) + f" and {values[-1]}"]
                for probe in probes:
                    if probe and probe in rendered:
                        leaks.append((agent, key, edge["source"], edge["kind"], probe))

    delta = after - before
    print(f"    prompts rendered: {2 * len(dfs)}   gained a cross-account block: {gained}")
    print(f"    chars before={before:,}  after={after:,}  delta={delta:,}"
          + (f"  ({delta / gained:.0f} chars per changed prompt)" if gained else ""))
    print(f"    digest-only prompt delta={digest_only_delta}  (must be 0)")
    print(f"    unsubstituted placeholders  before={len(bad_before)}  after={len(bad_after)}"
          f"  (must be equal){'  <-- REGRESSION' if len(bad_after) > len(bad_before) else ''}")
    print(f"    REJECTED relationships leaking into a rendered prompt: {len(leaks)}"
          f"{'  <-- RULE VIOLATED' if leaks else '  (rule holds)'}")
    for leak in leaks[:5]:
        print(f"        {leak}")
    return 1 if (leaks or len(bad_after) > len(bad_before)
                 or graph["malformed"] or digest_only_delta) else 0


if __name__ == "__main__":
    raise SystemExit(main())
