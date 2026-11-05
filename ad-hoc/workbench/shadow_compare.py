"""Legacy run against the agent path's capabilities, side by side. Offline, zero tokens.

    python ad-hoc/workbench/shadow_compare.py <databook.xlsx> --base-run RUN_ID [--json OUT]

What can be compared WITHOUT a model, on one workbook and one earlier run:

  trace coverage     of the amounts the legacy commentary prints, how many are some fact of the
                     account (by value), and how many sit under a period the sentence does not name
                     (a right number under the wrong date, which the pool search never noticed)
  legacy's own flags the run's clause reviews: hallucination / reasoning counts, residual after retries
  agent capability   accounts whose balance can be stated at all (tab and Financials agree, or the deck
                     follows Financials), accounts with a replaced period, detail that does or does not
                     add up -- what a verifier would refuse to let a claim assert
  artifact recall    export a deck from validated balance claims for every commented account, shift each
                     account's cited fact in turn and ask which paragraphs the returned-deck reader flags:
                     recall (the account's own paragraph is flagged) and precision (only it is)
  usage              the legacy run's recorded calls; the agent path's offline calls (0)

What it CANNOT compare, and says so in its output: errors in text a MODEL writes through the agent path,
useful analysis kept or lost, layout of agent-written text in real PowerPoint, how a person's edits in
PowerPoint survive a real save. Those need fresh provider calls and a Windows PowerPoint.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
os.chdir(ROOT)


def legacy_text(result) -> str:
    from fdd_utils.financial_common import get_pipeline_result_text
    return get_pipeline_result_text(result)


def trace_coverage(session, base_results):
    from fdd_utils.ai.claims import mention_matches, period_mentions, text_amounts
    rows, totals = [], {"amounts": 0, "by_value": 0, "wrong_period": 0, "unmatched": 0}
    for account, result in base_results.items():
        contract = session.contracts.get(account)
        if str(account).startswith("__") or contract is None or not isinstance(result, dict):
            continue
        text = legacy_text(result)
        facts = [f for f in contract.facts if f.get("value") is not None]
        counts = {"amounts": 0, "by_value": 0, "wrong_period": 0, "unmatched": 0}
        for clause in re.split(r"(?<=[。；;])", text):
            mentions = period_mentions(clause)
            for span in text_amounts(clause):
                counts["amounts"] += 1
                hits = [f for f in facts if abs(abs(f["value"]) - abs(span["value"])) <= span["tolerance"]]
                if not hits:
                    counts["unmatched"] += 1
                    continue
                counts["by_value"] += 1
                if mentions and not any(mention_matches(m, f["period"]) for m in mentions for f in hits):
                    counts["wrong_period"] += 1
        for k in totals:
            totals[k] += counts[k]
        rows.append({"account": account, **counts})
    return rows, totals


def capability(session):
    out = {"accounts": len(session.contracts), "balance_statable_latest": 0, "replaced": 0, "detail_untied_latest": 0, "detail_ambiguous_latest": 0}
    for contract in session.contracts.values():
        if not contract.periods:
            continue
        last = contract.periods[-1]
        perm = contract.permissions.get(last) or {}
        out["balance_statable_latest"] += bool(perm.get("describe_total"))
        out["replaced"] += any(c.get("overridden") for c in contract.checks)
        out["detail_untied_latest"] += perm.get("detail_basis") == "untied"
        out["detail_ambiguous_latest"] += perm.get("detail_basis") == "ambiguous"
    return out


def artifact_recall(session, tmp):
    from fdd_utils.agent import tools as T
    from fdd_utils.agent import updates as U
    accounts = [a for a in session.base_results if not str(a).startswith("__") and a in session.contracts and session.contracts[a].periods]
    for a in accounts:
        c = session.contracts[a]
        last = c.periods[-1]
        total = next((f for f in c.facts if f["period"] == last and f["role"] == "total" and f["basis"] == "authority_adjusted"), None)
        if total is None or not (c.permissions.get(last) or {}).get("describe_total"):
            continue
        y, m, d = (int(x) for x in last.split("-"))
        T.call_tool(session, "validate_candidate", {"account": a, "claims": [{"kind": "amount", "periods": [last], "cites": [total["ref"]],
                    "text": f"截至{y}年{m}月{d}日，{a}余额{abs(total['value']) / 1e4:,.1f}万元"}]})
    res = T.call_tool(session, "export_candidate", {"out_path": os.path.join(tmp, "shadow.pptx")})
    path = (res.data or {}).get("path")
    if not path or not os.path.exists(path):
        return {"error": (res.error or {}).get("message")}
    from fdd_utils.pptx import lineage as L
    public = L.read_manifest(path)
    cited = {e["account"] for e in public["elements"] if e["cites"]}
    hits = flagged_total = 0
    for a in sorted(cited):
        saved = [(f, f["value"]) for f in session.contracts[a].facts if f["role"] == "total"]
        for f, _ in saved:
            f["value"] += 1_000_000.0
        report = U.analyze_returned(path, session, public=public, internal=L.read_internal(path))
        flagged = {e["account"] for e in report["elements"] if e["source_changed"]}
        for f, v in saved:
            f["value"] = v
        hits += a in flagged
        flagged_total += len(flagged)
    n = len(cited)
    return {"cited_accounts": n, "recall": (hits / n) if n else None, "precision": (hits / flagged_total) if flagged_total else None,
            "deck_status": res.data.get("delivery", {}).get("status") if res.data else None}


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("databook")
    ap.add_argument("--base-run", required=True)
    ap.add_argument("--json", default=None)
    args = ap.parse_args()
    from fdd_utils import run_identity
    from fdd_utils.agent.session import SourceSession

    with run_identity.offline_provider_block() as attempts:
        session = SourceSession.load(args.databook, base_run=args.base_run)
        rows, totals = trace_coverage(session, session.base_results or {})
        cap = capability(session)
        with tempfile.TemporaryDirectory() as tmp:
            recall = artifact_recall(session, tmp)
    health = (session.base_results or {}).get("__run_health__") or {}
    reviews = [r for k, v in (session.base_results or {}).items() if not str(k).startswith("__") and isinstance(v, dict)
               for r in (v.get("agent_4_validation") or {}).get("clause_reviews") or []]
    legacy_flags = {"hallucination": sum(1 for r in reviews if not r.get("supported") and r.get("category") == "hallucination"),
                    "reasoning": sum(1 for r in reviews if not r.get("supported") and r.get("category") == "reasoning"),
                    "clauses": len(reviews)}
    report = {"workbook": os.path.basename(args.databook), "base_run": args.base_run, "legacy_trace": totals, "legacy_flags": legacy_flags,
              "agent_capability": cap, "artifact_recall": recall,
              "usage": {"legacy_calls_succeeded": health.get("calls_succeeded"), "legacy_calls_failed": health.get("calls_failed"),
                        "agent_offline_provider_calls": len(attempts)}}
    print(f"{report['workbook']}  vs legacy run {args.base_run}  ({len(rows)} commented account(s))")
    t = totals
    print(f"legacy amounts printed {t['amounts']}: by value in the account {t['by_value']} "
          f"({100 * t['by_value'] / max(t['amounts'], 1):.0f}%), under a period the sentence does not name {t['wrong_period']}, no fact at all {t['unmatched']}")
    print(f"legacy own flags: {legacy_flags}")
    print(f"agent capability: balance statable (latest) {cap['balance_statable_latest']}/{cap['accounts']}, accounts with a replaced period {cap['replaced']}, "
          f"detail untied {cap['detail_untied_latest']}, nesting unknown {cap['detail_ambiguous_latest']}")
    print(f"artifact recall (shift each cited account's total): {recall}")
    print(f"usage: legacy calls ok/failed {health.get('calls_succeeded')}/{health.get('calls_failed')}; agent path offline provider calls {len(attempts)}")
    print("NOT COMPARABLE OFFLINE (needs fresh model calls / Windows PowerPoint): claim errors in model-written text, useful analysis retained, "
          "layout of agent-written text, human edits surviving a real PowerPoint save.")
    if args.json:
        Path(args.json).write_text(json.dumps(report, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    return 0 if not attempts else 1


if __name__ == "__main__":
    sys.exit(main())
