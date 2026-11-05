"""Evidence contract for every account of a real workbook. Deterministic, zero tokens.

    python ad-hoc/workbench/contract_report.py <databook.xlsx> [--sheet NAME] [--account KEY] [--json OUT]

Prints, per account, the check outcome counts (account x period x kind) and the
periods where the deck's total is not the tab's; with --account, the full status
matrix, the three bases of every total and the trace of one fact back to its
sheet and row. Output names accounts and figures: keep it local.
"""
from __future__ import annotations

import argparse
import collections
import json
import re
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook")
    ap.add_argument("--sheet", default=None)
    ap.add_argument("--account", default=None)
    ap.add_argument("--json", default=None)
    args = ap.parse_args()

    from fdd_utils.ai import evidence_contract as ec
    from fdd_utils.ai.facts import build_financials_by_key
    from fdd_utils.run_identity import build_identity
    from fdd_utils.workbook import get_effective_mappings, load_mappings, process_workbook_data

    from fdd_utils.workbook.visibility import find_visible_financials_sheet
    sheet = args.sheet or find_visible_financials_sheet(args.workbook)   # a hidden sheet is never chosen
    state = process_workbook_data(temp_path=args.workbook, entity_name="", selected_sheet=sheet)
    dfs, bs_is = state["dfs"], state.get("bs_is_results") or {}
    mappings = get_effective_mappings(load_mappings(), state.get("resolution"))
    fin_by_key = {}
    for statement in ("balance_sheet", "income_statement"):
        for key, values in build_financials_by_key({statement: bs_is.get(statement)}, dfs, mappings).items():
            fin_by_key.setdefault(key, {}).update(values)
    identity = build_identity(workbook=args.workbook, entity="", sheet=sheet, financials_from=None,
                              financials_sheet=sheet, language=state.get("language", ""), model_type="", model_name=None,
                              accounts=list(dfs))
    binding = ec.SourceBinding.from_identity(identity, statement_source=str(bs_is.get("statement_source") or ""))
    digest = (state.get("resolution") or {}).get("workbook_digest")
    print(f"{Path(args.workbook).name}: {len(dfs)} account(s), source_version {binding.source_version}, "
          f"statements {binding.statement_source or '?'}")

    out = {}
    totals = collections.Counter()
    for key, df in dfs.items():
        stype = str(((df.attrs.get("integrity") or {}).get("statement_type")) or "")
        contract = ec.build_contract(key, df, statement_type=stype, financials=fin_by_key.get(key), binding=binding,
                                     digest=digest)
        out[key] = contract.to_dict()
        counts = collections.Counter(f"{c['kind'].split('_')[0]}:{c['status']}" for c in contract.checks)
        totals.update(counts)
        replaced = [p for p, perm in contract.permissions.items() if not perm["explain_with_detail"] and p in
                    {c["period"] for c in contract.checks if c.get("overridden")}]
        print(f"  {key[:34]:34s} periods {len(contract.periods)}  facts {len(contract.facts):3d}  "
              + " ".join(f"{k}={v}" for k, v in sorted(counts.items()))
              + (f"  REPLACED {replaced}" if replaced else "")
              + (f"  ISSUE {[i['code'] for i in contract.issues if i['severity'] != 'info']}"
                 if any(i["severity"] != "info" for i in contract.issues) else ""))
    print("TOTAL checks:", dict(sorted(totals.items())))

    if args.account:
        contract = ec.EvidenceContract.from_dict(out[args.account])
        print(f"\n== {args.account}: account x period x check")
        for period, row in contract.status_matrix().items():
            print(f"  {period}: {row}   permissions {{" + ", ".join(
                f"{k}={'Y' if v else 'n'}" for k, v in contract.permissions[period].items() if k != 'reasons') + "}")
        print("\n== bases of each total")
        for period in contract.periods:
            print(f"  {period}: raw {contract.total(period, ec.BASIS_RAW)} | financials "
                  f"{contract.total(period, ec.BASIS_FINANCIALS)} | adjusted {contract.total(period, ec.BASIS_ADJUSTED)}")
        latest = contract.periods[-1]
        fact = next((f for f in contract.facts if f["period"] == latest and f["role"] == "total" and f["basis"] == ec.BASIS_RAW), None)
        if fact:
            print("\n== trace of", fact["id"])
            print(json.dumps(ec.trace_fact(contract, fact["ref"]), ensure_ascii=False, indent=1, default=str)[:1800])
    if args.json:
        Path(args.json).write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
