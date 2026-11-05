"""What real material this machine has to validate the agent on. Free, no model.

    python ad-hoc/workbench/inventory_sources.py [--json OUT]

For every workbook in the repo root: its hash, visible and hidden sheets, the Financials sheet, units, the
accounts in scope, how many tie to Financials on their latest period, the periods; and for every archived run
whose identity names that workbook (by hash): how many accounts it commented on, and what kind of run it was: "real" (the usage
ledger holds provider-reported calls inside the run's window), "stub-like" (several accounts carry the same canned commentary) or
"unknown" (no ledger on this machine to say; the run's own "calls ok" line is not used, it counts a stub as a success). Plus the
template. baseline_pack.py is the tool that proves a run describes a workbook and packs them together.

Workbooks are printed as W1, W2 ... so the output can be pasted back without client names; --json writes the
real file names, keep that one local. A run with no identity cannot be tied to a workbook and is counted, not used.
"""
from __future__ import annotations

import argparse
import collections
import glob
import json
import os
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))


def run_facts(folder: str) -> dict:
    """Health and size of one archived run, from what the run itself wrote and (when this machine has one) the usage ledger."""
    import yaml
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import baseline_pack as B
    out = {"run": os.path.basename(folder), "accounts": 0, "model_calls_ok": None, "model_calls_failed": None,
           "real_model": None, "kind": "unknown", "unresolved_flags": None, "has_summary": os.path.exists(os.path.join(folder, "section_summaries.json"))}
    ledger = ROOT / "fdd_utils" / "logs" / "usage.jsonl"
    identity = B.read_json(Path(folder) / "run_identity.json") or {}
    start, end = B.run_window(Path(folder))
    rows = B.ledger_rows(ledger, {identity.get("workbook")}, start, end) if ledger.exists() else []
    out["kind"] = B.run_kind(Path(folder), rows)
    out["real_model"] = out["kind"] == "real" if out["kind"] != "unknown" else None
    try:
        results = yaml.safe_load(open(os.path.join(folder, "results.yml"), encoding="utf-8")) or {}
    except Exception:
        return out
    out["accounts"] = len([k for k in results if not str(k).startswith("__")])
    summary = os.path.join(folder, "summary.txt")
    if os.path.exists(summary):
        head = open(summary, encoding="utf-8").read(400)
        m = re.search(r"calls ok (\d+) / failed (\d+)", head)
        if m:
            out["model_calls_ok"], out["model_calls_failed"] = int(m.group(1)), int(m.group(2))
    log = os.path.join(folder, "processing.log")
    if os.path.exists(log):
        text = open(log, encoding="utf-8", errors="ignore").read()
        out["provider_in_log"] = (re.findall(r"model[=: ]+([a-z\-]+)", text[:4000], re.I) or [None])[0]
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--json", default=None)
    args = ap.parse_args()
    os.chdir(ROOT)
    from fdd_utils.run_identity import file_sha256
    from fdd_utils.agent.session import SourceSession
    from fdd_utils.agent.tools import financials_pairing
    from fdd_utils.workbook.visibility import find_visible_financials_sheet, sheet_states

    template = ROOT / "fdd_utils" / "template.pptx"
    print(f"template.pptx: {'present' if template.exists() else 'MISSING'}"
          + (f" ({template.stat().st_size // 1024} KB)" if template.exists() else ""))
    identities = {}
    for path in glob.glob(str(ROOT / "fdd_utils/logs/run_*/run_identity.json")):
        try:
            identities[os.path.dirname(path)] = json.load(open(path, encoding="utf-8"))
        except Exception:
            continue
    no_identity = [d for d in glob.glob(str(ROOT / "fdd_utils/logs/run_*")) if d not in identities]

    report = {}
    books = [p for p in sorted(ROOT.glob("*.xlsx")) if "reconciliation_report" not in p.name and not p.name.startswith("~$")]
    for n, book in enumerate(books, 1):
        alias = f"W{n}"
        sha = file_sha256(str(book))
        states = sheet_states(str(book)) or {}
        hidden = [s for s, st in states.items() if st != "visible"]
        sheet = find_visible_financials_sheet(str(book))
        row = {"file": book.name, "sha256": sha[:12], "sheets": len(states), "hidden_sheets": len(hidden),
               "financials_sheet": sheet, "alias": alias}
        try:
            session = SourceSession.load(str(book))
        except Exception as exc:                                    # noqa: BLE001
            row["load_error"] = f"{type(exc).__name__}: {exc}"
            report[alias] = row
            print(f"{alias}: could not be loaded: {row['load_error'][:120]}")
            continue
        periods = sorted({p for c in session.contracts.values() for p in c.periods})
        pairing = financials_pairing(session)
        units = sorted({str(u) for s in ((session.digest or {}).get("sheets") or {}).values()      # digest-ok: unit markers only
                        if isinstance(s, dict) and not s.get("hidden") for u in (s.get("unit_markers") or [])})
        row.update({
            "accounts": len(session.contracts), "in_scope": sum(session.in_scope(k) for k in session.contracts),
            "statement_source": session.binding.statement_source, "language": session.language, "entity": session.entity_name or "(single)",
            "periods": [periods[0], periods[-1]] if periods else [], "n_periods": len(periods),
            "tie_latest": f"{pairing['agreeing_latest']}/{pairing['compared']}", "pairing_suspect": pairing["suspect"],
            "unit_markers_seen": units[:6], "source_version": session.source_version,
            "replaced_periods": sum(1 for c in session.contracts.values() if any(x.get("overridden") for x in c.checks)),
        })
        runs = [os.path.basename(d) for d, i in identities.items() if i.get("workbook_sha256") == sha]
        facts = [run_facts(str(ROOT / "fdd_utils/logs" / r)) for r in sorted(runs)]
        row["runs_by_hash"] = len(facts)
        row["real_model_runs"] = [f for f in facts if f["kind"] == "real"][-3:]
        row["stub_runs"] = sum(1 for f in facts if f["kind"] == "stub-like")
        row["unknown_kind_runs"] = sum(1 for f in facts if f["kind"] == "unknown")
        report[alias] = row
        print(f"{alias}: {row['sheets']} sheets ({row['hidden_sheets']} hidden)  Financials sheet: {'yes' if sheet else 'NONE'}  "
              f"entity: {row['entity']}  lang {row['language']}  units {row['unit_markers_seen'] or '?'}")
        print(f"     {row['accounts']} accounts ({row['in_scope']} in scope)  periods {row['n_periods']} "
              f"[{row['periods'][0] if periods else '-'} .. {row['periods'][-1] if periods else '-'}]  statements {row['statement_source']}  "
              f"tie to Financials on latest period {row['tie_latest']}{'  SUSPECT' if row['pairing_suspect'] else ''}  replaced-total accounts {row['replaced_periods']}")
        print(f"     archived runs naming this file by hash: {row['runs_by_hash']}  (stub-like: {row['stub_runs']}, "
              f"real (ledger shows provider calls): {len([f for f in facts if f['kind'] == 'real'])}, kind unknown: {row['unknown_kind_runs']})")
        for f in row["real_model_runs"]:
            print(f"       real run {f['run']}: {f['accounts']} accounts, calls ok {f['model_calls_ok']} / failed {f['model_calls_failed']}, "
                  f"section summaries {'yes' if f['has_summary'] else 'no'}")
    print(f"\narchived runs with no identity (cannot be tied to a workbook, not used): {len(no_identity)}")
    if args.json:
        Path(args.json).write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8")
        print("written:", args.json, "(real file names inside: keep local)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
