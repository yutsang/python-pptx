"""Free; no model call. Two questions about a workbook the agent has read.

    python ad-hoc/workbench/gap_probe.py --session trial_<arm>_<stamp>.agent-session.json --dir databooks
        For every account whose evidence says nothing on something its paragraph is expected to carry (knowledge.py: a loan's rate
        and maturity, ageing, a depreciation policy, what cash is made of): is that something written anywhere on a visible
        sheet? One search per word, on its own sheet and elsewhere, with the cell's address and the first words of it. A gap with
        no hit anywhere is a source gap; a gap with a hit on another sheet is knowledge the writer was not given.

    python ad-hoc/workbench/gap_probe.py --session ... --dir databooks --account 应付账款
        How one account's lines sit in the evidence: for its latest period each detail line with its ref, whether the writer is
        shown it, the parent it was found nested in (blank: none found), and the sums Python derived (the formula and the refs it
        adds). A sentence that names a parent and its own child, or "the rest" beside lines of which one holds another, fails
        REST_INCOMPLETE / SUM_STATED here; this shows which.

The session file names the workbook by a copy that is gone; --dir is where to find it by sha256 (as period_audit.py does).
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Any, Dict, List

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parents[1]))


def text_of(cell: Dict[str, Any], width: int = 56) -> str:
    return re.sub(r"\s+", " ", str(cell.get("text") or ""))[:width]


def gaps(session: Any, width: int) -> List[str]:
    from fdd_utils.agent import investigate as I
    from fdd_utils.agent import knowledge as K
    from fdd_utils.agent import tools as T
    out: List[str] = []
    seen = real = 0
    for key in sorted(session.contracts):
        if str(key).startswith("__") or not session.in_scope(key) or not K.needs_for(key):
            continue
        result = T.call_tool(session, "get_account_evidence", {"account": key})
        if not result.ok:
            continue
        missing = K.missing(key, result.data)
        if not missing:
            continue
        seen += 1
        own = str((session.dfs[key].attrs or {}).get("source_sheet_name") or "") if key in session.dfs else ""
        out.append(f"{key}  (own sheet {own or '?'})")
        found_any = False
        for g in missing:
            words = [w.strip() for w in g["query"].split("/") if w.strip()]
            hits: Dict[str, Dict[str, Any]] = {}
            for w in words:
                try:
                    found = I.search_notes(session, w, account="", max_hits=8)
                except Exception as exc:                                  # noqa: BLE001
                    out.append(f"    search {w!r} failed: {type(exc).__name__}: {exc}")
                    continue
                for h in found["hits"]:
                    hits.setdefault(h["location"], h)
            here = [h for h in hits.values() if h["sheet"] == own]
            elsewhere = [h for h in hits.values() if h["sheet"] != own]
            found_any = found_any or bool(hits)
            out.append(f"  - {g['label']}: {len(here)} hit(s) on its own sheet, {len(elsewhere)} elsewhere (words {', '.join(words)})")
            for h in (here[:2] + elsewhere[:3]):
                out.append(f"      {h['location']}: {text_of(h, width)}")
        real += 0 if found_any else 1
    out.insert(0, f"GAP PROBE  {seen} account(s) with something their evidence does not say; {real} of them have no matching cell on any visible sheet")
    return out


def lines_of(session: Any, account: str) -> List[str]:
    from fdd_utils.ai import evidence_contract as ec
    from fdd_utils.agent import tools as T
    key = session.resolve_account(account)
    if not key:
        return [f"no single account matches {account!r}: {', '.join(sorted(session.contracts))[:300]}"]
    contract = session.contracts[key]
    result = T.call_tool(session, "get_account_evidence", {"account": key})
    view = result.data if result.ok else {}
    shown = T._refs_in(view.get("components") or {})
    latest = contract.periods[-1] if contract.periods else ""
    out = [f"ACCOUNT {key}  latest period {latest}  periods {contract.periods}"]
    nested = {f["ref"]: f.get("nested_in") for f in contract.facts}
    rows = [f for f in contract.facts if f["period"] == latest and f.get("role") == ec.ROLE_COMPONENT]
    out.append(f"  {len(rows)} detail line(s); the writer is shown {sum(1 for f in rows if f['ref'] in shown)}; "
               f"lines found nested in another: {sum(1 for f in rows if f.get('nested_in'))}")
    for f in sorted(rows, key=lambda f: -abs(f.get("value") or 0))[:40]:
        out.append(f"  {f['ref']:>5} {'shown' if f['ref'] in shown else 'cut  '} {('in ' + str(f.get('nested_in'))) if f.get('nested_in') else '':8s} "
                   f"{(f.get('value') or 0) / 10000:>12,.2f}万  {str(f.get('line'))[:44]}")
    out.append("  derived: " + ("; ".join(f"{d['ref']} {d.get('kind')} {d.get('formula') or ''}"[:110] for d in view.get('derived') or []) or "none"))
    totals = [t for t in view.get("totals") or [] if t.get("period") == latest]
    out.append("  totals: " + "; ".join(f"{t.get('ref')} {t.get('basis')} {(t.get('yuan') or 0) / 10000:,.2f}万" for t in totals))
    checks = (view.get("checks") or {}).get(latest)
    out.append(f"  checks {checks}   not shown: {view.get('components_not_shown')} line(s), {view.get('material_not_shown')} of them at least 1%")
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--session", default="")
    ap.add_argument("--dir", default="databooks")
    ap.add_argument("--databook", default="")
    ap.add_argument("--entity", default="")
    ap.add_argument("--sheet", default="")
    ap.add_argument("--financials-from", default="")
    ap.add_argument("--financials-sheet", default="")
    ap.add_argument("--account", default="", help="show how this account's lines nest instead of listing gaps")
    ap.add_argument("--width", type=int, default=56, help="characters of a hit's text")
    args = ap.parse_args()
    if not args.session and not args.databook:
        ap.error("pass --session or --databook")
    import warnings
    warnings.filterwarnings("ignore")
    import period_audit as PA
    session, _params = PA.load_session(args)
    print("\n".join(lines_of(session, args.account) if args.account else gaps(session, args.width)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
