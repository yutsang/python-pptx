"""What the writer is given, per account, and whether a reference report's figures are in it. No model call; free.

    python inspect_evidence.py --databook <file.xlsx> [--reference <report text>] [--verbose]

Per account in scope: the detail lines shown (and summed as the rest), the tab's other tables read beside the schedule, what
was filed from related accounts, the notes and how many are MUST ADDRESS, the size of the evidence the writer gets, and which
model the fixed dispatch rule would give it (the paid run asks the stronger model instead; this is only the fallback).

With --reference (a text dump of a finished report, e.g. `inspect_pptx.py <deck> --dump-text` saved to a file), every amount
and percentage in each account's paragraph is looked up: IN VIEW (the writer is shown it), CUT (in the evidence, but not in
what the writer is shown), or NOT READ (nowhere in this account's evidence). The paragraph of an account is the text from its
name followed by a dash ("应收账款 –") to the next such header; text before the first header is checked against every account.
This tells, before any paid run, how much of a reference the writer could possibly state.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Tuple

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))


def _bar(items, desc):
    try:
        from tqdm import tqdm
        return tqdm(items, desc=desc, leave=False)
    except ImportError:
        return items


def _values(contract: Any, view: Dict[str, Any]) -> Tuple[List[Tuple[float, str]], List[Tuple[float, str]], List[float]]:
    """(amounts shown, amounts in the evidence but not shown, percentages shown), amounts in yuan with their ref."""
    from fdd_utils.agent.tools import _refs_in
    from fdd_utils.ai import claims as cl
    shown_refs = _refs_in(view)
    shown, cut = [], []
    for f in contract.facts:
        (shown if f["ref"] in shown_refs else cut).append((float(f["value"]), f["ref"]))
    pcts: List[float] = []
    for d in contract.derived:
        if d.get("value") is not None:
            (shown if d.get("ref") in shown_refs else cut).append((float(d["value"]), d["ref"]))
        if d.get("pct") is not None and d.get("ref") in shown_refs:
            pcts.append(abs(float(d["pct"])))
    for n in contract.notes:
        fig = cl.note_figures(n, contract.periods)
        for a in fig["amounts"]:
            if a.get("currency") in (None, "CNY"):
                shown.append((float(a["value"]), n["ref"]))
        pcts.extend(fig["percents"])
    return shown, cut, pcts


def _paragraphs(text: str, accounts: List[str]) -> Dict[str, str]:
    heads = sorted((m.start(), a) for a in accounts for m in re.finditer(re.escape(a) + r"(?:（续）)?\s*[–—-]", text))
    out: Dict[str, str] = {"(before the first account)": text[:heads[0][0]] if heads else text}
    for i, (pos, a) in enumerate(heads):
        end = heads[i + 1][0] if i + 1 < len(heads) else len(text)
        out[a] = out.get(a, "") + text[pos:end]
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--databook", required=True)
    ap.add_argument("--reference", default="", help="a finished report's text: which of its figures the writer is shown")
    ap.add_argument("--entity", default="")
    ap.add_argument("--verbose", action="store_true", help="every figure, not only the ones the writer is not shown")
    args = ap.parse_args()

    from fdd_utils.agent.controller import AgentController
    from fdd_utils.agent.dispatch import account_profiles, rule_plan
    from fdd_utils.agent.session import SourceSession
    from fdd_utils.agent.tools import get_account_evidence
    from fdd_utils.ai import claims as cl

    print(f"reading {args.databook} (no model call) ...", flush=True)
    s = SourceSession.load(args.databook, entity=args.entity)
    accounts = [k for k in sorted(s.contracts) if s.in_scope(k)]
    views = {k: get_account_evidence(s, k) for k in _bar(accounts, "evidence")}
    plan = rule_plan(account_profiles(s, accounts))
    print(f"\n{len(accounts)} accounts in scope; source {s.source_version}")
    print("account                lines shown/total rest | tables (rows) | related bal/notes | notes must | chars | rule")
    total_chars = 0
    for k in accounts:
        v, c = views[k], s.contracts[k]
        comp = v["components"]
        lines = len(comp["lines"])
        rest = (comp.get("rest") or {}).get("lines", 0)
        tables = v.get("supporting_tables") or []
        rel_notes = sum(1 for n in c.notes if n.get("kind") == "related_note")
        chars = len(json.dumps(v, ensure_ascii=False, separators=(",", ":")))
        total_chars += chars
        must = len(AgentController.__new__(AgentController)._must_address(v))
        flag = "  CUT>=1%: " + str(v["material_not_shown"]) if v.get("material_not_shown") else ""
        print(f"{k[:20]:20s} {lines:5d}/{lines + v.get('components_not_shown', 0):<5d} {rest:4d} | "
              f"{len(tables)} ({sum(len(t['rows']) for t in tables)}) | {len(v.get('related') or [])}/{rel_notes} | "
              f"{len(v['notes']):3d} {must:3d} | {chars:6,d} | {plan.get(k, '?')}{flag}")
        if args.verbose:
            for t in tables:
                print(f"      table {t['where']} {str(t['title'])[:40]!r}: {[c_['head'][:12] for c_ in t['columns']]}")
    print(f"evidence for all accounts: {total_chars:,} characters (about {total_chars // 2:,} tokens of Chinese JSON)")
    print(f"rule plan: {sum(1 for x in plan.values() if x == 'pro')} of {len(plan)} to the stronger model")

    if not args.reference:
        return 0
    text = Path(args.reference).read_text(encoding="utf-8")
    paras = _paragraphs(text, accounts)
    vals = {k: _values(s.contracts[k], views[k]) for k in accounts}
    counts = {"in view": 0, "cut": 0, "not read": 0}
    print("\nREFERENCE figures per account: in view / cut / not read" + ("" if args.verbose else " (not-in-view listed)"))
    for name, para in paras.items():
        targets = [k for k in accounts if k == name] or accounts
        found: List[Tuple[str, str]] = []
        for sp in cl.text_amounts(para):
            if not sp["value"]:
                continue
            status = "not read"
            for k in targets:
                shown, cut, _p = vals[k]
                if any(abs(abs(sp["value"]) - abs(x)) <= sp["tolerance"] for x, _r in shown):
                    status = "in view"
                    break
                if any(abs(abs(sp["value"]) - abs(x)) <= sp["tolerance"] for x, _r in cut):
                    status = "cut"
            found.append((sp["written"], status))
        for sp in cl.text_percents(para):
            if sp.get("points"):
                continue
            ok = any(any(abs(abs(sp["value"]) - p) <= 0.05 for p in vals[k][2]) for k in targets)
            found.append((sp["written"], "in view" if ok else "not read"))
        if not found:
            continue
        n = {st: sum(1 for _w, x in found if x == st) for st in counts}
        for st in counts:
            counts[st] += n[st]
        listed = [f"{w} ({st})" for w, st in found if args.verbose or st != "in view"]
        print(f"  {name[:20]:20s} {n['in view']:3d} / {n['cut']:3d} / {n['not read']:3d}" + (f"   {', '.join(listed)}" if listed else ""))
    total = sum(counts.values()) or 1
    print(f"TOTAL: {counts['in view']} in view, {counts['cut']} cut, {counts['not read']} not read, of {total} "
          f"({counts['in view'] / total:.0%} of the reference's figures are in what the writer is shown)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
