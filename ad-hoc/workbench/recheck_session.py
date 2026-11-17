"""Run the verifier as the code is now over the drafts a saved session holds, and print which sentences the deck would mark
before and after. No model call; free. The workbook is read again (the evidence the drafts were checked against).

    python ad-hoc/workbench/recheck_session.py --databook <xlsx> --session <trial_<arm>_<stamp>.agent-session.json> [--width 400]

Phase A (2026-11-10) froze one draft of the client file; a new check is measured on that same text, not on a new draft
that differs in its own ways. The saved session is read from a copy and never written. The output holds client sentences:
it stays in the chat it is pasted to.
"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
import tempfile
from pathlib import Path
from types import SimpleNamespace
from typing import Any, Dict, List

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ROOT))


def recheck(databook: str, session: str) -> Dict[str, Any]:
    """{"id", "notice", "accounts": {account: {"before": marks, "after": marks, "codes": new codes}}}."""
    from entity_trial import deck_marks
    from fdd_utils.agent.workspace import WorkSession
    from fdd_utils.ai import claims as cl
    saved = json.loads(Path(session).read_text(encoding="utf-8"))
    with tempfile.TemporaryDirectory(prefix="recheck_") as tmp:
        copy = Path(tmp) / Path(session).name                  # open() would save into the file it resumed from
        shutil.copyfile(session, copy)
        work = WorkSession.open(databook, session_path=str(copy), resume=True)
        work.path = None
        out: Dict[str, Any] = {"id": saved.get("id"), "notice": getattr(work, "opening_notice", "") or "", "accounts": {}}
        for key, c in (saved.get("candidates") or {}).items():
            claims, flags = list(c.get("claims") or []), list(c.get("flags") or [])
            stamped = [dict(x, account=key, source_version=work.source.source_version) for x in claims]
            now = cl.verify_candidate(stamped, work.source.contracts, default_account=key) if claims else {}
            before = deck_marks(key, SimpleNamespace(claims=claims, outcome=c.get("outcome") or {}, flags=flags))
            after = deck_marks(key, SimpleNamespace(claims=claims, outcome=now, flags=flags))
            was = {m["claim"] for m in before}
            new_codes = sorted({f["code"] for n, v in enumerate(now.get("verdicts") or []) if n not in was
                                for f in v.get("findings") or [] if f.get("severity") in ("failed", "review")})
            red = {m["claim"] for m in after if m["colour"] == "red"}
            out["accounts"][key] = {"before": before, "after": after, "codes": new_codes,
                                    "red": [red_detail(n, key, claims[n], (now.get("verdicts") or [])[n], work.source.contracts)
                                            for n in sorted(red) if n < len(claims) and n < len(now.get("verdicts") or [])]}
    return out


def red_detail(n: int, key: str, claim: Dict[str, Any], verdict: Dict[str, Any], contracts: Dict[str, Any]) -> Dict[str, Any]:
    """What a person needs to tell a wrong sentence from a wrong refusal: every failed finding with its message, the period each
    amount in the sentence is read as written under, and every fact the claim cites with its line, period and value."""
    from fdd_utils.ai import claims as cl
    text = str(claim.get("text") or "")
    mentions = cl.period_mentions(text)
    binds = []
    for sp in cl.text_amounts(text):
        m = cl.attached_period(text, sp["start"], sp["end"], mentions)
        binds.append(f"{sp['written']} -> {m['text'] if m else 'no period'}")
    cited = []
    for ref in claim.get("cites") or []:
        item, _owner = cl._lookup(str(ref), contracts, key)
        if item is None:
            note = (contracts.get(key).note(str(ref)) if contracts.get(key) is not None else None)
            if note is not None:        # a note: its text and the figures the verifier reads from it, each with the period it gives them
                figs = cl.note_figures(note, contracts[key].periods)
                shown = ", ".join(f"{a['written']}@{a.get('period') or '?'}" for a in figs["amounts"][:6])
                cited.append(f"{ref} note[{figs.get('period') or '?'}] {str(note.get('text') or '')[:130]!r} figures: {shown}")
            else:
                cited.append(f"{ref}: not found")
            continue
        value = item.get("value")
        cited.append(f"{item.get('ref') or ref} {item.get('period') or '/'.join(map(str, item.get('periods') or []))} "
                     f"{str(item.get('line') or item.get('kind') or '')[:30]} "
                     f"{'' if value is None else format(float(value) / 10000, ',.2f') + '万'}")
    return {"claim": n,
            "findings": [f"{f.get('code')}: {str(f.get('message') or '')[:300]}"
                         for f in verdict.get("findings") or [] if f.get("severity") == "failed"],
            "binds": binds, "cited": cited, "declared": list(claim.get("periods") or [])}


def report(result: Dict[str, Any], width: int = 400) -> List[str]:
    rows = result["accounts"]
    count = lambda side: (sum(m["colour"] == "red" for r in rows.values() for m in r[side]),
                          sum(m["colour"] == "orange" for r in rows.values() for m in r[side]))
    (rb, ob), (ra, oa) = count("before"), count("after")
    gained = {k: [m for m in r["after"] if m["claim"] not in {x["claim"] for x in r["before"]}] for k, r in rows.items()}
    lost = {k: [m for m in r["before"] if m["claim"] not in {x["claim"] for x in r["after"]}] for k, r in rows.items()}
    # marked before and after, but in another colour (a red period refusal that is now an orange cause, say)
    moved = {k: [(b, a) for b in r["before"] for a in r["after"] if a["claim"] == b["claim"] and a["colour"] != b["colour"]]
             for k, r in rows.items()}
    out = [f"RECHECK {result.get('id') or '?'}: {len(rows)} accounts; marks before {rb} red {ob} orange, now {ra} red {oa} orange; "
           f"new {sum(map(len, gained.values()))}, gone {sum(map(len, lost.values()))}, recoloured {sum(map(len, moved.values()))}"]
    if result.get("notice"):
        out.append(f"note: {result['notice']}")
    for key in rows:
        if not gained[key] and not lost[key] and not moved[key]:
            continue
        out.append(key)
        for b, a in moved[key]:
            cut = a["text"] if len(a["text"]) <= width else a["text"][:width] + "…"
            out.append(f"  ~ {b['colour'][0].upper()}>{a['colour'][0].upper()} [{a['claim']}] {a['why']}")
            out.append(f"      {cut}")
        for sign, marks in (("+", gained[key]), ("-", lost[key])):
            for m in marks:
                cut = m["text"] if len(m["text"]) <= width else m["text"][:width] + "…"
                out.append(f"  {sign} {m['colour'][0].upper()} [{m['claim']}] {m['why']}")
                out.append(f"      {cut}")
    # red marks that are still red: they do not appear above (nothing changed), and a person cannot judge one without its numbers
    for key, r in rows.items():
        for d in r.get("red") or []:
            mark = next((m for m in r["after"] if m["claim"] == d["claim"]), None)
            if mark is None:
                continue
            cut = mark["text"] if len(mark["text"]) <= width else mark["text"][:width] + "…"
            out.append(f"STILL RED {key} [{d['claim']}] (claim declares periods {', '.join(d['declared']) or 'none'})")
            out.append(f"      {cut}")
            out.extend(f"    failed  {line}" for line in d["findings"])
            out.append("    amount read under:  " + "; ".join(d["binds"]))
            out.extend(f"    cites   {line}" for line in d["cited"])
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--databook", required=True, help="the workbook the session was written on")
    ap.add_argument("--session", required=True, help="trial_<arm>_<stamp>.agent-session.json")
    ap.add_argument("--width", type=int, default=400, help="characters of each marked sentence to show")
    args = ap.parse_args()
    print("\n".join(report(recheck(args.databook, args.session), args.width)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
