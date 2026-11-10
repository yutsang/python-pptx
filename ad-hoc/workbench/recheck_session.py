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
            out["accounts"][key] = {"before": before, "after": after, "codes": new_codes}
    return out


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
