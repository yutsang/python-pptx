"""Free; no model call. What the writer was told when its draft was refused, and what became of it: the evidence for deciding which
refusals an agent can settle by itself and which it cannot.

    python ad-hoc/workbench/repair_pack.py --session <trial_<arm>_<stamp>.agent-session.json> [--width 140]
    python ad-hoc/workbench/portfolio_trial.py --dir databooks --entities A,B --learn          every book's newest paid trial

Per session:
    BY CODE      each refusal code: in how many accounts it came up, how many of those accounts ended accepted, how many did not
    NOT ACCEPTED each account that ended with a refused sentence: its rounds (the codes of each), and the sentences still refused
                 with the verifier's message, i.e. what the writer could not fix in the rounds it had
    REPAIRED     accounts that were refused at least once and ended accepted: the codes of each round
    INFERENCE    reasons no note supports that were sent back once (agent.repair_inference), and how many were still orange after

A session written before 2026-11-17 holds each round's refusal message (cut at 600 characters) and no drafts; one written after also
holds the sentences refused and what the rewrite put in their place. The pack prints what there is. It holds client sentences:
it stays in the chat it is pasted to.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Dict, List

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parents[1]))

_CODE = re.compile(r"\b([A-Z][A-Z0-9_]{3,}):")


def round_codes(repair: Dict[str, Any]) -> List[str]:
    """The refusal codes of one round: from the trace when the round has one, else from the message's own 'CODE:' prefixes."""
    traced = sorted({c for r in repair.get("refused") or [] for c in r.get("codes") or []})
    return traced or sorted(set(_CODE.findall(str(repair.get("feedback") or ""))))


def by_account(session: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    """{account: {'repairs': [...], 'rewrites': [...], 'ok': bool, 'stop': str}}. A turn that did not end accepted lists no account, so
    the turn is found from the candidate's own turn number; a turn that names exactly one account is taken as that account's."""
    turns = {t.get("n"): t for t in session.get("turns") or []}
    out: Dict[str, Dict[str, Any]] = {}

    def take(account: str, turn: Dict[str, Any]) -> None:
        report = turn.get("report") or {}
        slot = out.setdefault(account, {"repairs": [], "rewrites": [], "ok": False, "stop": ""})
        slot["repairs"].extend(report.get("repairs") or [])
        slot["rewrites"].extend(report.get("rewrites") or [])
        slot["ok"], slot["stop"] = bool(turn.get("ok")), str(turn.get("stop_reason") or "")
    for account, cand in (session.get("candidates") or {}).items():
        turn = turns.get(cand.get("turn"))
        if turn is not None:
            take(account, turn)
    for turn in session.get("turns") or []:
        accounts = turn.get("accounts") or []
        if len(accounts) == 1 and accounts[0] not in out:
            take(accounts[0], turn)
    return out


def residual(candidate: Dict[str, Any], width: int) -> List[str]:
    """The sentences of the last draft the verifier still refused, each with its failed findings."""
    out = []
    claims = candidate.get("claims") or []
    for n, v in enumerate((candidate.get("outcome") or {}).get("verdicts") or []):
        if v.get("verdict") in ("verified", "verified_inference") or n >= len(claims):
            continue
        bad = [f"{f.get('code')}: {str(f.get('message') or '')[:260]}" for f in v.get("findings") or [] if f.get("severity") == "failed"]
        out.append(f"      [{n}] {str(claims[n].get('text') or '')[:width]}")
        out.extend(f"          {b}" for b in bad[:3])
    return out


def pack(session: Dict[str, Any], width: int = 140) -> List[str]:
    cands = session.get("candidates") or {}
    accounts = by_account(session)
    refused = {a: s for a, s in accounts.items() if any(not r.get("inference_repair") for r in s["repairs"])}
    accepted = {a for a, c in cands.items() if c.get("accepted")}
    out = [f"REPAIR PACK {session.get('id') or '?'}: {len(cands)} accounts, accepted {len(accepted)}, refused at least once {len(refused)}, "
           f"ended not accepted {len(cands) - len(accepted)}"]
    seen: Dict[str, Dict[str, set]] = defaultdict(lambda: {"accounts": set(), "ended_ok": set()})
    for a, s in refused.items():
        for r in s["repairs"]:
            if r.get("inference_repair"):
                continue
            for code in round_codes(r):
                seen[code]["accounts"].add(a)
                if a in accepted:
                    seen[code]["ended_ok"].add(a)
    if seen:
        out.append("BY CODE (accounts it came up in / of them ended accepted / not accepted)")
        for code, d in sorted(seen.items(), key=lambda kv: -len(kv[1]["accounts"])):
            n, ok = len(d["accounts"]), len(d["ended_ok"])
            out.append(f"  {code:32s} {n:>3} / {ok:>3} / {n - ok:>3}")
    bad = [a for a in cands if a not in accepted]
    if bad:
        out.append("NOT ACCEPTED (rounds; the sentences still refused at the end)")
        for a in bad:
            s = accounts.get(a) or {"repairs": [], "rewrites": []}
            rounds = [r for r in s["repairs"] if not r.get("inference_repair")]
            out.append(f"  {a}: {len(rounds)} round(s) " + " -> ".join("+".join(round_codes(r)) or "?" for r in rounds)
                       + (f"; stopped: {s.get('stop')}" if s.get("stop") else ""))
            left = residual(cands[a], width)
            out.extend(left)
            if not left and rounds:          # an account-level refusal (a required figure, a note left unaddressed) has no sentence to show
                out.append(f"      last refusal: {str(rounds[-1].get('feedback') or '')[:600]}")
            for w in s["rewrites"][:3]:
                if w.get("rewrote"):
                    out.append(f"      step {w.get('step')} the writer's rewrite: " + " | ".join(str(t)[:width] for t in w["rewrote"][:2]))
    cleared = [(a, n, c["judged"]) for a, cand in cands.items() for n, c in enumerate(cand.get("claims") or []) if (c.get("judged") or {}).get("verdict") == "faithful"]
    if cleared:
        out.append(f"READ BY THE AGENT AND CLEARED ({len(cleared)}): refused only for how the prose was read; the reading held against the evidence")
        for a, n, j in cleared:
            out.append(f"  {a}[{n}] {str(cands[a]['claims'][n].get('text') or '')[:width]}")
    confirmed = [(a, n, c["judged"]) for a, cand in cands.items() for n, c in enumerate(cand.get("claims") or []) if (c.get("judged") or {}).get("verdict") == "wrong"]
    if confirmed:
        out.append(f"FOUND WRONG BY THE AGENT ({len(confirmed)})")
        for a, n, j in confirmed:
            out.append(f"  {a}[{n}] {str(cands[a]['claims'][n].get('text') or '')[:width]}")
            out.append(f"      {str(j.get('why') or '')[:260]}")
    fixed = [a for a, s in refused.items() if a in accepted]
    if fixed:
        out.append("REPAIRED (refused, then accepted; the codes of each round)")
        for a in fixed:
            rounds = [r for r in accounts[a]["repairs"] if not r.get("inference_repair")]
            out.append(f"  {a}: " + " -> ".join("+".join(round_codes(r)) or "?" for r in rounds))
    asked = {a: sum(len(r.get("inference_repair") or []) for r in s["repairs"]) for a, s in accounts.items()}
    asked = {a: n for a, n in asked.items() if n}
    if asked:
        left = Counter()
        for a in asked:
            claims = (cands.get(a) or {}).get("claims") or []
            for n, v in enumerate(((cands.get(a) or {}).get("outcome") or {}).get("verdicts") or []):
                if v.get("verdict") == "verified_inference" and n < len(claims):
                    left[a] += 1
        out.append(f"INFERENCE: reasons no note supports, sent back once: {sum(asked.values())} sentence(s) in {len(asked)} account(s); "
                   f"still orange after: {sum(left.values())} sentence(s) in {len(left)} account(s)")
        for a in left:
            claims = cands[a]["claims"]
            for n, v in enumerate(cands[a]["outcome"]["verdicts"]):
                if v.get("verdict") == "verified_inference" and n < len(claims):
                    out.append(f"      {a}[{n}] {str(claims[n].get('text') or '')[:width]}")
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--session", required=True)
    ap.add_argument("--width", type=int, default=140)
    args = ap.parse_args()
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    print("\n".join(pack(json.loads(Path(args.session).read_text(encoding="utf-8")), args.width)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
