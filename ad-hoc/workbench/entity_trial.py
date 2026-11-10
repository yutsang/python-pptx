#!/usr/bin/env python3
"""A whole databook written from zero by the agent, once per model arm, on the same workbook bytes: what it cost and what it wrote.

    python ad-hoc/workbench/entity_trial.py plan --databook <xlsx> [--legacy-run RUN_ID]          free, seconds: nothing is sent
    python ad-hoc/workbench/entity_trial.py run  --databook <xlsx> --scripted                     free: the whole flow, scripted writer
    python ad-hoc/workbench/entity_trial.py run  --databook <xlsx> [--legacy-run RUN_ID] --i-accept-paid-calls \\
            [--arms pro,flash,mixed] [--limit N] [--escalate-after 1] [--max-output 4000]           PAID

Arms. The models are config.yml agent.models: primary is the cheaper one (Flash), escalation the stronger (Pro); --flash and --pro
(provider::model) name them without editing the file.

    pro     the stronger model writes every account
    flash   the cheaper model writes every account
    mixed   the cheaper model writes; after --escalate-after rejected drafts the stronger one repairs what was rejected
    planned mixed, but the stronger model first reads each account's evidence once and writes a plan (points to make, traps)
    seeded  mixed, but Python writes the balance and movement sentences itself and the cheaper model only adds the rest
    reviewed  seeded, and a verified paragraph with something only a reader can judge (a cause, notes, a basis caveat, part-year
            figures; Python decides) is read once by the stronger model for meaning; the cheaper one repairs what it objects to,
            and if that repair does not verify the verified paragraph stands with the objections listed as unresolved
    loop    (opt-in) cheaper then stronger through the tool loop, with no prepared write: the path the 2026-11-06 paid run measured,
            here to show what the prepared write changes

Each arm: a fresh copy of the workbook and a fresh session; Python decides which accounts are in scope (the selection the legacy
pipeline uses); every account is written by WorkSession.draft_entity (prepared write: the controller reads the evidence, the model
writes once, the verifier checks, a rejected claim is repaired alone); one deck is exported from the accepted paragraphs ONLY (no
legacy text fills an account the agent did not finish; each such account blocks the deck by name).

The executive summary band of that deck is the exporter's splice of first sentences (no summary call is made or counted); the
legacy run's numbers include its summary calls. That difference is printed with the results.

Output (all git-ignored, under fdd_utils/output/agent_trial/):
    trial_<ts>.json                 numbers only: per arm, per account (by index), per role; no account names, no text
    trial_<ts>.txt                  the summary printed at the end, to paste back
    trial_quality_<ts>.md           per account: the legacy text and each arm's text side by side, verdicts, what is unresolved,
                                    legacy figures an arm left out. Holds client commentary: keep it local
    trial_<arm>_<ts>[_DRAFT-BLOCKED].pptx and .agent-session.json   the deck and the session behind it

--legacy-run reads that run's commentary for the side-by-side and, from fdd_utils/logs/usage.jsonl, the calls the ledger recorded
for that run's file in that run's time window (model, input, cached input, output, priced from config). Without the ledger on
this machine it says so; nothing is estimated in its place.

No cap unless one is given (--max-calls, --max-tokens, --max-cost: one ceiling over every arm together). Each turn still ends at
the product's own per-turn limits. A bar per arm shows accounts done, tokens and cost so far.
"""
from __future__ import annotations

import argparse
import copy
import json
import shutil
import sys
import tempfile
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ROOT))

import agent_acceptance as A  # noqa: E402

OUT_DIR = ROOT / "fdd_utils" / "output" / "agent_trial"
ARMS = ("pro_lead", "flash_lead", "mixed_lead", "pro", "flash", "mixed", "planned", "routed", "seeded", "reviewed", "checked", "loop")
# the model leads (understands, decides, checks); Python reads, computes and verifies as a tool. pro / flash / mixed are the
# Python-gated prepared write, kept to compare against; seeded / reviewed were measured 2026-11-08 and dropped
DEFAULT_ARMS = "pro_lead,flash_lead,mixed_lead"


# ---------------------------------------------------------------------------------------------------------------------------
# arms
# ---------------------------------------------------------------------------------------------------------------------------

def arm_setup(cfg: Dict[str, Any], name: str, args: argparse.Namespace) -> Dict[str, Any]:
    """Which model writes, which repairs, and how: the arm as a config the session reads (agent.models), so the session's own
    routing builds the roles, with the output allowance the trial gives every role."""
    r = A.roles(cfg)
    cheap, strong = dict(r["primary"]), dict(r["escalation"])
    for sel in (cheap, strong):
        sel["max_output"] = int(args.max_output) if args.max_output else sel.get("max_output")
    if name == "pro_lead":
        return {"primary": strong, "escalation": None, "routing": "single", "prepare": True, "lead": True}
    if name == "flash_lead":
        return {"primary": cheap, "escalation": None, "routing": "single", "prepare": True, "lead": True}
    if name == "mixed_lead":                     # the strong model leads (understands, checks, decides), the cheap one writes
        return {"primary": cheap, "escalation": strong, "routing": "escalate", "prepare": True, "lead": True}
    if name == "pro":
        return {"primary": strong, "escalation": None, "routing": "single", "prepare": True}
    if name == "flash":
        return {"primary": cheap, "escalation": None, "routing": "single", "prepare": True}
    if name in ("mixed", "seeded", "reviewed"):
        return {"primary": cheap, "escalation": strong, "routing": "escalate", "prepare": True, "escalate_after": int(args.escalate_after),
                "seed": name in ("seeded", "reviewed"), "review": name == "reviewed"}
    if name == "planned":                        # the stronger model plans each account once (points, traps); the cheaper writes
        return {"primary": cheap, "escalation": strong, "routing": "escalate", "prepare": True, "escalate_after": int(args.escalate_after),
                "plan": True}
    if name == "routed":                         # the stronger model assigns each account (dispatch.py); the cheaper writes the rest
        return {"primary": cheap, "escalation": strong, "routing": "escalate", "prepare": True, "escalate_after": int(args.escalate_after),
                "dispatch": True}
    if name == "checked":                        # the cheaper writes alone; the stronger reads each verified paragraph once and
        # marks the sentences it objects to (agent.review_meaning: mark), changing nothing; it never takes over a write
        return {"primary": cheap, "escalation": strong, "routing": "escalate", "prepare": True, "escalate_after": 99, "review": "mark"}
    if name == "loop":
        return {"primary": cheap, "escalation": strong, "routing": "escalate", "prepare": False}
    raise ValueError(f"unknown arm {name!r} (arms: {', '.join(ARMS)})")


def arm_config(cfg: Dict[str, Any], setup: Dict[str, Any]) -> Dict[str, Any]:
    out = copy.deepcopy(cfg or {})
    agent = out.setdefault("agent", {})
    models = agent.setdefault("models", {})
    models["primary"] = {k: v for k, v in setup["primary"].items() if v not in (None, "")}
    if setup.get("escalation"):
        models["escalation"] = {k: v for k, v in setup["escalation"].items() if v not in (None, "")}
    agent["routing"] = setup["routing"]
    agent["prepared_write"] = bool(setup["prepare"])
    agent["seed_facts"] = bool(setup.get("seed"))
    agent["review_meaning"] = setup.get("review") if setup.get("review") == "mark" else bool(setup.get("review"))
    agent["lead_write"] = bool(setup.get("lead"))
    agent["plan_first"] = bool(setup.get("plan"))
    return out


class _Bar:
    """tqdm with manual updates (the session drives the loop), or a silent stand-in without tqdm."""

    def __init__(self, total: int, desc: str):
        try:
            from tqdm import tqdm
            self.bar = tqdm(total=total, desc=desc, file=sys.stderr, dynamic_ncols=True, leave=True)
        except Exception:                                               # noqa: BLE001
            self.bar = None

    def step(self, **postfix: Any) -> None:
        if self.bar is not None:
            self.bar.update(1)
            self.bar.set_postfix(**postfix)

    def close(self) -> None:
        if self.bar is not None:
            self.bar.close()


def code_fingerprint() -> Dict[str, Any]:
    """sha256 of the code a trial depends on, line endings normalised: the company folder is not a git checkout, so this is how a
    run says which frozen version it used."""
    import hashlib
    files = sorted(list((ROOT / "fdd_utils" / "agent").glob("*.py"))
                   + [ROOT / "fdd_utils" / "ai" / n for n in ("claims.py", "validator.py", "evidence_contract.py")]
                   + [ROOT / "fdd_utils" / "workbook" / "supporting.py"]          # the tab's other tables are evidence too
                   + [ROOT / "fdd_utils" / "report_lengths.yml"]                   # the writer's LENGTH per account
                   + [HERE / n for n in ("entity_trial.py", "agent_acceptance.py")])
    total = hashlib.sha256()
    for f in files:
        total.update(f.relative_to(ROOT).as_posix().encode())
        total.update(f.read_bytes().replace(b"\r\n", b"\n"))
    return {"sha256": total.hexdigest(), "short": total.hexdigest()[:12], "files": len(files)}


def coverage(source: Any, key: str, candidate: Any, language: str) -> Dict[str, Any]:
    """What the paragraph covers, by Python, for the record (not a gate beyond the balance): the latest balance stated (its figure
    written in a verified claim that cites it), the movement into it stated, which notes it cites, how many detail lines the
    evidence the writer saw left out (the writer and the reviewer saw the same cut)."""
    from fdd_utils.agent import controller as C
    from fdd_utils.agent import tools as T
    view = T.get_account_evidence(source, key)
    ctl = C.AgentController(source, lambda *a: {}, language=language)
    required = ctl._required_refs(view)
    latest = (view.get("periods") or [""])[-1]
    move = next((d["ref"] for d in view.get("derived") or [] if d.get("kind") == "movement" and d.get("valid", True) and d.get("value")
                 and list(d.get("periods") or [])[-1:] == [latest]), None)
    notes = [n["ref"] for n in view.get("notes") or [] if n.get("ref")]
    must = [m["ref"] for m in ctl._must_address(view)]
    out = {"required": required, "balance_stated": None, "movement": move, "movement_stated": None, "notes": len(notes),
           "notes_cited": 0, "notes_not_cited": notes, "components_not_shown": view.get("components_not_shown"),
           "must": len(must), "must_stated": 0, "must_refs": must}
    if candidate is None:
        return out
    claims, outcome = candidate.claims, candidate.outcome
    cited = {r for c in claims for r in list(c.get("cites") or []) + list(c.get("cause_cites") or [])}
    out["balance_stated"] = (not ctl._required_unstated(required, claims, outcome, view)) if required else None
    out["movement_stated"] = (not ctl._required_unstated([move], claims, outcome, view)) if move else None
    out["notes_cited"] = len([n for n in notes if n in cited])
    out["must_stated"] = len([m for m in must if m in cited])
    out["notes_not_cited"] = [n for n in notes if n not in cited]
    return out


def sources_for_review(source: Any, key: str, candidate: Any) -> Dict[str, Any]:
    """What a source review needs without the workbook at hand: each cited fact (period, basis, line, value, sheet and row) and
    note (its text), and the notes the paragraph did not cite, in full. Written only to the local side-by-side file."""
    contract = source.contracts.get(key)
    if contract is None:
        return {}
    cited, seen = [], set()
    for c in (candidate.claims if candidate is not None else []):
        for ref in list(c.get("cites") or []) + list(c.get("cause_cites") or []):
            if ref in seen:
                continue
            seen.add(ref)
            item = contract.fact(ref) or contract.note(ref) or next((d for d in contract.derived if d.get("ref") == ref), None) or {}
            if str(ref).startswith("N"):
                cited.append({"ref": ref, "basis": "note"})          # its text is listed once per case
                continue
            cited.append({"ref": ref, "period": item.get("period") or item.get("periods"), "basis": item.get("basis") or item.get("kind"),
                          "line": item.get("line"), "value": item.get("value"), "where": f"{item.get('sheet') or ''}!{item.get('row') or ''}".strip("!"),
                          "text": (item.get("text") or "")[:300] or None, "formula": item.get("formula")})
    notes = [{"ref": n.get("ref"), "kind": n.get("kind"), "period": n.get("period"), "text": n.get("text")}
             for n in contract.notes if n.get("ref") and n.get("ref") not in seen]
    return {"cited": cited, "notes_not_cited": notes,
            "all_notes": [{"ref": n.get("ref"), "kind": n.get("kind"), "text": n.get("text")} for n in contract.notes if n.get("ref")]}


def run_arm(name: str, setup: Dict[str, Any], cfg: Dict[str, Any], databook: str, workdir: Path, limits: Any, *,
            select: Callable[[Any], List[str]], scripted: Optional[Dict[str, Any]] = None, stamp: str = "", deck: bool = True,
            checkpoint: Optional[Callable[[Dict[str, Any]], None]] = None, overview: str = "both") -> Dict[str, Any]:
    """One arm on its own copy and session. Every account's row is checkpointed as it finishes; an exception ends the arm with
    what it spent so far recorded (it is never lost, and never counted as a success)."""
    from fdd_utils.agent import routing as R
    from fdd_utils.agent import tools as T
    from fdd_utils.agent.flag import agent_budget
    from fdd_utils.agent.workspace import WorkSession
    from fdd_utils.ai import claims as cl
    copy_path = workdir / f"{name}.xlsx"
    shutil.copyfile(databook, copy_path)
    acfg = arm_config(cfg, setup)
    policy = R.Policy(mode=setup["routing"], escalate_after=int(setup.get("escalate_after") or R.Policy().escalate_after))
    kwargs: Dict[str, Any] = {"budget": agent_budget(acfg), "limits": limits, "policy": policy, "config": acfg, "prepare": setup["prepare"]}
    if scripted is not None:
        kwargs.update(model=scripted["primary"], escalation=scripted.get("escalation"), model_name=f"scripted-{name}")
    started_all = time.perf_counter()
    result: Dict[str, Any] = {"arm": name, "routing": setup["routing"], "prepared_write": setup["prepare"], "seed": bool(setup.get("seed")),
                              "review": bool(setup.get("review")), "escalate_after": setup.get("escalate_after"),
                              "models": {role: {"model": (sel or {}).get("model"), "provider": (sel or {}).get("provider"),
                                                "max_output": (sel or {}).get("max_output")}
                                         for role, sel in (("primary", setup["primary"]), ("escalation", setup.get("escalation"))) if sel},
                              "accounts": 0, "done": 0, "ok": 0, "per_account": [], "_texts": {}}
    work = None
    try:
        work = WorkSession.open(str(copy_path), **kwargs)
        if scripted is not None and scripted.get("holder") is not None:
            scripted["holder"]["source"] = work.source
        books = work.shared
        accounts = select(work)
        ranked = sorted(work.source.contracts)
        result.update(accounts=len(accounts), order=list(accounts), sorted_index={k: ranked.index(k) for k in accounts},
                      in_scope={k: work.source.in_scope(k) for k in accounts})
        if setup.get("dispatch"):
            plan = work.dispatch(accounts)
            result["dispatch"] = {"by": plan["by"], "error": plan["error"], "pro": len(plan["pro"]), "of": len(accounts),
                                  "pro_index": sorted(ranked.index(k) for k in plan["pro"])}
            result["_dispatch"] = {"pro": plan["pro"], "why": plan["why"]}
        marks, clock = [books.calls], [time.perf_counter()]
        per: List[Dict[str, Any]] = result["per_account"]
        bar = _Bar(len(accounts), f"{name}")

        def on_account(key: str, turn: Any) -> None:
            d = A._books_delta(books, marks[-1])
            pr = d.get("provider_reported") or {}
            routing = (turn.report or {}).get("routing") or {}
            c = work.source.candidates.get(key)
            review = dict((c.review if c is not None else {}) or {})
            row = {"case": accounts.index(key) + 1, "sorted_index": ranked.index(key), "ok": bool(turn.ok), "stop_reason": turn.stop_reason,
                   "stop_class": routing.get("stop_class"), "mechanical": ({k: c.outcome.get(k) for k in ("verified", "inference", "failed", "unverified")}
                                                                             | {"validated": bool(c.validated)}) if c is not None else None,
                   "review_status": review.get("status"), "review_objections": len(review.get("objections") or []) + len(review.get("missing") or []),
                   "flags": _flag_counts((c.flags if c is not None else []) or []),
                   "failure_codes": sorted({code for r in refusals(c) for code in r["codes"]}),
                   "coverage": {k: v for k, v in coverage(work.source, key, c, work.source.language).items()
                                if k not in ("required", "notes_not_cited", "must_refs")},
                   "unresolved": len(turn.unresolved or []), "calls": d["calls"], "in": pr.get("prompt", 0), "cached": pr.get("cached_prompt", 0),
                   "out": pr.get("completion", 0), "charged": d["charged_tokens"], "cost": d.get("cost"), "escalations": d["escalations"],
                   "review_calls": ((d.get("by_purpose") or {}).get("review") or {}).get("calls", 0),
                   "review_cost": ((d.get("by_purpose") or {}).get("review") or {}).get("cost"),
                   "fallback": (turn.report or {}).get("prepared_fallback"), "seeded": ((turn.report or {}).get("prepared") or {}).get("seeded"),
                   "seconds": round(time.perf_counter() - clock[-1], 2)}
            per.append(row)
            result["_texts"][key] = {"text": cl.render_claims(c.claims) if c is not None else "", "accepted": bool(c is not None and c.accepted and c.validated),
                                     "unresolved": list(turn.unresolved or []), "review": review,
                                     "sources": sources_for_review(work.source, key, c),
                                     "flags": list((c.flags if c is not None else []) or []),
                                     "refusals": refusals(c) if not (c is not None and c.accepted and c.validated) else [],
                                     "deck_marks": deck_marks(key, c), "plan": (turn.report or {}).get("plan") or {}}
            marks.append(books.calls)
            clock.append(time.perf_counter())
            result.update(done=len(per), ok=sum(p["ok"] for p in per), books=A._books_delta(books, 0))
            if checkpoint is not None:
                checkpoint(dict(row, arm=name))
            bar.step(ok=result["ok"], tokens=f"{books.charged_tokens:,}", cost=f"{books.cost or 0:.3f}")

        review_bar: List[Any] = []

        def on_review(key: str, review: Dict[str, Any]) -> None:   # the marks-only pass after the drafts (agent.review_meaning: mark)
            if not review_bar:
                bar.close()
                review_bar.append(_Bar(len(accounts), f"{name} review"))
            review_bar[0].step(status=review.get("status"), cost=f"{books.cost or 0:.3f}")

        try:
            done = work.draft_entity(accounts, on_account=on_account, on_review=on_review)
            result["stopped"] = next((r.get("error") or "" for r in done["rows"] if r.get("stop_reason") == "budget_session"), "")
        finally:
            bar.close()
            for b in review_bar:
                b.close()
        if done.get("review"):
            # the review ran after every draft was written: its state is read now, not when each account finished
            for row in per:
                c = work.source.candidates.get(accounts[row["case"] - 1])
                review = dict((c.review if c is not None else {}) or {})
                row.update(review_status=review.get("status"),
                           review_objections=len(review.get("objections") or []) + len(review.get("missing") or []),
                           flags=_flag_counts((c.flags if c is not None else []) or []))
        if overview != "none" and not result.get("stopped"):
            # the summary band from the finished paragraphs; with two models, the cheaper one's is written too, for comparison,
            # and the stronger one's (written last) is what the deck carries
            roles = list(work._router().roles)
            order = (["primary", "escalation"] if overview == "both" and "escalation" in roles else
                     ["escalation" if overview == "pro" and "escalation" in roles else "primary"])
            result["_overview"], result["overview"] = {}, {}
            for role in order:
                before = A._books_delta(work.shared, 0)
                written = work.write_overview(role, accounts)
                after = A._books_delta(work.shared, 0)
                label = "pro" if role == "escalation" else ("pro" if name in ("pro", "pro_lead") else "flash")
                result["_overview"][label] = {s_: e["text"] for s_, e in written.items()}
                result["overview"][label] = {"cost": round((after.get("cost") or 0) - (before.get("cost") or 0), 4),
                                             "calls": after["calls"] - before["calls"],
                                             "sections": {s_: {"chars": len(e["text"]), "kept": len(e["claims"]), "dropped": e["dropped"],
                                                               "error": e["error"]} for s_, e in written.items()}}
        if deck:
            OUT_DIR.mkdir(parents=True, exist_ok=True)
            exported = work.export(out_path=str(OUT_DIR / f"trial_{name}_{stamp}.pptx"), first_generation=True)
            data = exported.get("data") or {}
            findings: Dict[str, int] = {}
            for f in (data.get("delivery") or {}).get("findings") or []:
                findings[f.get("code")] = findings.get(f.get("code"), 0) + 1
            result.update(deck=_rel(data.get("path") or ""), export_ok=bool(exported.get("ok")), export_error=(exported.get("error") or {}).get("code"),
                          delivery=(data.get("delivery") or {}).get("status"), findings=findings)
        # drafts the gate did not accept, carried in the deck with their rejected sentences marked (export.py, ACCOUNT_NOT_ACCEPTED)
        marked = {"accounts": 0, "red": 0, "orange": 0}
        for key in T.drafts_to_mark(work.source):
            marked["accounts"] += 1
            for v in work.source.candidates[key].outcome.get("verdicts") or []:
                if v.get("verdict") in ("failed", "unverified"):
                    red = any(f.get("severity") == "failed" and f.get("dimension") in cl._RED_DIMENSIONS for f in v.get("findings") or [])
                    marked["red" if red else "orange"] += 1
        result["marked"] = marked
    except Exception as exc:                                            # noqa: BLE001
        import traceback
        result["error"] = f"{type(exc).__name__}: {str(exc)[:300]}"
        result["traceback_tail"] = traceback.format_exc().splitlines()[-6:]
    finally:
        if work is not None:
            result["books"] = A._books_delta(work.shared, 0)
            try:
                result["session_file"] = _rel(work.save(str(OUT_DIR / f"trial_{name}_{stamp}.agent-session.json")))
            except Exception as exc:                                    # noqa: BLE001
                result["session_file_error"] = f"{type(exc).__name__}: {str(exc)[:160]}"
        result["wall_seconds"] = round(time.perf_counter() - started_all, 1)
    return result


def deck_marks(account: str, candidate: Any) -> List[Dict[str, Any]]:
    """The sentences the deck colours, with the colour and the reason, as the export decides them (claims.to_legacy_result):
    what a person is actually shown, which the state's flags alone are not (most flags no longer colour a sentence)."""
    from fdd_utils.ai import claims as cl
    if candidate is None or not candidate.claims or not (candidate.outcome or {}).get("verdicts"):
        return []
    reviews = cl.to_legacy_result(account, candidate.claims, candidate.outcome, flags=candidate.flags)["agent_4_validation"]["clause_reviews"]
    colour = {"hallucination": "red", "reasoning": "orange"}
    return [{"claim": n, "colour": colour[r["category"]], "text": r["clause"], "why": str(r.get("reason") or "")[:200]}
            for n, r in enumerate(reviews) if r["category"] in colour]


def refusals(candidate: Any) -> List[Dict[str, Any]]:
    """Why the verifier refused each claim it did not pass: the claim's text and every failed / review finding. Without this a
    failed case said only how many claims failed (the 2026-11-08 whole-databook run could not be diagnosed from its report)."""
    if candidate is None:
        return []
    out = []
    for n, v in enumerate(candidate.outcome.get("verdicts") or []):
        if v.get("verdict") in ("verified", "verified_inference"):
            continue
        bad = [f for f in v.get("findings") or [] if f.get("severity") in ("failed", "review")]
        out.append({"claim": n, "text": str(v.get("text") or "")[:300], "codes": sorted({f.get("code") for f in bad}),
                    "messages": [f"{f.get('code')}: {str(f.get('message') or '')[:300]}" for f in bad]})
    return out


def _flag_counts(flags: List[Dict[str, Any]]) -> Dict[str, int]:
    out: Dict[str, int] = {}
    for f in flags:
        out[f.get("kind", "?")] = out.get(f.get("kind", "?"), 0) + 1
    return out


def _rel(path: str) -> Optional[str]:
    if not path:
        return None
    try:
        return str(Path(path).resolve().relative_to(ROOT))
    except ValueError:
        return str(path)


# ---------------------------------------------------------------------------------------------------------------------------
# a scripted writer (free mode): proves the flow, the books and the deck; says nothing about how a model writes
# ---------------------------------------------------------------------------------------------------------------------------

def _writer_view(user: str) -> Optional[Dict[str, Any]]:
    if "EVIDENCE for" not in user:
        return None
    try:
        return json.loads(user.split("EVIDENCE for", 1)[1].split("\n", 1)[1].split("\n\n", 1)[0])
    except (ValueError, IndexError):
        return None


def _balance_claim(view: Dict[str, Any], wrong_by: float = 0.0) -> Optional[Dict[str, Any]]:
    from flow_helpers import date_text
    periods = view.get("periods") or []
    if not periods:
        return None
    latest = periods[-1]
    total = next((t for t in view.get("totals") or [] if t.get("period") == latest and t.get("basis") == "authority_adjusted"), None)
    if total is None:
        return None
    wan = abs(float(total.get("yuan") or 0.0)) / 1e4 + wrong_by
    return {"kind": "amount", "periods": [latest], "cites": [total["ref"]], "text": f"截至{date_text(latest)}，余额{wan:,.1f}万元。"}


class ScriptedWriter:
    """Writes the latest balance from the evidence it is shown. `wrong_first` writes a wrong figure in the first draft of every
    account (so the mixed arm's repair and escalation are exercised)."""

    def __init__(self, label: str, wrong_first: bool = False):
        self.label, self.wrong_first, self.calls = label, wrong_first, 0

    def __call__(self, system: str, user: str, max_tokens: int) -> Dict[str, Any]:
        self.calls += 1
        view = _writer_view(user)
        if "assign each account" in system:
            reply: Dict[str, Any] = {"pro": [], "why": {}}
        elif "executive summary" in system:
            lines = [json.loads(x) for x in user.split("ACCOUNTS:\n", 1)[-1].split("\n\n", 1)[0].splitlines() if x.startswith("{")]
            reply = {"claims": [{"account": x["account"], "kind": "amount", "periods": [x["balance"]["period"]], "cites": [x["balance"]["ref"]],
                                 "text": f"{x['account']}余额{x['balance']['wan']:,.1f}万元。"} for x in lines[:3]]}
        elif "You are the senior" in system:
            reply = {"points": [{"point": "scripted: state the balance", "refs": []}], "cautions": ["scripted: no trap"]}
        elif "You lead the writing" in system:
            reply = {"points": [], "cautions": [], "difficulty": "simple", "writer": "cheap"}
        elif "You check ONE paragraph" in system:
            n = len([line for line in user.splitlines() if line.startswith("  [")])
            reply = {"claims": [{"index": i, "verdict": "ok", "why": "scripted"} for i in range(n)], "missing": []}
        elif "Judge meaning only" in system:
            reply = {"objections": [], "missing": []}
        elif "KEPT (Python wrote these" in user and "REJECTED" not in user:
            reply = {"add": [], "unresolved": [f"{r}: scripted writer defers it" for r in _must_refs(user)]}
        elif view is None:
            reply = {"need_evidence": "scripted writer: no EVIDENCE block"}
        else:
            repairing = "REJECTED" in user
            claim = _balance_claim(view, wrong_by=1.0 if (self.wrong_first and not repairing) else 0.0)
            if claim is None:
                reply = {"draft": {"claims": [], "unresolved": ["no latest balance on the deck's basis"]}}
            elif repairing:
                reply = {"fix": [{"index": 0, "claim": claim}], "add": [] if "REQUIRED_NOT_STATED" not in user else [claim],
                         "unresolved": [f"{r}: scripted writer defers it" for r in _must_refs(user)]}
            else:
                reply = {"draft": {"claims": [claim], "unresolved": [f"{r}: scripted writer defers it" for r in _must_refs(user)]}}
        text = json.dumps(reply, ensure_ascii=False)
        return {"content": text, "prompt_tokens": None, "completion_tokens": None}


def _must_refs(user: str) -> List[str]:
    if "MUST ADDRESS" not in user:
        return []
    block = user.split("MUST ADDRESS", 1)[1].split("\n\n", 1)[0]
    return [line.split()[0] for line in block.splitlines()[1:] if line.strip()]


def scripted_for(name: str) -> Dict[str, Any]:
    if name == "loop":
        holder: Dict[str, Any] = {}
        primary, good = A.scripted_models(lambda: holder["source"], wrong_first=True)
        return {"primary": A.Scripted(primary, "scripted-cheap"), "escalation": A.Scripted(good, "scripted-strong"), "holder": holder}
    if name in ("mixed", "seeded", "reviewed", "mixed_lead", "routed", "planned", "checked"):
        return {"primary": ScriptedWriter("cheap", wrong_first=(name == "mixed")), "escalation": ScriptedWriter("strong")}
    return {"primary": ScriptedWriter(name)}


# ---------------------------------------------------------------------------------------------------------------------------
# the legacy run: its text, and what the ledger recorded for it
# ---------------------------------------------------------------------------------------------------------------------------

def legacy_texts(run_id: str) -> Dict[str, str]:
    if not run_id:
        return {}
    from fdd_utils.financial_common import get_pipeline_result_text
    from fdd_utils.run_identity import load_run, resolve_run_folder
    folder = resolve_run_folder(run_id)
    if not folder:
        return {}
    results = load_run(folder).get("results") or {}
    return {k: str(get_pipeline_result_text(v) or "") for k, v in results.items() if not str(k).startswith("__")}


def legacy_usage(run_id: str, cfg: Dict[str, Any], ledger: Optional[Path] = None) -> Dict[str, Any]:
    """The ledger's calls for that run: the run's workbook file, from the run's start (its id) to when its identity was written,
    a few minutes either side. Provider-reported tokens; priced from config.yml like the agent's books."""
    from fdd_utils.agent import flag as F
    from fdd_utils.run_identity import resolve_run_folder
    if not run_id:
        return {}
    folder = resolve_run_folder(run_id)
    if not folder:
        return {"error": f"no run folder for {run_id}"}
    ledger = ledger or ROOT / "fdd_utils" / "logs" / "usage.jsonl"
    if not ledger.exists():
        return {"error": "no usage ledger on this machine (fdd_utils/logs/usage.jsonl)"}
    try:
        start = datetime.strptime(Path(folder).name.replace("run_", ""), "%Y%m%d_%H%M%S")
    except ValueError:
        return {"error": f"cannot read a start time from {run_id}"}
    ident = {}
    try:
        ident = json.loads((Path(folder) / "run_identity.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        pass
    end = None
    for key in ("written_at",):
        try:
            end = datetime.fromisoformat(str(ident.get(key) or "")[:19])
        except ValueError:
            end = None
    if end is None:
        end = datetime.fromtimestamp((Path(folder) / "results.yml").stat().st_mtime)
    file_name = Path(str(ident.get("workbook") or "")).name
    attribution = (f"calls for file {file_name!r} from the run's start to its identity time, plus or minus five minutes" if file_name else
                   "time window only: the run identity names no workbook, so calls for other files in that window are included")
    lo, hi = (start - timedelta(minutes=5)).isoformat(timespec="seconds"), (end + timedelta(minutes=5)).isoformat(timespec="seconds")
    by_model: Dict[str, Dict[str, Any]] = {}
    rows = 0
    with ledger.open(encoding="utf-8") as fh:
        for line in fh:
            try:
                row = json.loads(line)
            except ValueError:
                continue
            stamp = str(row.get("time") or "")
            if not (lo <= stamp <= hi) or (file_name and str(row.get("file") or "") != file_name):
                continue
            rows += 1
            m = by_model.setdefault(str(row.get("model") or "?"), {"provider": row.get("provider"), "calls": 0, "failed": 0, "response_cache": 0,
                                                                    "in": 0, "cached": 0, "out": 0, "estimated_calls": 0})
            if row.get("token_source") == "response_cache":
                m["response_cache"] += 1
                continue
            if row.get("ok") is False:
                m["failed"] += 1
                m["failed_estimated_in"] = m.get("failed_estimated_in", 0) + int(row.get("estimated_prompt_tokens") or 0)
                continue
            m["calls"] += 1
            if row.get("prompt_tokens") is None:
                m["estimated_calls"] += 1
            m["in"] += int(row.get("prompt_tokens") or 0)
            m["out"] += int(row.get("completion_tokens") or 0)
            m["cached"] += int(row.get("cached_prompt_tokens") or 0)
    for model, m in by_model.items():
        price = F.price_for(cfg, str(m.get("provider") or ""), model)
        m["cost"] = None if price is None else round((m["in"] * price[0] + m["out"] * price[1]) / 1e6, 4)
    total = {k: sum(m.get(k, 0) for m in by_model.values()) for k in ("calls", "failed", "failed_estimated_in", "response_cache", "in", "cached", "out",
                                                                     "estimated_calls")}
    costs = [m["cost"] for m in by_model.values() if m["cost"] is not None]
    total["cost"] = round(sum(costs), 4) if costs and len(costs) == len(by_model) else None
    return {"run": run_id, "file": file_name, "window": [lo, hi], "rows": rows, "by_model": by_model, "total": total, "attribution": attribution,
            "basis": "ledger rows for the run's file in its time window; provider-reported tokens; cached input priced at the full price_in"}


# ---------------------------------------------------------------------------------------------------------------------------
# reports
# ---------------------------------------------------------------------------------------------------------------------------

def _num(n: Any) -> str:
    return "-" if n is None else (f"{n:,}" if isinstance(n, int) else f"{n:,.4f}")


def summary_text(report: Dict[str, Any]) -> str:
    cases = report.get("cases")
    fp = report.get("code_fingerprint") or {}
    lines = [f"TRIAL {report['stamp']}  ({'SCRIPTED: plumbing only, no model wrote anything' if report.get('scripted') else 'real models'})  "
             f"code {fp.get('short')}  workbook {str(report.get('databook_sha256') or '')[:12]}",
             (f"  cases (sorted-contract indices): {cases}" if cases else f"  accounts in scope: {report.get('accounts_in_scope')}"
              + (f" (limited to {report['limit']})" if report.get("limit") else "")),
             f"  declared allowance: {report.get('caps')} (one ceiling over every arm; spent {((report.get('allowance') or {}).get('spent'))})",
             "  arm      ok/acc  calls        in    cached       out   cost(yuan)  tok/acc  yuan/acc  esc  fallbk   secs  deck"]
    for name, res in (report.get("arms") or {}).items():
        if res.get("not_started"):
            lines.append(f"  {name:7s} NOT STARTED: {res['not_started']}")
            continue
        b = res.get("books") or {}
        pr = b.get("provider_reported") or {}
        n = max(res.get("done") or 0, 1)
        tok = (pr.get("prompt", 0) + pr.get("completion", 0)) or b.get("charged_tokens", 0)
        cost = b.get("cost")
        fall = sum(1 for p in res.get("per_account") or [] if p.get("fallback"))
        lines.append(f"  {name:7s} {res.get('ok')}/{res.get('accounts'):<4} {b.get('calls', 0):5d} {pr.get('prompt', 0):9,d} {pr.get('cached_prompt', 0):9,d} "
                     f"{pr.get('completion', 0):9,d}   {_num(cost):>9}  {tok // n:7,d}  {(cost or 0) / n:8.4f}  {b.get('escalations', 0):3d}  {fall:6d} "
                     f"{res.get('wall_seconds', 0):6.0f}  {res.get('delivery') or res.get('export_error') or '-'}"
                     + (f"  STOPPED {res['stopped']}" if res.get("stopped") else "") + (f"  ERROR {res['error']}" if res.get("error") else ""))
        roles = []
        for role, v in (b.get("by_role") or {}).items():
            roles.append(f"{role} {v.get('models_reported') or v.get('model')}: {v['calls']} calls, in {v.get('prompt_provider', 0):,} "
                         f"(cached {v.get('cached_prompt', 0):,}), out {v.get('completion_provider', 0):,}, {_num(v.get('cost'))} yuan")
        lines.append("           " + " | ".join(roles))
        stops: Dict[str, int] = {}
        for p in res.get("per_account") or []:
            stops[str(p.get("stop_reason"))] = stops.get(str(p.get("stop_reason")), 0) + 1
        lines.append(f"           end states {stops}")
        codes: Dict[str, int] = {}
        for p in res.get("per_account") or []:
            for code in p.get("failure_codes") or []:
                codes[code] = codes.get(code, 0) + 1
        if codes:
            lines.append(f"           refused by (accounts per code): {dict(sorted(codes.items(), key=lambda kv: -kv[1]))}")
        must = sum((p.get("coverage") or {}).get("must", 0) for p in res.get("per_account") or [])
        stated = sum((p.get("coverage") or {}).get("must_stated", 0) for p in res.get("per_account") or [])
        deferred = sum((p.get("flags") or {}).get("deferred_note", 0) for p in res.get("per_account") or [])
        flagged = sum(1 for p in res.get("per_account") or [] if any(k in (p.get("flags") or {}) for k in
                                                                       ("deferred_note", "cause_not_proven", "note_wording_unchecked")))
        lines.append(f"           material notes: {stated} stated, {deferred} deferred (flagged), of {must}; accounts with a human check: {flagged}")
        mk = res.get("marked") or {}
        if mk.get("accounts"):
            lines.append(f"           not accepted but in the draft, rejected sentences marked: {mk['accounts']} account(s), "
                         f"{mk['red']} red, {mk['orange']} orange sentence(s)")
        purposes = {k: v for k, v in (b.get("by_purpose") or {}).items() if k in ("understand", "write", "check", "fix")}
        if purposes:
            lines.append("           by step: " + "; ".join(f"{k} {v['calls']} calls, in {v['prompt_provider']:,} out {v['completion_provider']:,}, "
                                                          f"{_num(v.get('cost'))} yuan" for k, v in purposes.items()))
        rev = (b.get("by_purpose") or {}).get("review")
        if res.get("review"):
            statuses: Dict[str, int] = {}
            for p in res.get("per_account") or []:
                statuses[str(p.get("review_status"))] = statuses.get(str(p.get("review_status")), 0) + 1
            lines.append(f"           review: {statuses}" + (f"; {rev['calls']} calls, in {rev['prompt_provider']:,}, out {rev['completion_provider']:,}, "
                                                            f"{_num(rev.get('cost'))} yuan (inside the totals above)" if rev else "; no review call"))
        if res.get("seed"):
            lines.append(f"           Python wrote {sum(int(p.get('seeded') or 0) for p in res.get('per_account') or [])} sentence(s) itself")
        if res.get("dispatch"):
            d = res["dispatch"]
            lines.append(f"           dispatch ({d['by']}{': ' + d['error'] if d.get('error') else ''}): {d['pro']} of {d['of']} accounts to the "
                         f"stronger model (sorted indices {d['pro_index']})")
        for label, o in (res.get("overview") or {}).items():
            lines.append(f"           overview by {label}: {o['calls']} calls, {_num(o.get('cost'))} yuan; "
                         + "; ".join(f"{s_} {x['chars']} chars, {x['kept']} kept, {x['dropped']} dropped{(' ERROR ' + x['error']) if x.get('error') else ''}"
                                     for s_, x in o["sections"].items()))
        if res.get("findings"):
            lines.append(f"           deck findings {res['findings']}")
        if b.get("calls_without_provider_usage"):
            lines.append(f"           {b['calls_without_provider_usage']} call(s) without provider usage (estimated, not in the columns above)")
        if (b.get("cached_input_pricing") or "").startswith(tuple("0123456789")):
            lines.append("           " + b["cached_input_pricing"])
    if cases:
        lines.append("  per case: arm case(index) end | mech v/i/f/u | review | balance movement notes | must stated/total | flags | unresolved | calls tokens yuan secs")
        for name, res in (report.get("arms") or {}).items():
            for p in res.get("per_account") or []:
                m = p.get("mechanical") or {}
                cov = p.get("coverage") or {}
                lines.append(f"    {name:7s} {p['case']:2d}({p['sorted_index']:2d}) {str(p.get('stop_reason')):17s} | "
                             f"{m.get('verified', '-')}/{m.get('inference', '-')}/{m.get('failed', '-')}/{m.get('unverified', '-')} | "
                             f"{str(p.get('review_status') or '-'):19s} | {_yn(cov.get('balance_stated'))} {_yn(cov.get('movement_stated'))} "
                             f"{cov.get('notes_cited', 0)}/{cov.get('notes', 0)}{'' if not cov.get('components_not_shown') else ' +' + str(cov['components_not_shown']) + ' lines not shown'} | "
                             f"{cov.get('must_stated', 0)}/{cov.get('must', 0)} | "
                             f"{','.join(f'{k}:{v}' for k, v in (p.get('flags') or {}).items() if k in ('deferred_note', 'note_amount', 'cause_not_proven', 'note_wording_unchecked', 'detail_caveat', 'basis_caveat')) or '-'} | "
                             f"{p.get('unresolved', 0)} | {p.get('calls')} {p.get('in', 0) + p.get('out', 0):,} {_num(p.get('cost'))} {p.get('seconds')}")
    leg = report.get("legacy_usage") or {}
    if leg.get("total"):
        t = leg["total"]
        lines.append(f"  legacy run {leg['run']} (ledger): {t['calls']} calls, in {t['in']:,} (cached {t['cached']:,}), out {t['out']:,}, "
                     f"{_num(t.get('cost'))} yuan; failed {t['failed']} (est. input {t.get('failed_estimated_in', 0):,}, not costed), "
                     f"response cache {t['response_cache']}")
        lines.append(f"    NOT directly comparable: {leg.get('attribution')}; failed calls are not costed and cached input is at full price "
                     "there; it includes summary and mapping calls the arms do not make")
    elif leg.get("error"):
        lines.append(f"  legacy run: {leg['error']}")
    if report.get("quality"):
        lines.append(f"  side by side with a review form per case (local, holds client text): {report['quality']}")
    return "\n".join(lines)


def _yn(value: Any) -> str:
    return "-" if value is None else ("yes" if value else "NO")


def _numbers(text: str) -> set:
    return A._numbers(text)


REVIEW_FORM = ("來源內容 review（逐項填寫，對照 workbook 原始資料）：\n"
               "- 數字、期間、口徑：\n"
               "- 意思是否與來源一致（有無沒有依據的原因、把餘額當流量、父子項重複、跨期套用）：\n"
               "- 遺漏的重要事實、附註或限制：\n"
               "- unresolved 是否合理：\n"
               "- 結論（可用 / 需修改 / 不可用）：")


def quality_markdown(report: Dict[str, Any], legacy: Dict[str, str], order: List[str]) -> str:
    arms = report.get("arms") or {}
    fp = report.get("code_fingerprint") or {}
    out = [f"# Trial {report['stamp']}", "",
           "Local file: it holds the workbook's commentary. Do not commit it or paste it anywhere public.", "",
           f"Arms: {', '.join(arms)}. Code {fp.get('short')}. Workbook {str(report.get('databook_sha256') or '')[:12]}. "
           f"Legacy run: {report.get('legacy_run') or 'none given'}. Cases: {report.get('cases') or 'all in scope'}.", "",
           "A failed or blocked case is kept: its text, its objections and what it cost are below like any other.", ""]
    for name, res in arms.items():
        if res.get("_dispatch"):
            out.append(f"## {name}: accounts the stronger model took ({len(res['_dispatch']['pro'])})")
            out.extend(f"- {k}: {res['_dispatch']['why'].get(k, '')}" for k in res["_dispatch"]["pro"])
            out.append("")
        for label, texts in (res.get("_overview") or {}).items():
            out.append(f"## {name}: executive summary written by {label}")
            out.extend(f"**{s_}**: {t or '(empty)'}" for s_, t in texts.items())
            out.append("")
    keys: List[str] = list(order)
    for res in arms.values():
        for k in res.get("_texts") or {}:
            if k not in keys:
                keys.append(k)
    for key in keys:
        all_notes = next((t.get("sources", {}).get("all_notes") for res in arms.values() for k, t in (res.get("_texts") or {}).items()
                          if k == key and t.get("sources", {}).get("all_notes")), None)
        rows = {name: next((p for p in res.get("per_account") or [] if res.get("order") and key in res["order"]
                            and p["case"] == res["order"].index(key) + 1), {}) for name, res in arms.items()}
        any_row = next((r for r in rows.values() if r), {})
        out.append(f"## case {any_row.get('case', '?')}: {key} (sorted index {any_row.get('sorted_index', '?')})")
        if all_notes:                                   # once per case, not once per arm (the report was too long to carry)
            out.append("_notes in the evidence_:")
            out.extend(f"- {n['ref']} ({n.get('kind')}): {n.get('text')}" for n in all_notes)
            out.append("")
        old = legacy.get(key, "")
        if legacy:
            out.append(f"**legacy** ({len(old)} chars)\n\n{old or '(no legacy paragraph)'}\n")
        for name, res in arms.items():
            t = (res.get("_texts") or {}).get(key)
            row = rows.get(name) or {}
            if not t:
                out.append(f"**{name}**: no draft" + (f" (arm error: {res['error']})" if res.get("error") else "") + "\n")
                continue
            m, cov, review = row.get("mechanical") or {}, row.get("coverage") or {}, t.get("review") or {}
            out.append(f"**{name}**: end `{row.get('stop_reason')}`; mechanical {m}; review `{review.get('status') or '-'}`; "
                       f"balance stated {_yn(cov.get('balance_stated'))}, movement stated {_yn(cov.get('movement_stated'))}, "
                       f"notes cited {cov.get('notes_cited', 0)}/{cov.get('notes', 0)}, detail lines not shown to the writer {cov.get('components_not_shown')}; "
                       f"calls {row.get('calls')}, tokens in {row.get('in')} (cached {row.get('cached')}) out {row.get('out')}, "
                       f"{_num(row.get('cost'))} yuan, {row.get('seconds')} s"
                       + (f"; fell back to the tool loop: {row.get('fallback')}" if row.get("fallback") else "") + "\n")
            out.append(t["text"] or "(empty)")
            if review.get("text_reviewed") and review["text_reviewed"] != t["text"]:
                out.append(f"\n_the paragraph the reviewer read (before the repair)_: {review['text_reviewed']}")
            if review.get("objections") or review.get("missing"):
                out.append("\n_reviewer objections_: " + "; ".join([f"[{o.get('index')}] {o.get('why')}" for o in review.get("objections") or []]
                                                                  + [f"missing {x.get('ref')}: {x.get('why')}" for x in review.get("missing") or []])
                           + (f" ({review.get('settled')})" if review.get("settled") else ""))
            plan = t.get("plan") or {}
            if plan.get("points") or plan.get("cautions"):
                out.append("\n_plan by the stronger model_: " + " | ".join(p["point"] for p in plan.get("points") or [])
                           + ("\n_traps it named_: " + " | ".join(plan.get("cautions") or []) if plan.get("cautions") else ""))
            out.append("\n_marked in the deck_ (what a person is shown): " + ("none" if not t.get("deck_marks") else ""))
            out.extend(f"- [{m['claim']}] {m['colour'].upper()}: {m['text']}\n  why: {m['why']}" for m in t.get("deck_marks") or [])
            if t.get("refusals"):
                out.append("\n_why the verifier refused (the last draft)_:")
                out.extend(f"- [{r['claim']}] {r['text']}\n  " + "\n  ".join(r["messages"]) for r in t["refusals"])
            if t.get("flags"):
                out.append("\n_state flags (on the record; only the ones under 'marked in the deck' are coloured)_:")
                out.extend(f"- {f.get('kind')}{(' ' + f['ref']) if f.get('ref') else ''}{(' claim ' + str(f['claim'])) if f.get('claim') is not None else ''}: "
                           f"{f.get('why') or ''} {f.get('text') or ''}".rstrip() for f in t["flags"])
            if t.get("unresolved"):
                out.append("\n_unresolved_: " + "; ".join(t["unresolved"]))
            if cov.get("notes_not_cited"):
                out.append(f"\n_notes not cited_: {', '.join(cov['notes_not_cited'][:30])}")
            if old:
                missing = sorted(_numbers(old) - _numbers(t["text"]))
                if missing:
                    out.append(f"\n_legacy figures not in this text_: {', '.join(missing[:20])}")
            src = t.get("sources") or {}
            if src.get("cited"):
                out.append("\n_cited, as read from the workbook_:")
                out.extend(f"- {x['ref']}: {x.get('period')} {x.get('basis')} {x.get('line') or ''} {_num(x.get('value')) if isinstance(x.get('value'), (int, float)) else ''}"
                           f" {('(' + x['where'] + ')') if x.get('where') else ''} {x.get('text') or ''}{(' = ' + x['formula']) if x.get('formula') else ''}".rstrip()
                           for x in src["cited"])
            if src.get("notes_not_cited"):
                out.append("\n_notes this paragraph does not cite_: " + ", ".join(n["ref"] for n in src["notes_not_cited"]))
            out.append("\n" + REVIEW_FORM + "\n")
    return "\n".join(out)


# ---------------------------------------------------------------------------------------------------------------------------
# commands
# ---------------------------------------------------------------------------------------------------------------------------

def _arms(args: argparse.Namespace) -> List[str]:
    wanted = [a.strip() for a in str(args.arms).split(",") if a.strip()]
    bad = [a for a in wanted if a not in ARMS]
    if bad:
        raise SystemExit(f"unknown arm(s) {bad}; arms: {', '.join(ARMS)}")
    return wanted


def _cases(args: argparse.Namespace) -> List[int]:
    return [int(x) for x in str(args.cases or "").split(",") if x.strip()]


def selector(args: argparse.Namespace) -> Callable[[Any], List[str]]:
    """Which accounts an arm writes: the cases by their index in the sorted contract list (the ten-case study's numbering), or
    every account in scope (optionally the first --limit)."""
    cases = _cases(args)

    def select(work: Any) -> List[str]:
        if cases:
            ranked = sorted(work.source.contracts)
            bad = [i for i in cases if not 0 <= i < len(ranked)]
            if bad:
                raise ValueError(f"case index out of range {bad}: this workbook has {len(ranked)} contracts")
            return [ranked[i] for i in cases]
        accounts = work.entity_accounts()
        return accounts[:args.limit] if args.limit else accounts
    return select


def _sha_problem(args: argparse.Namespace, sha: str) -> str:
    want = str(args.expect_sha or "").strip().lower()
    return "" if not want or sha.lower().startswith(want) else f"the workbook's sha256 is {sha[:12]}..., not {want}...: not the same bytes"


def cmd_plan(args: argparse.Namespace) -> int:
    from fdd_utils.agent.session import SourceSession
    cfg = A._config_or_empty()
    sha = _sha(args.databook)
    fp = code_fingerprint()
    print(f"TRIAL PLAN (nothing sent)  code {fp['short']}  workbook {sha[:12]}")
    source = SourceSession.load(args.databook)
    cases = _cases(args)
    if cases:
        ranked = sorted(source.contracts)
        print(f"  cases {cases}: " + ", ".join(f"{i}{'' if source.in_scope(ranked[i]) else ' (out of scope)'}" for i in cases if 0 <= i < len(ranked))
              + (f"; out of range: {[i for i in cases if not 0 <= i < len(ranked)]}" if any(not 0 <= i < len(ranked) for i in cases) else ""))
    else:
        accounts = [k for k in source.contracts if not str(k).startswith("__") and source.in_scope(k)]
        print(f"  accounts in scope: {len(accounts)}" + (f" (the run would take the first {args.limit})" if args.limit else ""))
    from fdd_utils.agent import flag as F
    for name in _arms(args):
        setup = arm_setup(cfg, name, args)
        prices = {role: F.price_for(cfg, (sel or {}).get("provider") or "", (sel or {}).get("model") or "")
                  for role, sel in (("primary", setup["primary"]), ("escalation", setup.get("escalation"))) if sel}
        print(f"  {name:8s} writes: {setup['primary'].get('provider')}::{setup['primary'].get('model')}"
              + (f"; repairs after {setup.get('escalate_after') or 'the policy'} rejection(s): {setup['escalation'].get('provider')}::{setup['escalation'].get('model')}"
                 if setup.get("escalation") else "")
              + f"; prepared write {'on' if setup['prepare'] else 'OFF (tool loop)'}; seed {bool(setup.get('seed'))}; review {bool(setup.get('review'))}; "
              f"max_output {setup['primary'].get('max_output')}; prices {prices}")
    caps = {"max_calls": int(args.max_calls) or None, "max_tokens": int(args.max_tokens) or None, "max_cost": float(args.max_cost) or None}
    print(f"  declared allowance: {caps} (one ceiling over every arm)")
    problems = [] if args.scripted else A.preflight(cfg, A.preview_roles(cfg))
    if _sha_problem(args, sha):
        problems.append(_sha_problem(args, sha))
    leg = legacy_usage(args.legacy_run, cfg) if args.legacy_run else {}
    if leg:
        print("  legacy run: " + (leg.get("error") or f"{leg['total']['calls']} ledger calls in its window, {_num(leg['total'].get('cost'))} yuan; "
                                  f"{leg.get('attribution')}"))
    print("  a run may start: " + ("yes" if not problems else "NO -- " + " | ".join(problems)))
    return 0 if not problems else 1


def _gate() -> str:
    """Run the free offline checks (migration_checks.py) before money is spent; '' when all green, else why not. This machine and
    the one that runs the trial differ (pandas 3's string dtype, Windows, the gitignored config), and a test that fails only
    there has silently cost whole runs before (CLAUDE.md); one command now does both, the gate first."""
    import subprocess
    print("gate: migration_checks.py (free, 1-3 minutes) ...", flush=True)
    done = subprocess.run([sys.executable, str(HERE / "migration_checks.py")], capture_output=True, text=True, encoding="utf-8",
                          errors="replace")
    last = next((line for line in reversed((done.stdout or "").splitlines()) if line.strip()), "")
    print(f"gate: {last}", flush=True)
    if done.returncode != 0 or not last.startswith("ALL GREEN"):
        failing = [line for line in (done.stdout or "").splitlines() if line.startswith("FAIL")]
        return "the offline checks are not all green: " + ("; ".join(failing)[:400] or last or "no summary line")
    return ""


def cmd_run(args: argparse.Namespace) -> int:
    from fdd_utils.agent.routing import Limits
    cfg = A._config_or_empty()
    names = _arms(args)
    sha = _sha(args.databook)
    problems = [] if args.scripted else A.preflight(cfg, A.preview_roles(cfg))
    if not args.scripted and not args.i_accept_paid_calls:
        problems.append("--i-accept-paid-calls was not given (or use --scripted for the free run)")
    if _sha_problem(args, sha):
        problems.append(_sha_problem(args, sha))
    if not problems and not args.scripted and not getattr(args, "skip_gate", False):
        gate = _gate()
        if gate:
            problems.append(gate)
    if problems:
        print("NOT STARTED (nothing was sent):")
        for p in problems:
            print("  -", p)
        return 2
    stamp = time.strftime("%Y%m%d_%H%M%S")
    caps = {"max_calls": int(args.max_calls) or None, "max_tokens": int(args.max_tokens) or None, "max_cost": float(args.max_cost) or None}
    # The cap is one ceiling against a runaway, never a quota per arm: each arm may spend all that the arms before it left.
    # Split evenly (runs=len(names)), a 6-yuan cap gave the stronger model 2 yuan and stopped it at 17 of 26 accounts, so the
    # arm had to be paid for again (2026-11-09). Put the cheap arms first; set the cap well above the estimate.
    allowance = A.Allowance(caps, 1)
    deck = not args.no_deck and not _cases(args)
    report: Dict[str, Any] = {"stamp": stamp, "scripted": bool(args.scripted), "databook_sha256": sha, "limit": args.limit,
                              "cases": _cases(args) or None, "legacy_run": args.legacy_run or None, "arms": {}, "caps": caps,
                              "code_fingerprint": code_fingerprint(), "deck": deck}
    workdir = Path(tempfile.mkdtemp(prefix="entity_trial_"))
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    progress_file = OUT_DIR / f"trial_{stamp}.progress.jsonl"
    legacy = legacy_texts(args.legacy_run)

    def checkpoint(row: Dict[str, Any]) -> None:              # numbers only, one line per finished account, as it finishes
        with progress_file.open("a", encoding="utf-8") as fh:
            fh.write(json.dumps({k: v for k, v in row.items() if k != "coverage"} | {"coverage": {k: v for k, v in (row.get("coverage") or {}).items()
                                                                                                    if k != "notes_not_cited"}},
                                ensure_ascii=False, default=str) + "\n")

    print(f"TRIAL: arms {', '.join(names)}" + (" (scripted, no model calls)" if args.scripted else " (PAID)")
          + f"; code {report['code_fingerprint']['short']}; allowance {caps}", flush=True)
    try:
        for name in names:
            share = allowance.next(name)
            if share is None:
                report["arms"][name] = {"arm": name, "not_started": allowance.given[-1]["not_started"]}
                continue
            setup = arm_setup(cfg, name, args)
            res = run_arm(name, setup, cfg, args.databook, workdir, Limits(**share), select=selector(args),
                          scripted=scripted_for(name) if args.scripted else None, stamp=stamp, deck=deck, checkpoint=checkpoint,
                          overview="none" if args.cases else args.overview)
            allowance.charge(res.get("books"))
            report["arms"][name] = res
            report["accounts_in_scope"] = report.get("accounts_in_scope") or res.get("accounts")
            print(f"  {name}: {res.get('ok')}/{res.get('accounts')} finished" + (f", deck {res.get('delivery') or res.get('export_error')}" if deck else "")
                  + (f"; ERROR {res['error']}" if res.get("error") else ""), flush=True)
            _write_reports(report, legacy, stamp)                       # after every arm: a later failure loses nothing
    finally:
        shutil.rmtree(workdir, ignore_errors=True)
        report["allowance"] = allowance.report()
        report["legacy_usage"] = legacy_usage(args.legacy_run, cfg) if args.legacy_run else {}
        text = _write_reports(report, legacy, stamp)
    print(text)
    print(f"numbers (local, git-ignored, no account names): {(OUT_DIR / f'trial_{stamp}.json').relative_to(ROOT)}; "
          f"per account as it finished: {progress_file.relative_to(ROOT)}")
    # what the review file says in a screenful, so the run's output is what gets pasted back, not the file (several hundred KB)
    quality = report.get("quality")
    if quality and (ROOT / quality).exists():
        from trial_digest import _read, digest
        print("\n".join(digest(_read(ROOT / quality))))
    return 0


def _write_reports(report: Dict[str, Any], legacy: Dict[str, str], stamp: str) -> str:
    order: List[str] = []
    for res in (report.get("arms") or {}).values():
        order = order or list(res.get("order") or [])
    quality = OUT_DIR / f"trial_quality_{stamp}.md"
    quality.write_text(quality_markdown(report, legacy, order), encoding="utf-8")
    report["quality"] = str(quality.relative_to(ROOT))
    numbers = {k: v for k, v in report.items()}
    numbers["arms"] = {n: {k: v for k, v in r.items() if k not in ("_texts", "_overview", "_dispatch", "order", "sorted_index", "in_scope")}
                       | {"in_scope_count": sum(1 for v in (r.get("in_scope") or {}).values() if v)} for n, r in report["arms"].items()}
    (OUT_DIR / f"trial_{stamp}.json").write_text(json.dumps(numbers, ensure_ascii=False, separators=(",", ":"), default=str), encoding="utf-8")
    text = summary_text(report)
    (OUT_DIR / f"trial_{stamp}.txt").write_text(text + "\n", encoding="utf-8")
    return text


def _sha(path: str) -> str:
    from fdd_utils.run_identity import file_sha256
    try:
        return file_sha256(path)
    except OSError:
        return ""


def main(argv: Optional[List[str]] = None) -> int:
    A._utf8()
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    sub = ap.add_subparsers(dest="cmd", required=True)
    for name in ("plan", "run"):
        p = sub.add_parser(name)
        p.add_argument("--databook", required=True)
        p.add_argument("--legacy-run", default="", help="a legacy run of this workbook: its text side by side, its ledger usage")
        p.add_argument("--arms", default=DEFAULT_ARMS, help=f"comma list of {', '.join(ARMS)} (default {DEFAULT_ARMS})")
        p.add_argument("--cases", default="", help="comma list of indices into the sorted contract list (the ten-case study); no deck")
        p.add_argument("--expect-sha", default="", help="refuse to start unless the workbook's sha256 starts with this")
        p.add_argument("--no-deck", action="store_true", help="do not export a deck (always off with --cases)")
        p.add_argument("--overview", default="both", choices=("both", "pro", "flash", "none"),
                       help="the executive summary from the finished paragraphs: both = the cheaper and the stronger model each write one "
                            "(the deck carries the stronger's), on arms with two models (default both; off with --cases)")
        p.add_argument("--limit", type=int, default=0, help="only the first N accounts in scope (a cheap pilot); 0 = all")
        # 2 since 2026-11-08: with 1, nearly every cheap draft went to the stronger model (23 of 26), many of them for refusals the
        # verifier has since stopped making; the cheaper model now gets one repair of its own first
        p.add_argument("--escalate-after", type=int, default=2, help="mixed/seeded/reviewed: rejected drafts before the stronger model repairs (default 2)")
        p.add_argument("--max-output", type=int, default=4000, help="output allowance per call for every role (0 = config's)")
        p.add_argument("--flash", default="", help="provider::model for the cheaper role (default agent.models.primary)")
        p.add_argument("--pro", default="", help="provider::model for the stronger role (default agent.models.escalation)")
        p.add_argument("--scripted", action="store_true", help="a scripted writer instead of models: free, proves the flow only")
        p.add_argument("--max-calls", type=int, default=0, help="0 (default) = no cap; one ceiling over every arm")
        p.add_argument("--max-tokens", type=int, default=0)
        p.add_argument("--max-cost", type=float, default=0.0, help="yuan at config prices; one ceiling over every arm")
        if name == "run":
            p.add_argument("--i-accept-paid-calls", action="store_true")
            p.add_argument("--skip-gate", action="store_true", help="do not run migration_checks.py before a paid run")
    args = ap.parse_args(argv)
    from fdd_utils.agent.flag import parse_role
    A._OVERRIDES.clear()
    if args.flash:
        A._OVERRIDES["primary"] = parse_role(args.flash)
    if args.pro:
        A._OVERRIDES["escalation"] = parse_role(args.pro)
    return {"plan": cmd_plan, "run": cmd_run}[args.cmd](args)


if __name__ == "__main__":
    raise SystemExit(main())
