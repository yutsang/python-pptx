#!/usr/bin/env python3
"""One entry for the company-side acceptance of the agent path: what would run, what it may spend, and a short summary to send back.

    python ad-hoc/workbench/agent_acceptance.py plan  --databook <xlsx> [--base-run <RUN_ID>]       free, seconds: nothing is sent
    python ad-hoc/workbench/agent_acceptance.py probe                                             free: what each role WOULD send
    python ad-hoc/workbench/agent_acceptance.py probe --send                                      PAID, tiny: one call per role
    python ad-hoc/workbench/agent_acceptance.py free  --databook <xlsx> [--base-run <RUN_ID>]       free, minutes: zero model calls
    python ad-hoc/workbench/agent_acceptance.py paid  --databook <xlsx> --base-run <RUN_ID> --i-accept-paid-calls \\
            [--max-calls 90] [--max-tokens 600000] [--max-cost 0]                                  PAID, capped, stops itself

The three configurations come from config.yml (agent.models): flash_only (primary alone), pro_only (the escalation model
alone) and routing (primary, escalating to the other on the code's rule). Each runs the SAME tasks on the SAME account of a
COPY of the workbook, so a task is compared only with itself:

    investigate   a question about the account, then a follow-up
    rewrite       "put it in the report, shorter" (a model edit), accept, export a deck
    update        the same controlled change to one cell of the copy, then a rewrite of the same paragraph from the new figures

A whole entity's first draft is NOT one of them: the agent does not write an entity from nothing; that comparison is the legacy
pipeline's (compare_models.py), printed by `plan`, never run here. Do not compare a task's tokens with a whole entity's.

`free` runs everything that costs nothing: the baseline search (identity hash and per-account fingerprints, reported apart),
the offline gate, the real-workbook flows with scripted models, the three configurations with scripted stand-ins (to check the
books and the plumbing on real frames, not how a model writes), and the reading tools. `paid` refuses to start unless every role
resolves to the provider it names (no silent fallback), the base run describes the workbook, and --i-accept-paid-calls is given.
Results go to fdd_utils/output/agent_acceptance/ (git-ignored); the printed summary carries no workbook text.
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ROOT))

OUT_DIR = ROOT / "fdd_utils" / "output" / "agent_acceptance"
CONFIGS = ("flash_only", "pro_only", "routing")
TASKS = ("investigate", "rewrite", "update")
TURNS_PER_CONFIG = 5                       # investigate 2 + rewrite 1 + update 1, and one spare for a refusal


def progress(iterable, desc: str, total: Optional[int] = None):
    """A tqdm bar on stderr for every long loop here (configurations, tasks, accounts); a plain pass-through without tqdm."""
    try:
        from tqdm import tqdm
        return tqdm(iterable, desc=desc, total=total, file=sys.stderr, dynamic_ncols=True, leave=True)
    except Exception:                                                   # noqa: BLE001
        return iterable


def _utf8() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except (ValueError, OSError):
                pass


def _config() -> Dict[str, Any]:
    from fdd_utils.ai import FDDConfig
    return FDDConfig().config or {}


#: --flash / --pro on the command line (provider::model), so the models already configured for the app can be used without editing config.yml
_OVERRIDES: Dict[str, Dict[str, Any]] = {}


def roles(cfg: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    from fdd_utils.agent import flag as F
    out = {"primary": F.role_config(cfg, "primary"), "escalation": F.role_config(cfg, "escalation")}
    for role, sel in _OVERRIDES.items():
        out[role] = dict(out[role], provider=sel["provider"], model=sel["model"])
    return out


def configurations(cfg: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    r = roles(cfg)
    return {"flash_only": {"primary": r["primary"], "escalation": None, "routing": "single"},
            "pro_only": {"primary": r["escalation"], "escalation": None, "routing": "single"},
            "routing": {"primary": r["primary"], "escalation": r["escalation"], "routing": "escalate"}}


def preview_roles(cfg: Dict[str, Any], send: bool = False) -> Dict[str, Dict[str, Any]]:
    from fdd_utils.agent import probe
    from fdd_utils.agent.controller import AIClientModel
    out = {}
    for name, sel in roles(cfg).items():
        if not sel.get("provider"):
            out[name] = {"error": f"agent.models.{name}.provider is not set in config.yml"}
            continue
        try:
            role = AIClientModel(sel["provider"], sel.get("model") or None, thinking=sel.get("thinking") or "", ledger_stage=f"agent_probe_{name}",
                                 reasoning_effort=sel.get("reasoning_effort") or "", max_output=sel.get("max_output"))
        except Exception as exc:                                       # noqa: BLE001
            out[name] = {"error": f"{type(exc).__name__}: {str(exc)[:200]}"}
            continue
        out[name] = probe.preview(role)
        if send and not out[name].get("error") and not out[name].get("provider_fallback"):
            out[name]["live"] = probe.send(role)
    return out


def preflight(cfg: Dict[str, Any], previews: Dict[str, Dict[str, Any]]) -> List[str]:
    """Why a paid run must not start (empty: it may)."""
    problems = []
    for name, p in previews.items():
        if p.get("error"):
            problems.append(f"{name}: {p['error']}")
        elif p.get("provider_fallback"):
            problems.append(f"{name}: asks for '{p.get('provider_requested')}' but the client would use '{p.get('provider')}' "
                            "(not configured here); a different provider cannot stand in for the gateway")
    r = roles(cfg)
    if r["primary"].get("provider") and r["primary"] == r["escalation"]:
        problems.append("agent.models.primary and agent.models.escalation are the same model: the comparison would compare nothing")
    return problems


# ---------------------------------------------------------------------------------------------------------------------------
# baseline (baseline_pack does the work; identity and fingerprints are reported apart)
# ---------------------------------------------------------------------------------------------------------------------------

def baseline(databook: str, deep: bool = True) -> Dict[str, Any]:
    import baseline_pack as B
    argv = ["--databook", databook] + ([] if deep else ["--no-deep"])
    plan = B.build_plan(B.make_parser().parse_args(argv))
    proof = plan.get("proof") or {}
    return {"status": plan.get("status"), "run": (plan.get("run") or {}).get("id"), "proof_kind": proof.get("kind"),
            "proof_status": proof.get("status"), "problems": [p[:300] for p in plan.get("problems") or []],
            "meaning": {"identity": "the run's identity file names these exact bytes (sha256)",
                        "fingerprints": "every account the run wrote reads the same in the workbook now; that proves the parsed data agree, "
                                        "not that the files are byte-identical",
                        "none": "no run is shown to describe this workbook: the new architecture can still be tested on it, but no "
                                "historical regression is claimed"}.get(proof.get("kind") or "none")}


# ---------------------------------------------------------------------------------------------------------------------------
# the tasks
# ---------------------------------------------------------------------------------------------------------------------------

def _books_delta(books: Any, start: int) -> Dict[str, Any]:
    from fdd_utils.agent.routing import SharedBudget
    part = SharedBudget(books.limits)
    part.records = list(books.records[start:])
    part.escalations = [e for e in books.escalations if int(e.get("after_call") or 0) >= start]
    summary = part.summary()
    summary.pop("limits", None)
    summary["seconds"] = round(sum(float(r.seconds or 0.0) for r in part.records), 2)
    return summary


def choose_account(source: Any, copy_path: str) -> Optional[Tuple[str, str, str, Dict[str, Any], int, float]]:
    """The account the base run commented on that suits all three tasks; without a base run, any account that suits them."""
    from flow_helpers import pick_account
    return pick_account(source, copy_path) or pick_account(source, copy_path, require_base=False)


def _turn_brief(turn: Any) -> Dict[str, Any]:
    routing = (turn.report or {}).get("routing") or {}
    return {"ok": bool(turn.ok), "stop": turn.stop_reason, "stop_class": routing.get("stop_class"), "tools": list(turn.tool_path or []),
            "escalations": [{k: e.get(k) for k in ("reason", "trigger", "to")} for e in routing.get("escalations") or []],
            "unresolved": len(turn.unresolved or []), "verification": dict(turn.verification or {})}


def run_configuration(name: str, setup: Dict[str, Any], databook: str, base_run: str, workdir: Path, limits: Any,
                      scripted: Optional[Dict[str, Any]] = None, on_open=None) -> Dict[str, Any]:
    """The three tasks, in order, on this configuration's own copy of the workbook, under its own caps."""
    from fdd_utils.agent import controller as C
    from fdd_utils.agent import routing as R
    from fdd_utils.agent.workspace import WorkSession
    from flow_helpers import bump_cached_values
    copy = workdir / f"{name}.xlsx"
    shutil.copyfile(databook, copy)
    from fdd_utils.agent.flag import agent_budget
    # the per-turn limits the product runs with (agent.max_steps, max_tokens, max_seconds ... in config.yml), so a turn ends as it would in use
    kwargs: Dict[str, Any] = {"budget": agent_budget(), "limits": limits, "policy": R.Policy(mode=setup["routing"])}
    if scripted is not None:
        kwargs.update(model=scripted["primary"], escalation=scripted.get("escalation"), model_name=f"scripted-{name}")
    else:
        kwargs.update(model_type=setup["primary"]["provider"], model_name=setup["primary"].get("model") or None,
                      thinking=setup["primary"].get("thinking") or "", escalation=setup.get("escalation"))
    work = WorkSession.open(str(copy), base_run=base_run or None, **kwargs)
    if on_open is not None:
        on_open(work)
    picked = choose_account(work.source, str(copy))
    result: Dict[str, Any] = {"configuration": name, "routing": setup["routing"], "tasks": {}, "base_run_check": (work.source.base_run_check or {}).get("status")}
    if picked is None:
        result["error"] = "no account suits the tasks (needs a base commentary, two periods, an unreplaced total and a findable cell)"
        return result
    key, last, prev, raw, row, unit = picked
    result["account_index"] = sorted(work.source.contracts).index(key)          # which account, without its name
    books = work.shared

    def task(label: str, fn) -> None:
        start = books.calls
        started = time.perf_counter()
        try:
            briefs = fn()
            entry = {"ran": True, "turns": briefs}
        except R.BudgetExhausted as exc:
            entry = {"ran": False, "stopped": f"budget: {exc}"}
        except Exception as exc:                                       # noqa: BLE001
            entry = {"ran": False, "error": f"{type(exc).__name__}: {str(exc)[:200]}"}
        entry["books"] = _books_delta(books, start)
        entry["wall_seconds"] = round(time.perf_counter() - started, 2)
        result["tasks"][label] = entry

    def investigate():
        a = work.ask(f"{key}截至{last}的余额是多少？较{prev}变动多少？请说明数据来源。")
        b = work.ask("这个变动的幅度是多少？")
        return [_turn_brief(a), _turn_brief(b)]

    def rewrite():
        t = work.report(f"把{key}的分析放入报告，精简为两句", account=key)
        exported = work.export(out_path=str(workdir / f"{name}.pptx")) if t.ok else {"ok": False, "error": {"code": "NOT_DRAFTED"}}
        return [_turn_brief(t) | {"export_ok": bool(exported.get("ok")), "export_error": (exported.get("error") or {}).get("code")}]

    def update():
        delta = 1000.0 if unit == 1.0 else (1.0 if unit == 1e3 else 0.1)
        bumped = bump_cached_values(str(copy), raw["sheet"], row, raw["value"] / unit, delta)
        notice = work.ensure_fresh() or {}
        effects = [e.get("status") for e in notice.get("candidate_effects") or [] if e.get("account") == key]
        t = work.edit(f"工作簿已更新，请根据最新数据重写{key}的段落", key, accept=True)
        return [_turn_brief(t) | {"cells_changed": len(bumped), "draft_after_change": effects}]

    for label, fn in progress((("investigate", investigate), ("rewrite", rewrite), ("update", update)), desc=f"{name} tasks", total=3):
        if any(not e.get("ran") and str(e.get("stopped", "")).startswith("budget") for e in result["tasks"].values()):
            result["tasks"][label] = {"ran": False, "stopped": "an earlier task reached this configuration's cap"}
            continue
        task(label, fn)
    result["books"] = books.summary()
    return result


def run_entity(name: str, setup: Dict[str, Any], databook: str, base_run: str, workdir: Path, limits: Any,
               scripted: Optional[Dict[str, Any]] = None, on_open=None) -> Dict[str, Any]:
    """Every account the base run commented on, rewritten from its evidence by the agent, one edit per account, on a fresh copy and a
    fresh session (nothing remembered from the other tasks), then one deck exported. This is the agent's whole-entity cost with the
    base run's text as the starting point; it is NOT a from-zero generation, and the legacy run it is compared with wrote from zero."""
    from fdd_utils.agent import routing as R
    from fdd_utils.agent.flag import agent_budget
    from fdd_utils.agent.workspace import WorkSession
    copy = workdir / f"{name}_entity.xlsx"
    shutil.copyfile(databook, copy)
    kwargs: Dict[str, Any] = {"budget": agent_budget(), "limits": limits, "policy": R.Policy(mode=setup["routing"])}
    if scripted is not None:
        kwargs.update(model=scripted["primary"], escalation=scripted.get("escalation"), model_name=f"scripted-{name}")
    else:
        kwargs.update(model_type=setup["primary"]["provider"], model_name=setup["primary"].get("model") or None,
                      thinking=setup["primary"].get("thinking") or "", escalation=setup.get("escalation"))
    work = WorkSession.open(str(copy), base_run=base_run or None, **kwargs)
    if on_open is not None:
        on_open(work)
    source, books = work.source, work.shared
    ordered = sorted(source.contracts)
    accounts = [k for k in (source.base_results or {}) if k in source.contracts and not str(k).startswith("__")] or \
               [k for k in ordered if source.in_scope(k)]
    started_all, start_all, per = time.perf_counter(), books.calls, []
    stopped = ""
    bar = progress(accounts, desc=f"entity {name}", total=len(accounts))
    for key in bar:
        start, started = books.calls, time.perf_counter()
        try:
            turn = work.edit(f"根据工作簿中的证据，重写{key}的报告段落", key, accept=True)
            brief = _turn_brief(turn)
        except R.BudgetExhausted as exc:
            stopped = f"budget: {exc}"
            break
        except Exception as exc:                                       # noqa: BLE001
            brief = {"ok": False, "stop_class": "error", "error": f"{type(exc).__name__}: {str(exc)[:160]}"}
        d = _books_delta(books, start)
        per.append({"account_index": ordered.index(key), "ok": bool(brief.get("ok")), "stop_class": brief.get("stop_class"),
                    "calls": d["calls"], "charged_tokens": d["charged_tokens"], "provider": d["provider_reported"], "estimated": d["estimated"],
                    "escalations": d["escalations"], "seconds": round(time.perf_counter() - started, 2), "error": brief.get("error"),
                    "verification": brief.get("verification"), "unresolved": brief.get("unresolved")})
        if hasattr(bar, "set_postfix"):
            bar.set_postfix(tokens=books.charged_tokens, ok=sum(1 for p in per if p["ok"]))
    OUT_DIR.mkdir(parents=True, exist_ok=True)                       # kept (git-ignored) so the deck can be opened and judged
    deck_path = OUT_DIR / f"entity_{name}_{time.strftime('%Y%m%d_%H%M%S')}.pptx"
    exported = work.export(out_path=str(deck_path)) if not stopped else {"ok": False}
    total = _books_delta(books, start_all)
    # the session (drafts, verdicts, unresolved items) and the deck are kept, git-ignored, so the quality can be judged afterwards
    session_file = work.save(str(OUT_DIR / f"entity_{name}_{time.strftime('%Y%m%d_%H%M%S')}.agent-session.json"))
    written = (exported.get("data") or {}).get("path")
    return {"configuration": name, "routing": setup["routing"], "accounts": len(accounts), "done": len(per),
            "ok": sum(1 for p in per if p["ok"]), "stopped": stopped, "books": total, "per_account": per,
            "wall_seconds": round(time.perf_counter() - started_all, 1), "export_ok": bool(exported.get("ok")),
            "deck": str(Path(written).resolve().relative_to(ROOT)) if written else None, "session_file": str(Path(session_file).relative_to(ROOT)),
            "delivery": ((exported.get("data") or {}).get("delivery") or {}).get("findings"),
            "export_error": (exported.get("error") or {}).get("code"), "base_run_check": (source.base_run_check or {}).get("status")}


# ---------------------------------------------------------------------------------------------------------------------------
# scripted stand-ins (free mode): they read the evidence and answer from it, so the plumbing and the books run on real frames
# ---------------------------------------------------------------------------------------------------------------------------

class Scripted:
    """A stand-in model: returns the JSON a policy builds, and reports no usage, so the books hold estimates only."""

    def __init__(self, fn, label: str):
        self.fn, self.label, self.calls = fn, label, 0

    def __call__(self, system: str, user: str, max_tokens: int) -> Dict[str, Any]:
        from test_agent_controller import returned
        self.calls += 1
        return {"content": json.dumps(self.fn(system, user, returned(user)), ensure_ascii=False)}


def scripted_models(get_source, wrong_first: bool = False) -> Tuple[Any, Any]:
    """(primary, escalation) policies that read the evidence and answer from it, in ask and in edit mode. `wrong_first` makes the
    primary misstate the figure twice (a capability failure the gate catches), so the routing path escalates."""
    from flow_helpers import date_text, wan
    from fdd_utils.ai import claims as cl
    state = {"failures": 0}

    def claims_for(source, key, bad=False):
        contract = source.contracts[key]
        view = cl.evidence_view(contract)
        last = contract.periods[-1]
        total = next((t for t in view["totals"] if t["period"] == last and t["basis"] == "authority_adjusted"), None) or view["totals"][-1]
        value = total["yuan"] * (1.5 if bad else 1.0)
        return [{"account": key, "kind": "amount", "periods": [total["period"]], "cites": [total["ref"]],
                 "text": f"截至{date_text(total['period'])}，{key}余额{wan(value)}"}]

    def policy(system, user, results, bad=False):
        source = get_source()
        names = [n for n, _ in results]
        # the account this turn is about: named in the request line (an edit also carries "[account: X; ...]"); the earlier turns the
        # prompt quotes name other accounts and must not be read as this one
        request = next((line for line in user.splitlines() if line.startswith("REQUEST:")), user)
        names = sorted(source.contracts, key=len, reverse=True)
        key = next((a for a in names if f"[account: {a};" in user), None) or next((a for a in names if a in request), None) \
            or next((a for a in names if a in user), None) or next(iter(source.contracts))
        # read the transcript itself: a large evidence result is cut to its size limit in the prompt and no longer parses as JSON
        if "get_account_evidence" not in names and "TOOL get_account_evidence(" not in user and "TASK STATE" not in user \
                and "fetched for you" not in user:
            return {"tool": "get_account_evidence", "args": {"account": key}}
        if "EDIT MODE" in system:
            at = user.rfind("TOOL validate_candidate(")
            if at >= 0 and any(mark in user[at:at + 600] for mark in ('"ok":true', '"ok": true', "ok=True")):   # compact JSON since the prompt diet
                return {"done": {"account": key, "export": False, "summary": "scripted"}}
            return {"tool": "validate_candidate", "args": {"account": key, "claims": claims_for(source, key, bad)}}
        return {"answer": {"claims": claims_for(source, key, bad), "unresolved": [f"{key}：变动原因未能在工作簿中核实"]}}

    def primary(system, user, results):
        bad = wrong_first and state["failures"] < 2 and "TASK STATE" not in user
        action = policy(system, user, results, bad=bad)
        if bad and (action.get("answer") or (action.get("tool") == "validate_candidate")):
            state["failures"] += 1                                     # only a misstated figure counts, not the fetch before it
        return action
    return primary, policy


# ---------------------------------------------------------------------------------------------------------------------------
# commands
# ---------------------------------------------------------------------------------------------------------------------------

def _caps(args: argparse.Namespace) -> Dict[str, Any]:
    """0 (the default) is no cap: the run spends what the work takes, and the books say what that was. A cap that IS declared
    is one ceiling over the whole command (see Allowance), not a figure each run gets afresh."""
    total = {"max_calls": int(args.max_calls) or None, "max_tokens": int(args.max_tokens) or None, "max_cost": float(args.max_cost) or None}
    return {"total": total, "rule": "one ceiling over the whole command: each run draws an even share of what is left, rounded up "
                                    "and never more than is left; what a run leaves passes to the runs after it"}


class Allowance:
    """The declared total, shared by every run of one command (each configuration's tasks, then each entity). Each run is
    opened with its own Limits, so its books stay its own, but the Limits it gets are cut from what the runs before it left:
    the runs together cannot pass the total (a single call can still pass a token cap by the gap between its estimated and
    its counted input; the books say so). Before this the entity runs got fresh caps of their own, so a declared 90 calls
    allowed 120 with the routing entity and 180 with all three."""

    KEYS = ("max_calls", "max_tokens", "max_cost")

    def __init__(self, total: Dict[str, Any], runs: int):
        self.total = {k: total.get(k) for k in self.KEYS}
        self.runs_left = max(1, int(runs))
        self.spent: Dict[str, float] = {"max_calls": 0, "max_tokens": 0, "max_cost": 0.0}
        self.given: List[Dict[str, Any]] = []

    def next(self, label: str) -> Optional[Dict[str, Any]]:
        """The caps for the next run, or None when a declared cap has nothing left (the run is then not started)."""
        share: Dict[str, Any] = {}
        for key in self.KEYS:
            cap = self.total[key]
            if cap is None:
                share[key] = None
                continue
            left = cap - self.spent[key]
            if left <= 0 or (key == "max_calls" and left < 1):
                self.given.append({"run": label, "not_started": f"{key} total {cap} is used up"})
                self.runs_left = max(1, self.runs_left - 1)
                return None
            share[key] = min(left, -(-int(left) // self.runs_left)) if key != "max_cost" else left / self.runs_left
        self.runs_left = max(1, self.runs_left - 1)
        self.given.append({"run": label, "caps": dict(share)})
        return share

    def charge(self, books: Optional[Dict[str, Any]]) -> None:
        books = books or {}
        self.spent["max_calls"] += int(books.get("calls") or 0)
        self.spent["max_tokens"] += int(books.get("charged_tokens") or 0)
        self.spent["max_cost"] += float(books.get("cost") or 0.0)
        if self.given:
            self.given[-1]["spent"] = {"calls": int(books.get("calls") or 0), "charged_tokens": int(books.get("charged_tokens") or 0),
                                       "cost": books.get("cost")}

    def report(self) -> Dict[str, Any]:
        return {"total": self.total, "spent": {"calls": self.spent["max_calls"], "charged_tokens": self.spent["max_tokens"],
                                               "cost": round(self.spent["max_cost"], 4)}, "runs": self.given}


def cmd_plan(args: argparse.Namespace) -> int:
    cfg = _config()
    previews = preview_roles(cfg)
    caps = _caps(args)
    print("AGENT ACCEPTANCE PLAN (nothing was sent)")
    for name, p in previews.items():
        print(f"  {name:10s} provider {p.get('provider')} model {p.get('model') or p.get('configured_model')}  output<= {p.get('max_output')}  "
              f"thinking requested {p.get('thinking_requested') or '-'} sent {p.get('thinking_sent')}"
              + (" (unverified)" if p.get("thinking_unverified") else "") + (f"  extra_body {p.get('extra_body')}" if p.get("extra_body") else "")
              + (f"  FALLBACK from {p.get('provider_requested')}" if p.get("provider_fallback") else "") + (f"  ERROR {p['error']}" if p.get("error") else ""))
    problems = preflight(cfg, previews)
    print("  configurations: flash_only (primary), pro_only (escalation model alone), routing (primary -> escalation on the code's rule)")
    print(f"  tasks per configuration: {', '.join(TASKS)} ({TURNS_PER_CONFIG} turns at most); all on the same account of a copy of the workbook")
    if any(v is not None for v in caps["total"].values()):
        print(f"  caps: total {caps['total']} over the whole command, task runs and entity runs together ({caps['rule']})")
    else:
        print("  caps: none (the run spends what the work takes); each turn still ends at the product's own per-turn limits (agent.max_steps ...)")
    print("  --entity adds: every account the base run commented on, rewritten by the agent on a fresh copy, then one deck "
          f"(configurations: {getattr(args, 'entity_configs', 'routing')})")
    print("  stop conditions: a cap (calls, tokens, cost) -> that configuration stops; a provider error on its first call -> everything stops")
    print("  first full entity (NOT run here; legacy pipeline, compare like with like):")
    r = roles(cfg)
    if r["primary"].get("provider") and r["escalation"].get("provider"):
        print(f"    python ad-hoc/workbench/compare_models.py {args.databook or '<databook.xlsx>'} --all "
              f"--models {r['primary']['provider']}::{r['primary'].get('model')} --baseline {r['escalation']['provider']}::{r['escalation'].get('model')}")
    else:
        print("    (set agent.models.primary and agent.models.escalation in config.yml first)")
    print("  base run: " + (args.base_run or "none given (run `free` to search for one: identity and fingerprints are reported apart)"))
    print("  paid run may start: " + ("yes" if not problems else "NO -- " + " | ".join(problems)))
    return 0 if not problems else 1


def cmd_probe(args: argparse.Namespace) -> int:
    cfg = _config()
    previews = preview_roles(cfg, send=args.send)
    print(json.dumps(previews, ensure_ascii=False, indent=1, default=str))
    if args.send:
        print("PAID: one small call per role was made (see 'live'); the ledger has them under stage agent_probe_<role>.")
    return 0 if not preflight(cfg, previews) else 1


def cmd_free(args: argparse.Namespace) -> int:
    started = time.perf_counter()
    report: Dict[str, Any] = {"mode": "free", "model_calls": 0}
    print("FREE ACCEPTANCE (zero model calls)")
    found = baseline(args.databook, deep=not args.no_deep)
    report["baseline"] = found
    # A run whose identity names these bytes (or whose every account still reads the same) can carry the base commentary for the
    # new-architecture flows even when the current extractor reads some accounts differently (edited_since); that is reported, and no
    # historical regression is claimed from it.
    usable = found["run"] if found["run"] and found["proof_status"] in ("verified", "edited_since") else ""
    base_run = args.base_run or usable
    if found["run"] and found["proof_status"] == "verified":
        report["historical_regression"] = (f"a run describes this workbook (proof: {found['proof_kind']}); pack {found['status']}"
                                           + ("; checked by identity hash only (--no-deep)" if args.no_deep else ""))
    elif found["run"] and found["proof_status"] == "edited_since":
        report["historical_regression"] = ("NOT claimed: the run was written from these bytes, but the current extractor reads some accounts "
                                           "differently (the differences are listed in the baseline problems)")
    else:
        report["historical_regression"] = "NOT claimed: no run is shown to describe this workbook"
    print(f"  baseline: {found['status']}; run {'found' if found['run'] else 'none'}; proof {found['proof_kind']} ({found['proof_status']}); {found['meaning']}")
    print(f"  historical regression: {report['historical_regression']}; base commentary for the flows: {'that run' if base_run else 'none'}")
    for p in found["problems"][:2]:
        print("    problem:", p[:240])
    gate = subprocess.run([sys.executable, str(HERE / "migration_checks.py")], capture_output=True, encoding="utf-8", errors="replace", cwd=str(ROOT))
    last = [l for l in gate.stdout.splitlines() if "GREEN" in l][-1:] or ["(no summary line)"]
    report["offline_gate"] = last[0]
    print("  offline gate:", last[0])
    if base_run:
        env = dict(os.environ, FDD_FLOW_WORKBOOK=args.databook, FDD_FLOW_BASE_RUN=base_run, PYTHONIOENCODING="utf-8")
        flows = subprocess.run([sys.executable, str(HERE / "migration_checks.py"), "--only", "test_workspace_flows"], capture_output=True,
                               encoding="utf-8", errors="replace", cwd=str(ROOT), env=env)
        summary = [l for l in flows.stdout.splitlines() if l.startswith(("ok ", "FAIL", "ALL GREEN", "NOT GREEN")) or l.strip().startswith("x ")]
        report["real_workbook_flows"] = summary
        print("  real-workbook flows (scripted model):", " | ".join(summary)[:400])
    else:
        print("  real-workbook flows: skipped (no base run describes this workbook)")
    from fdd_utils.agent.routing import Limits
    workdir = Path(tempfile.mkdtemp(prefix="agent_accept_"))
    try:
        holder: Dict[str, Any] = {}
        results = {}
        for name in CONFIGS:
            primary_fn, good_fn = scripted_models(lambda: holder["source"], wrong_first=(name == "routing"))
            primary = Scripted(primary_fn, f"scripted-{name}-primary")
            escalation = Scripted(good_fn, f"scripted-{name}-escalation") if name == "routing" else None
            setup = configurations(_config_or_empty())[name]
            results[name] = run_configuration(name, setup, args.databook, base_run, workdir, Limits(max_calls=40),
                                              scripted={"primary": primary, "escalation": escalation},
                                              on_open=lambda work: holder.__setitem__("source", work.source))
        report["scripted_configurations"] = results
        if args.entity:
            primary_fn, good_fn = scripted_models(lambda: holder["source"], wrong_first=True)
            setup = configurations(_config_or_empty())["routing"]
            report["entity"] = {"routing": run_entity("routing", setup, args.databook, base_run, workdir, Limits(),
                                                      scripted={"primary": Scripted(primary_fn, "scripted-entity-primary"),
                                                                "escalation": Scripted(good_fn, "scripted-entity-escalation")},
                                                      on_open=lambda work: holder.__setitem__("source", work.source))}
        for name, res in results.items():
            line = [f"{t}: {'ran' if e.get('ran') else (e.get('stopped') or e.get('error'))}, calls {e['books']['calls']}, "
                    f"est {e['books']['estimated']['prompt']}+{e['books']['estimated']['completion']}"
                    for t, e in res.get("tasks", {}).items()]
            print(f"  scripted {name}: " + ("; ".join(line) if line else res.get("error", "")) +
                  (f"; escalations {res['books']['escalations']}" if res.get("books") else ""))
    finally:
        shutil.rmtree(workdir, ignore_errors=True)
    if report.get("entity"):
        print("\n".join(summary_text({"entity": report["entity"]}).splitlines()[1:]))
        dumped = quality_dump(report["entity"]["routing"], args.databook, base_run)
        print(f"  quality dump (scripted text, for checking the dump itself): {dumped['path']}; {dumped['line']}")
    report["seconds"] = round(time.perf_counter() - started, 1)
    path = _save(report)
    print(f"  ({report['seconds']}s, no model call; full report: {path.relative_to(ROOT)})")
    print("  scripted stand-ins prove the plumbing and the books on real frames; they say nothing about how a real model writes or chooses")
    return 0


def _config_or_empty() -> Dict[str, Any]:
    try:
        return _config()
    except Exception:                                                   # noqa: BLE001
        return {}


def cmd_paid(args: argparse.Namespace) -> int:
    cfg = _config()
    previews = preview_roles(cfg)
    problems = preflight(cfg, previews)
    if not args.i_accept_paid_calls:
        problems.append("--i-accept-paid-calls was not given")
    if not args.base_run:
        problems.append("--base-run is required (the rewrite task exports a deck over the base run's commentary)")
    if problems:
        print("NOT STARTED (nothing was sent):")
        for p in problems:
            print("  -", p)
        return 2
    from fdd_utils.agent.routing import Limits
    caps = _caps(args)
    wanted = (CONFIGS if args.entity_configs == "all" else [c.strip() for c in args.entity_configs.split(",") if c.strip()]) \
        if args.entity else []
    allowance = Allowance(caps["total"], len(CONFIGS) + len([c for c in wanted if c in CONFIGS]))
    report: Dict[str, Any] = {"mode": "paid", "caps": caps, "previews": previews, "configurations": {}}
    workdir = Path(tempfile.mkdtemp(prefix="agent_paid_"))
    print(f"PAID ACCEPTANCE: {len(CONFIGS)} configurations x {len(TASKS)} tasks" + (" + the whole entity" if args.entity else "")
          + f"; caps {caps['total'] if any(v is not None for v in caps['total'].values()) else 'none'}", flush=True)
    try:
        for name, setup in progress(list(configurations(cfg).items()), desc="configurations"):
            share = allowance.next(f"{name} tasks")
            if share is None:
                report["configurations"][name] = {"configuration": name, "not_started": allowance.given[-1]["not_started"]}
                print(f"  {name}: not started ({allowance.given[-1]['not_started']})", flush=True)
                continue
            res = run_configuration(name, setup, args.databook, args.base_run, workdir, Limits(**share))
            allowance.charge(res.get("books"))
            report["configurations"][name] = res
            first = (res.get("tasks") or {}).get("investigate") or {}
            books = first.get("books") or {}
            if books.get("calls") and books.get("failed_calls") == books.get("calls"):
                report["stopped"] = f"{name}: every call of its first task failed at the provider; nothing further was run"
                print("  " + report["stopped"])
                break
            print(f"  {name}: done", flush=True)
        if args.entity and not report.get("stopped"):
            report["entity"] = {}
            for name, setup in configurations(cfg).items():
                if name in wanted:
                    share = allowance.next(f"{name} entity")
                    if share is None:
                        print(f"  entity {name}: not started ({allowance.given[-1]['not_started']})", flush=True)
                        continue
                    report["entity"][name] = run_entity(name, setup, args.databook, args.base_run, workdir, Limits(**share))
                    allowance.charge(report["entity"][name].get("books"))
                    print(f"  entity {name}: done", flush=True)
    finally:
        shutil.rmtree(workdir, ignore_errors=True)
        report["allowance"] = allowance.report()
    path = _save(report)
    print(summary_text(report))
    print(f"full report (local, git-ignored): {path.relative_to(ROOT)}")
    for name, res in (report.get("entity") or {}).items():
        dumped = quality_dump(res, args.databook, args.base_run)
        print(f"quality dump for entity {name} (local, git-ignored; holds the commentary): {dumped['path']}")
        print("  " + dumped["line"])
    return 0


def _cost_text(books: Dict[str, Any]) -> str:
    """The configured-price figure, never called a bill, with how cached input was priced."""
    if books.get("cost") is None:
        return "cost unknown (these models have no price in config)"
    cached = str(books.get("cached_input_pricing") or "")
    note = "cached input at full price" if "full price_in" in cached else ("cached input at price_cached_in" if "price_cached_in" in cached else "")
    return f"cost {books['cost']:.4f} yuan at config prices" + (f" ({note})" if note else "")


def _role_text(role: str, v: Dict[str, Any]) -> str:
    cost = "" if v.get("cost") is None else f"/{v['cost']:.4f}yuan"
    return f"{role}:{v['calls']}c/in {v.get('prompt_provider', 0)} (cached {v.get('cached_prompt', 0)}) out {v.get('completion_provider', 0)}{cost}"


def summary_text(report: Dict[str, Any]) -> str:
    """The short summary to send back: per task, per configuration. No account names, no workbook text."""
    lines = ["SUMMARY (per task; compare a task only with itself)"]
    for task in TASKS:
        lines.append(f"  {task}:")
        for name, res in (report.get("configurations") or {}).items():
            e = (res.get("tasks") or {}).get(task) or {}
            b = e.get("books") or {}
            pr = b.get("provider_reported") or {}
            roles_line = ", ".join(f"{r}:{v['calls']}c/{v['charged_tokens']}t" + (f"/{','.join(v['models_reported'])}" if v.get("models_reported") else "")
                                   for r, v in (b.get("by_role") or {}).items())
            turns = e.get("turns") or []
            lines.append(f"    {name:10s} {'ran' if e.get('ran') else (e.get('stopped') or e.get('error') or 'not run')}; "
                         f"ok {sum(1 for t in turns if t.get('ok'))}/{len(turns)}; calls {b.get('calls', 0)} (failed {b.get('failed_calls', 0)}, "
                         f"retries {sum((b.get('retries') or {}).values())}); provider {pr.get('prompt', 0)}+{pr.get('completion', 0)} "
                         f"(reasoning {pr.get('reasoning', 0)}, cache {pr.get('cached_prompt', 0)}); est {b.get('estimated', {}).get('prompt', 0)}+"
                         f"{b.get('estimated', {}).get('completion', 0)}; {_cost_text(b)}; escalations {b.get('escalations', 0)}; {b.get('seconds', 0)}s; "
                         f"stop {[t.get('stop_class') for t in turns]}; roles {roles_line}")
    al = report.get("allowance") or {}
    if any(v is not None for v in (al.get("total") or {}).values()):
        lines.append(f"  declared total {al['total']} over the whole command; spent {al.get('spent')}; "
                     + "; ".join(f"{g['run']}: {g.get('not_started') or g.get('caps')}" for g in al.get("runs") or []))
    entity = report.get("entity") or {}
    if entity:
        lines.append("  entity (every account rewritten from its evidence, the base run's text as the start; not a from-zero generation):")
        for name, res in entity.items():
            b = res.get("books") or {}
            pr = b.get("provider_reported") or {}
            per = [p for p in res.get("per_account") or [] if p.get("calls")]
            avg = round(sum(p["charged_tokens"] for p in per) / len(per)) if per else 0
            roles_line = ", ".join(_role_text(r, v) for r, v in (b.get("by_role") or {}).items())
            stops: Dict[str, int] = {}
            for p in res.get("per_account") or []:
                stops[str(p.get("stop_class"))] = stops.get(str(p.get("stop_class")), 0) + 1
            lines.append(f"    {name:10s} accounts {res.get('done')}/{res.get('accounts')} ok {res.get('ok')}; calls {b.get('calls', 0)} "
                         f"(failed {b.get('failed_calls', 0)}, retries {sum((b.get('retries') or {}).values())}); "
                         f"provider-reported in {pr.get('prompt', 0)} (of which cached {pr.get('cached_prompt', 0)}) out {pr.get('completion', 0)} "
                         f"(reasoning {pr.get('reasoning', 0)}); {_cost_text(b)}; "
                         f"{b.get('calls_without_provider_usage', 0)} call(s) without provider usage; charged (cap figure) {b.get('charged_tokens', 0)}; "
                         f"per account avg {avg}; escalations {b.get('escalations', 0)}; {res.get('wall_seconds')}s; "
                         f"deck {res.get('deck') if res.get('export_ok') else res.get('export_error')}; stops {stops}; roles {roles_line}"
                         + (f"; STOPPED {res['stopped']}" if res.get("stopped") else ""))
    return "\n".join(lines)


_NUM = __import__("re").compile(r"\d[\d,]*\.?\d*")


def _numbers(text: str) -> set:
    return {n.replace(",", "").rstrip(".") for n in _NUM.findall(text or "") if len(n.replace(",", "")) >= 2}


def _deck_texts(deck: str) -> Tuple[Dict[str, str], Dict[str, List[Dict[str, Any]]]]:
    """{account: the text on the slides}, {account: the claims behind it} from the deck's own lineage (no model, no extraction)."""
    from fdd_utils.pptx import lineage as L
    public, internal = L.read_manifest(deck) or {}, L.read_internal(deck) or {}
    texts: Dict[str, List[str]] = {}
    claims: Dict[str, List[Dict[str, Any]]] = {}
    for e in public.get("elements") or []:
        inner = (internal.get("elements") or {}).get(e["id"]) or {}
        texts.setdefault(e["account"], []).append(str(inner.get("text") or ""))
        claims.setdefault(e["account"], []).extend(inner.get("claims") or [])
    return {a: " ".join(t for t in parts if t) for a, parts in texts.items()}, claims


def quality_dump(entity: Dict[str, Any], databook: str, base_run: str, deck: str = "") -> Dict[str, Any]:
    """Side by side, per account: what the legacy run wrote, what the agent wrote, the claims and how they verified, and the plain
    measures (length, figures the legacy text quotes that the agent text does not). Written to a git-ignored markdown file for a
    person (or the reviewer) to read; the printed line carries numbers only."""
    from fdd_utils.agent.session import SourceSession
    from fdd_utils.financial_common import get_pipeline_result_text
    source = SourceSession.load(databook, base_run=base_run or None)
    ordered = sorted(source.contracts)
    deck = deck or str(ROOT / entity["deck"]) if (deck or entity.get("deck")) else ""
    agent_text, claims = _deck_texts(deck) if deck and os.path.exists(deck) else ({}, {})
    sess = {}
    if entity.get("session_file") and (ROOT / entity["session_file"]).exists():
        sess = json.loads((ROOT / entity["session_file"]).read_text(encoding="utf-8")).get("candidates") or {}
    rows, sections, dropped_total, legacy_chars, agent_chars, with_text = [], [], 0, 0, 0, 0
    for p in entity.get("per_account") or []:
        key = ordered[p["account_index"]] if p["account_index"] < len(ordered) else f"#{p['account_index']}"
        legacy = str(get_pipeline_result_text((source.base_results or {}).get(key)) or "")
        cand = sess.get(key) or {}
        mine = " ".join(c.get("text", "") for c in cand.get("claims") or []) if cand else ""
        on_deck = agent_text.get(key, "")
        written = mine or (on_deck if p.get("ok") else "")
        with_text += 1 if written else 0
        missing = sorted(_numbers(legacy) - _numbers(written)) if written else []
        dropped_total += len(missing)
        legacy_chars, agent_chars = legacy_chars + len(legacy), agent_chars + len(written)
        verdicts = (cand.get("outcome") or {}).get("verdicts") or []
        counts = {v: sum(1 for x in verdicts if x.get("verdict") == v) for v in ("verified", "verified_inference", "unverified", "failed")}
        rows.append(f"| {p['account_index']} | {key} | {p.get('stop_class')} | {p['calls']} | {p['charged_tokens']} | {p['escalations']} | "
                    f"{len(legacy)} | {len(written)} | {len(cand.get('claims') or claims.get(key) or [])} | "
                    f"{counts['verified']}/{counts['verified_inference']}/{counts['unverified']}/{counts['failed']} | {len(missing)} |")
        sections.append(f"### {key}  ({p.get('stop_class')}; {p['calls']} calls, {p['charged_tokens']} tokens)\n\n"
                        f"**legacy run**\n\n{legacy or '(none)'}\n\n**agent**\n\n"
                        f"{written or ('(accepted, but its text is on no slide of the deck)' if p.get('ok') else '(nothing accepted; the deck keeps the legacy text)')}\n\n"
                        + ("**claims**\n\n" + "\n".join(f"- [{c.get('kind')}] {c.get('text')}  (cites {', '.join(c.get('cites') or [])}"
                                                            f"{'; inference' if c.get('inference') else ''})" for c in (cand.get("claims") or claims.get(key) or []))
                           + "\n\n" if (cand.get("claims") or claims.get(key)) else "")
                        + (f"figures in the legacy text the agent text does not have: {', '.join(missing[:20])}\n\n" if missing else "")
                        + (f"unresolved: {p.get('unresolved')}\n\n" if p.get("unresolved") else ""))
    delivery = entity.get("delivery") or []
    if not delivery and deck:
        side = Path(deck + ".delivery.json")
        if side.exists():
            delivery = (json.loads(side.read_text(encoding="utf-8")) or {}).get("findings") or []
    head = [f"# Agent quality dump -- entity {entity.get('configuration')}", "",
            f"accounts {entity.get('done')}/{entity.get('accounts')}, ok {entity.get('ok')}; charged {entity.get('books', {}).get('charged_tokens')} tokens; "
            f"deck {Path(deck).name if deck else (entity.get('deck') or entity.get('export_error'))}", "",
            "## Why the deck is not deliverable" if delivery else "## Delivery", ""]
    head += [f"- [{f.get('severity')}] {f.get('code')} {f.get('account') or ''}: {f.get('message')}" for f in delivery] or ["- nothing blocks it"]
    head += ["", "## Per account", "", "| # | account | stop | calls | tokens | esc | legacy chars | agent chars | claims | verified/inference/unverified/failed | legacy figures missing |",
             "|---|---|---|---|---|---|---|---|---|---|---|"] + rows + ["", "## Side by side", ""]
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    path = OUT_DIR / f"quality_{entity.get('configuration')}_{time.strftime('%Y%m%d_%H%M%S')}.md"
    path.write_text("\n".join(head) + "\n".join(sections), encoding="utf-8")
    codes: Dict[str, int] = {}
    for f in delivery:
        codes[str(f.get("code"))] = codes.get(str(f.get("code")), 0) + 1
    line = (f"accounts {len(rows)}; legacy chars {legacy_chars}, agent chars {agent_chars} "
            f"({round(100 * agent_chars / legacy_chars) if legacy_chars else 0}%); legacy figures missing from the agent text {dropped_total}; "
            f"delivery findings {codes or 'none'}; agent text found for {with_text} account(s)")
    return {"path": str(path.relative_to(ROOT)), "line": line}


def cmd_dump(args: argparse.Namespace) -> int:
    """The quality dump for a paid run already made (nothing is sent): the report names the accounts, the deck holds what was written."""
    report = json.loads(Path(args.report).read_text(encoding="utf-8"))
    entity = (report.get("entity") or {}).get(args.configuration)
    if not entity:
        print(f"the report has no entity run for {args.configuration}")
        return 2
    deck = args.deck
    if not deck and not entity.get("deck"):
        # the deck is written before the report is saved: take the newest entity deck from before the report
        stamp = Path(args.report).stem.split("_", 1)[-1][:15]
        found = sorted(p for p in OUT_DIR.glob(f"entity_{args.configuration}_*.pptx") if p.stem.split("_", 2)[-1][:15] <= stamp)
        deck = str(found[-1]) if found else ""
        print(f"deck: {Path(deck).name if deck else 'none found (pass --deck)'}")
    dumped = quality_dump(entity, args.databook, args.base_run, deck)
    print(f"quality dump (local, git-ignored; holds the commentary): {dumped['path']}")
    print("  " + dumped["line"])
    return 0


def _save(report: Dict[str, Any]) -> Path:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    path = OUT_DIR / f"{report['mode']}_{time.strftime('%Y%m%d_%H%M%S')}.json"
    path.write_text(json.dumps(report, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    return path


def main(argv: Optional[List[str]] = None) -> int:
    _utf8()
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    for name in ("plan", "free", "paid"):
        p = sub.add_parser(name)
        p.add_argument("--databook", required=(name != "plan"), default="")
        p.add_argument("--base-run", default="")
        p.add_argument("--max-calls", type=int, default=0, help="0 (default) = no cap")
        p.add_argument("--max-tokens", type=int, default=0, help="0 (default) = no cap")
        p.add_argument("--max-cost", type=float, default=0.0, help="0 (default) = no cap; needs prices in config")
        p.add_argument("--flash", default="", metavar="PROVIDER::MODEL", help="the first model (default: agent.models.primary)")
        p.add_argument("--pro", default="", metavar="PROVIDER::MODEL", help="the escalation model (default: agent.models.escalation)")
        p.add_argument("--entity", action="store_true", help="also rewrite every account the base run commented on (the whole-entity cost)")
        p.add_argument("--entity-configs", default="routing", help="which configurations run --entity: routing (default), all, or a comma list")
        if name == "free":
            p.add_argument("--no-deep", action="store_true", help="baseline search by identity hash only (skip the fingerprint proof)")
        if name == "paid":
            p.add_argument("--i-accept-paid-calls", action="store_true")
    du = sub.add_parser("dump", help="the side-by-side quality dump of a paid run already made (nothing is sent)")
    du.add_argument("--report", required=True)
    du.add_argument("--databook", required=True)
    du.add_argument("--base-run", default="")
    du.add_argument("--deck", default="")
    du.add_argument("--configuration", default="routing")
    pr = sub.add_parser("probe")
    pr.add_argument("--send", action="store_true", help="PAID: one small real call per role")
    pr.add_argument("--flash", default="", metavar="PROVIDER::MODEL")
    pr.add_argument("--pro", default="", metavar="PROVIDER::MODEL")
    args = ap.parse_args(argv)
    from fdd_utils.agent.flag import parse_role
    _OVERRIDES.clear()
    if getattr(args, "flash", ""):
        _OVERRIDES["primary"] = parse_role(args.flash)
    if getattr(args, "pro", ""):
        _OVERRIDES["escalation"] = parse_role(args.pro)
    return {"plan": cmd_plan, "probe": cmd_probe, "free": cmd_free, "paid": cmd_paid, "dump": cmd_dump}[args.cmd](args)


if __name__ == "__main__":
    sys.exit(main())
