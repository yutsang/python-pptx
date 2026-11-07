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
from typing import Any, Dict, List, Optional

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ROOT))

import agent_acceptance as A  # noqa: E402

OUT_DIR = ROOT / "fdd_utils" / "output" / "agent_trial"
ARMS = ("pro", "flash", "mixed", "seeded", "reviewed", "loop")
DEFAULT_ARMS = "pro,flash,mixed,seeded,reviewed"


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
    if name == "pro":
        return {"primary": strong, "escalation": None, "routing": "single", "prepare": True}
    if name == "flash":
        return {"primary": cheap, "escalation": None, "routing": "single", "prepare": True}
    if name in ("mixed", "seeded", "reviewed"):
        return {"primary": cheap, "escalation": strong, "routing": "escalate", "prepare": True, "escalate_after": int(args.escalate_after),
                "seed": name in ("seeded", "reviewed"), "review": name == "reviewed"}
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
    agent["review_meaning"] = bool(setup.get("review"))
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


def run_arm(name: str, setup: Dict[str, Any], cfg: Dict[str, Any], databook: str, workdir: Path, limits: Any, *, limit: int = 0,
            scripted: Optional[Dict[str, Any]] = None, stamp: str = "") -> Dict[str, Any]:
    from fdd_utils.agent import routing as R
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
    work = WorkSession.open(str(copy_path), **kwargs)
    if scripted is not None and scripted.get("holder") is not None:
        scripted["holder"]["source"] = work.source
    books = work.shared
    accounts = work.entity_accounts()
    if limit:
        accounts = accounts[:limit]
    order = list(work.source.contracts)
    marks = [books.calls]
    clock = [time.perf_counter()]
    per: List[Dict[str, Any]] = []
    bar = _Bar(len(accounts), f"{name}")

    def on_account(key: str, turn: Any) -> None:
        d = A._books_delta(books, marks[-1])
        pr = d.get("provider_reported") or {}
        routing = (turn.report or {}).get("routing") or {}
        per.append({"account_index": order.index(key), "ok": bool(turn.ok), "stop_class": routing.get("stop_class"),
                    "stop_reason": turn.stop_reason, "calls": d["calls"], "in": pr.get("prompt", 0), "cached": pr.get("cached_prompt", 0),
                    "out": pr.get("completion", 0), "charged": d["charged_tokens"], "cost": d.get("cost"), "escalations": d["escalations"],
                    "fallback": (turn.report or {}).get("prepared_fallback"), "unresolved": len(turn.unresolved or []),
                    "seeded": ((turn.report or {}).get("prepared") or {}).get("seeded"),
                    "review": {k: v for k, v in ((turn.report or {}).get("review") or {}).items() if k != "reasons"} or None,
                    "verification": dict(turn.verification or {}), "seconds": round(time.perf_counter() - clock[-1], 2)})
        marks.append(books.calls)
        clock.append(time.perf_counter())
        bar.step(ok=sum(p["ok"] for p in per), tokens=f"{books.charged_tokens:,}", cost=f"{books.cost or 0:.3f}")

    stopped = ""
    done: Dict[str, Any] = {"rows": []}
    try:
        done = work.draft_entity(accounts, on_account=on_account)
        stopped = next((r.get("error") or "" for r in done["rows"] if r.get("stop_reason") == "budget_session"), "")
    finally:
        bar.close()
    unresolved = {r["account"]: r.get("unresolved") or [] for r in done.get("rows") or []}
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    exported = work.export(out_path=str(OUT_DIR / f"trial_{name}_{stamp}.pptx"), first_generation=True)
    session_file = work.save(str(OUT_DIR / f"trial_{name}_{stamp}.agent-session.json"))
    data = exported.get("data") or {}
    findings: Dict[str, int] = {}
    for f in (data.get("delivery") or {}).get("findings") or []:
        findings[f.get("code")] = findings.get(f.get("code"), 0) + 1
    texts = {}
    for key in accounts:
        c = work.source.candidates.get(key)
        if c is not None:
            texts[key] = {"text": cl.render_claims(c.claims), "accepted": bool(c.accepted and c.validated),
                          "verdicts": {k: c.outcome.get(k) for k in ("verified", "inference", "failed", "unverified")},
                          "unresolved": unresolved.get(key) or []}
    books_all = A._books_delta(books, 0)
    return {"arm": name, "routing": setup["routing"], "prepared_write": setup["prepare"], "seed": bool(setup.get("seed")),
            "review": bool(setup.get("review")),
            "models": {role: {"model": (sel or {}).get("model"), "provider": (sel or {}).get("provider"), "max_output": (sel or {}).get("max_output")}
                       for role, sel in (("primary", setup["primary"]), ("escalation", setup.get("escalation"))) if sel},
            "escalate_after": setup.get("escalate_after"), "accounts": len(accounts), "done": len(per), "ok": sum(p["ok"] for p in per),
            "stopped": stopped, "books": books_all, "per_account": per, "wall_seconds": round(time.perf_counter() - started_all, 1),
            "deck": _rel((data.get("path") or "")), "export_ok": bool(exported.get("ok")), "export_error": (exported.get("error") or {}).get("code"),
            "delivery": (data.get("delivery") or {}).get("status"), "findings": findings, "session_file": _rel(session_file),
            "order": order, "_texts": texts}


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
        if "Judge meaning only" in system:
            reply: Dict[str, Any] = {"objections": [], "missing": []}
        elif "KEPT (Python wrote these" in user and "REJECTED" not in user:
            reply = {"add": [], "unresolved": []}
        elif view is None:
            reply = {"need_evidence": "scripted writer: no EVIDENCE block"}
        else:
            repairing = "REJECTED" in user
            claim = _balance_claim(view, wrong_by=1.0 if (self.wrong_first and not repairing) else 0.0)
            if claim is None:
                reply = {"draft": {"claims": [], "unresolved": ["no latest balance on the deck's basis"]}}
            elif repairing:
                reply = {"fix": [{"index": 0, "claim": claim}], "add": [] if "REQUIRED_NOT_STATED" not in user else [claim]}
            else:
                reply = {"draft": {"claims": [claim], "unresolved": []}}
        text = json.dumps(reply, ensure_ascii=False)
        return {"content": text, "prompt_tokens": None, "completion_tokens": None}


def scripted_for(name: str) -> Dict[str, Any]:
    if name == "loop":
        holder: Dict[str, Any] = {}
        primary, good = A.scripted_models(lambda: holder["source"], wrong_first=True)
        return {"primary": A.Scripted(primary, "scripted-cheap"), "escalation": A.Scripted(good, "scripted-strong"), "holder": holder}
    if name in ("mixed", "seeded", "reviewed"):
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
    total = {k: sum(m[k] for m in by_model.values()) for k in ("calls", "failed", "response_cache", "in", "cached", "out", "estimated_calls")}
    costs = [m["cost"] for m in by_model.values() if m["cost"] is not None]
    total["cost"] = round(sum(costs), 4) if costs and len(costs) == len(by_model) else None
    return {"run": run_id, "file": file_name, "window": [lo, hi], "rows": rows, "by_model": by_model, "total": total,
            "basis": "ledger rows for the run's file in its time window; provider-reported tokens; cached input priced at the full price_in"}


# ---------------------------------------------------------------------------------------------------------------------------
# reports
# ---------------------------------------------------------------------------------------------------------------------------

def _num(n: Any) -> str:
    return "-" if n is None else (f"{n:,}" if isinstance(n, int) else f"{n:,.4f}")


def summary_text(report: Dict[str, Any]) -> str:
    lines = [f"WHOLE-DATABOOK TRIAL {report['stamp']}  ({'SCRIPTED: plumbing only, no model wrote anything' if report.get('scripted') else 'real models'})",
             f"  accounts in scope: {report.get('accounts_in_scope')}" + (f" (limited to {report['limit']})" if report.get("limit") else ""),
             "  arm      ok/acc  calls        in    cached       out   cost(yuan)  tok/acc  yuan/acc  esc  fallbk   secs  deck"]
    for name, res in (report.get("arms") or {}).items():
        b = res.get("books") or {}
        pr = b.get("provider_reported") or {}
        n = max(res.get("done") or 0, 1)
        tok = (pr.get("prompt", 0) + pr.get("completion", 0)) or b.get("charged_tokens", 0)
        cost = b.get("cost")
        fall = sum(1 for p in res.get("per_account") or [] if p.get("fallback"))
        lines.append(f"  {name:7s} {res.get('ok')}/{res.get('accounts'):<4} {b.get('calls', 0):5d} {pr.get('prompt', 0):9,d} {pr.get('cached_prompt', 0):9,d} "
                     f"{pr.get('completion', 0):9,d}   {_num(cost):>9}  {tok // n:7,d}  {(cost or 0) / n:8.4f}  {b.get('escalations', 0):3d}  {fall:6d} "
                     f"{res.get('wall_seconds', 0):6.0f}  {res.get('delivery') or res.get('export_error')}"
                     + (f"  STOPPED {res['stopped']}" if res.get("stopped") else ""))
        roles = []
        for role, v in (b.get("by_role") or {}).items():
            roles.append(f"{role} {v.get('models_reported') or v.get('model')}: {v['calls']} calls, in {v.get('prompt_provider', 0):,} "
                         f"(cached {v.get('cached_prompt', 0):,}), out {v.get('completion_provider', 0):,}, {_num(v.get('cost'))} yuan")
        lines.append("           " + " | ".join(roles))
        rev = (b.get("by_purpose") or {}).get("review")
        if rev:
            outcomes: Dict[str, int] = {}
            for p in res.get("per_account") or []:
                o = str(((p.get("review") or {}).get("outcome") or "")).split(":")[0]
                if o:
                    outcomes[o] = outcomes.get(o, 0) + 1
            lines.append(f"           of which review by the stronger model: {rev['calls']} calls, in {rev['prompt_provider']:,}, out {rev['completion_provider']:,}, "
                         f"{_num(rev.get('cost'))} yuan; outcomes {outcomes}")
        if res.get("seed"):
            lines.append(f"           Python wrote {sum(int(p.get('seeded') or 0) for p in res.get('per_account') or [])} sentence(s) itself")
        if res.get("findings"):
            lines.append(f"           deck findings {res['findings']}")
        if b.get("calls_without_provider_usage"):
            lines.append(f"           {b['calls_without_provider_usage']} call(s) without provider usage (estimated, not in the columns above)")
        if (b.get("cached_input_pricing") or "").startswith(tuple("0123456789")):
            lines.append("           " + b["cached_input_pricing"])
    leg = report.get("legacy_usage") or {}
    if leg.get("total"):
        t = leg["total"]
        lines.append(f"  legacy run {leg['run']} (ledger, its time window): {t['calls']} calls, in {t['in']:,} (cached {t['cached']:,}), "
                     f"out {t['out']:,}, {_num(t.get('cost'))} yuan; failed {t['failed']}, response cache {t['response_cache']}; "
                     + "; ".join(f"{m}: {v['calls']} calls {_num(v.get('cost'))} yuan" for m, v in leg["by_model"].items()))
        lines.append("    the legacy run includes its executive-summary and mapping calls; the arms above write commentary only "
                     "(their summary band is spliced, no call)")
    elif leg.get("error"):
        lines.append(f"  legacy run: {leg['error']}")
    if report.get("quality"):
        lines.append(f"  side by side (local, holds client text): {report['quality']}")
    return "\n".join(lines)


def _numbers(text: str) -> set:
    return A._numbers(text)


def quality_markdown(report: Dict[str, Any], legacy: Dict[str, str], order: List[str]) -> str:
    arms = report.get("arms") or {}
    out = [f"# Whole-databook trial {report['stamp']}",
           "", "Local file: it holds the workbook's commentary. Do not commit it or paste it anywhere public.", "",
           f"Arms: {', '.join(arms)}. Legacy run: {report.get('legacy_run') or 'none given'}.", ""]
    keys = []
    for res in arms.values():
        for k in res.get("_texts") or {}:
            if k not in keys:
                keys.append(k)
    for key in sorted(keys, key=lambda k: order.index(k) if k in order else 10 ** 6):
        out.append(f"## {key} (index {order.index(key) if key in order else '?'})")
        old = legacy.get(key, "")
        if legacy:
            out.append(f"**legacy** ({len(old)} chars)\n\n{old or '(no legacy paragraph)'}\n")
        for name, res in arms.items():
            t = (res.get("_texts") or {}).get(key)
            if not t:
                out.append(f"**{name}**: no draft\n")
                continue
            row = next((p for p in res.get("per_account") or [] if order and p["account_index"] == order.index(key)), {})
            status = "accepted" if t["accepted"] else f"NOT ACCEPTED ({row.get('stop_reason') or 'not verified'})"
            missing = sorted(_numbers(old) - _numbers(t["text"])) if old else []
            out.append(f"**{name}** {status}; verdicts {t['verdicts']}; calls {row.get('calls')}, tokens in {row.get('in')} out {row.get('out')}"
                       + (f", fell back to the tool loop: {row.get('fallback')}" if row.get("fallback") else "") + f" ({len(t['text'])} chars)\n")
            out.append(t["text"] or "(empty)")
            if row.get("review"):
                out.append(f"\n_review by the stronger model_: {row['review']}")
            if t.get("unresolved"):
                out.append("\n_unresolved_: " + "; ".join(t["unresolved"]))
            if missing:
                out.append(f"\n_legacy figures not in this text_: {', '.join(missing[:20])}")
            out.append("")
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


def cmd_plan(args: argparse.Namespace) -> int:
    from fdd_utils.agent.session import SourceSession
    cfg = A._config_or_empty()
    print("WHOLE-DATABOOK TRIAL PLAN (nothing sent)")
    source = SourceSession.load(args.databook)
    accounts = [k for k in source.contracts if not str(k).startswith("__") and source.in_scope(k)]
    print(f"  accounts in scope: {len(accounts)}" + (f" (the run would take the first {args.limit})" if args.limit else ""))
    from fdd_utils.agent import flag as F
    for name in _arms(args):
        setup = arm_setup(cfg, name, args)
        prices = {role: F.price_for(cfg, (sel or {}).get("provider") or "", (sel or {}).get("model") or "")
                  for role, sel in (("primary", setup["primary"]), ("escalation", setup.get("escalation"))) if sel}
        print(f"  {name:6s} writes: {setup['primary'].get('provider')}::{setup['primary'].get('model')}"
              + (f"; repairs after {setup.get('escalate_after') or 'the policy'} rejection(s): {setup['escalation'].get('provider')}::{setup['escalation'].get('model')}"
                 if setup.get("escalation") else "")
              + f"; prepared write {'on' if setup['prepare'] else 'OFF (tool loop)'}; max_output {setup['primary'].get('max_output')}; prices {prices}")
    problems = [] if args.scripted else A.preflight(cfg, A.preview_roles(cfg))
    leg = legacy_usage(args.legacy_run, cfg) if args.legacy_run else {}
    if leg:
        print("  legacy run: " + (leg.get("error") or f"{leg['total']['calls']} ledger calls in its window, {_num(leg['total'].get('cost'))} yuan"))
    print("  a run may start: " + ("yes" if not problems else "NO -- " + " | ".join(problems)))
    return 0 if not problems else 1


def cmd_run(args: argparse.Namespace) -> int:
    from fdd_utils.agent.routing import Limits
    cfg = A._config_or_empty()
    names = _arms(args)
    if not args.scripted:
        problems = A.preflight(cfg, A.preview_roles(cfg))
        if not args.i_accept_paid_calls:
            problems.append("--i-accept-paid-calls was not given (or use --scripted for the free run)")
        if problems:
            print("NOT STARTED (nothing was sent):")
            for p in problems:
                print("  -", p)
            return 2
    stamp = time.strftime("%Y%m%d_%H%M%S")
    caps = {"max_calls": int(args.max_calls) or None, "max_tokens": int(args.max_tokens) or None, "max_cost": float(args.max_cost) or None}
    allowance = A.Allowance(caps, len(names))
    report: Dict[str, Any] = {"stamp": stamp, "scripted": bool(args.scripted), "databook_sha256": _sha(args.databook), "limit": args.limit,
                              "legacy_run": args.legacy_run or None, "arms": {}, "caps": caps}
    workdir = Path(tempfile.mkdtemp(prefix="entity_trial_"))
    legacy = legacy_texts(args.legacy_run)
    order: List[str] = []
    print(f"WHOLE-DATABOOK TRIAL: arms {', '.join(names)}" + (" (scripted, no model calls)" if args.scripted else " (PAID)"), flush=True)
    try:
        for name in names:
            share = allowance.next(name)
            if share is None:
                report["arms"][name] = {"arm": name, "not_started": allowance.given[-1]["not_started"]}
                continue
            setup = arm_setup(cfg, name, args)
            res = run_arm(name, setup, cfg, args.databook, workdir, Limits(**share), limit=args.limit,
                          scripted=scripted_for(name) if args.scripted else None, stamp=stamp)
            allowance.charge(res.get("books"))
            report["arms"][name] = res
            report["accounts_in_scope"] = res["accounts"] if not args.limit else report.get("accounts_in_scope") or res["accounts"]
            order = order or list(res.get("order") or [])
            print(f"  {name}: {res['ok']}/{res['accounts']} accepted, deck {res.get('delivery') or res.get('export_error')}", flush=True)
    finally:
        shutil.rmtree(workdir, ignore_errors=True)
        report["allowance"] = allowance.report()
    report["legacy_usage"] = legacy_usage(args.legacy_run, cfg) if args.legacy_run else {}
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    if not order:
        order = sorted({k for res in report["arms"].values() for k in (res.get("_texts") or {})})
    quality = OUT_DIR / f"trial_quality_{stamp}.md"
    quality.write_text(quality_markdown(report, legacy, order), encoding="utf-8")
    report["quality"] = str(quality.relative_to(ROOT))
    numbers = {k: v for k, v in report.items()}
    numbers["arms"] = {n: {k: v for k, v in r.items() if k not in ("_texts", "order")} for n, r in report["arms"].items()}
    (OUT_DIR / f"trial_{stamp}.json").write_text(json.dumps(numbers, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    text = summary_text(report)
    (OUT_DIR / f"trial_{stamp}.txt").write_text(text + "\n", encoding="utf-8")
    print(text)
    print(f"numbers (local, git-ignored, no account names): {(OUT_DIR / f'trial_{stamp}.json').relative_to(ROOT)}")
    return 0


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
        p.add_argument("--limit", type=int, default=0, help="only the first N accounts in scope (a cheap pilot); 0 = all")
        p.add_argument("--escalate-after", type=int, default=1, help="mixed: rejected drafts before the stronger model repairs (default 1)")
        p.add_argument("--max-output", type=int, default=4000, help="output allowance per call for every role (0 = config's)")
        p.add_argument("--flash", default="", help="provider::model for the cheaper role (default agent.models.primary)")
        p.add_argument("--pro", default="", help="provider::model for the stronger role (default agent.models.escalation)")
        p.add_argument("--scripted", action="store_true", help="a scripted writer instead of models: free, proves the flow only")
        p.add_argument("--max-calls", type=int, default=0, help="0 (default) = no cap; one ceiling over every arm")
        p.add_argument("--max-tokens", type=int, default=0)
        p.add_argument("--max-cost", type=float, default=0.0)
        if name == "run":
            p.add_argument("--i-accept-paid-calls", action="store_true")
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
