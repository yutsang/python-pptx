"""Read-only memory of one finished run, and the tools an agent may ask it.

The user ranked this above everything else: the memory 「肩負了核數和後續
agent chat」. The audit half already works -- audit.jsonl and __run__ hold every
account's phase path, typed defects, repairs with before/after. This is the
chat half: a loader for what a run left on disk, and a small set of
deterministic tools over it, each answer carrying its provenance.

Nothing here calls a model, and nothing crosses a run boundary. The tools are
plain functions so they can be unit-tested and so a caller that is NOT an LLM
(inspect_run.py, a Streamlit panel) can use them directly.

What it reads:

    results.yml        the shipped text and per-stage outputs, plus __run__
                       and __run_health__
    audit.jsonl        one state path per account
    evidence/*.json    the grounding pool each account was graded against
                       and what its Generator was shown (E2)
    data.yml           token counts per stage, if present

The one rule that makes a chat over this safe: a number an agent states must
ground against the SAME pool that graded the deck. `trace_amount` is that
check exposed as a tool, and ask_run.py runs it over the agent's whole answer
before printing it.
"""
from __future__ import annotations

import collections
import json
import os
from typing import Any, Dict, List, Optional

import yaml

from .evidence import AccountEvidence, list_evidence
from .validator import SourceIndex, extract_amount_spans, describe_fact

RUN_STATE_KEY = "__run__"
RUN_HEALTH_KEY = "__run_health__"


class RunMemory:
    """Everything one run wrote, loaded lazily, read-only."""

    def __init__(self, run_folder: str) -> None:
        if not os.path.isdir(run_folder):
            raise FileNotFoundError(run_folder)
        if not os.path.exists(os.path.join(run_folder, "results.yml")):
            present = sorted(os.listdir(run_folder))
            raise FileNotFoundError(
                "%s has no results.yml (the run did not reach finalize). Present: %s"
                % (run_folder, present)
            )
        self.run_folder = run_folder
        self.run_id = os.path.basename(run_folder).replace("run_", "")
        self._results: Optional[Dict[str, Any]] = None
        self._audit: Optional[Dict[str, Dict[str, Any]]] = None
        self._evidence: Optional[Dict[str, AccountEvidence]] = None
        self._data: Optional[Dict[str, Any]] = None

    # -- loading -------------------------------------------------------------

    @property
    def results(self) -> Dict[str, Any]:
        if self._results is None:
            with open(os.path.join(self.run_folder, "results.yml"), encoding="utf-8") as fh:
                self._results = yaml.safe_load(fh) or {}
        return self._results

    @property
    def audit(self) -> Dict[str, Dict[str, Any]]:
        if self._audit is None:
            self._audit = {}
            path = os.path.join(self.run_folder, "audit.jsonl")
            if os.path.exists(path):
                with open(path, encoding="utf-8") as fh:
                    for line in fh:
                        line = line.strip()
                        if not line:
                            continue
                        try:
                            rec = json.loads(line)
                        except Exception:
                            continue
                        self._audit[str(rec.get("account"))] = rec
        return self._audit

    @property
    def evidence(self) -> Dict[str, AccountEvidence]:
        if self._evidence is None:
            self._evidence = list_evidence(self.run_folder)
        return self._evidence

    @property
    def data(self) -> Dict[str, Any]:
        if self._data is None:
            path = os.path.join(self.run_folder, "data.yml")
            self._data = {}
            if os.path.exists(path):
                with open(path, encoding="utf-8") as fh:
                    self._data = yaml.safe_load(fh) or {}
        return self._data

    def accounts(self) -> List[str]:
        return [k for k, v in self.results.items()
                if isinstance(v, dict) and not str(k).startswith("__")]

    def language(self) -> str:
        for ev in self.evidence.values():
            if ev.language:
                return ev.language
        return ""

    def resolve(self, key: str) -> Optional[str]:
        """Exact, then case/space-insensitive, then substring, so a question can
        name the account loosely. Never guesses between two candidates."""
        want = str(key or "").strip()
        names = self.accounts()
        if want in names:
            return want
        low = want.lower().replace(" ", "")
        hits = [n for n in names if str(n).lower().replace(" ", "") == low]
        if len(hits) == 1:
            return hits[0]
        hits = [n for n in names if low and low in str(n).lower().replace(" ", "")]
        return hits[0] if len(hits) == 1 else None

    # -- tools ---------------------------------------------------------------
    #
    # Every tool returns a plain dict. `provenance` says where the numbers came
    # from; `error` names an account that does not exist rather than raising.

    def run_health(self) -> Dict[str, Any]:
        health = dict(self.results.get(RUN_HEALTH_KEY) or {})
        run = self.results.get(RUN_STATE_KEY) or {}
        return {
            "run_id": self.run_id,
            "accounts": len(self.accounts()),
            "health": health,
            "phase_tally": (run.get("state") or {}).get("phase_tally"),
            "degraded_accounts": (run.get("state") or {}).get("degraded_accounts"),
            "defect_codes": run.get("defect_codes"),
            "resumed_from": run.get("resumed_from"),
            "provenance": ["results.yml:__run_health__", "results.yml:__run__"],
        }

    def account_state(self, key: str) -> Dict[str, Any]:
        name = self.resolve(key)
        if name is None:
            return {"error": "no account named %r; accounts: %s" % (key, self.accounts())}
        rec = self.audit.get(name) or {}
        return {
            "account": name,
            "phase": rec.get("phase"),
            "stages_done": rec.get("stages_done"),
            "degraded": rec.get("degraded"),
            "transitions": rec.get("transitions"),
            "defects": rec.get("defects"),
            "repairs": rec.get("repairs"),
            "feedback_retries": rec.get("feedback_retries"),
            "feedback_arbiter": rec.get("feedback_arbiter"),
            "evidence_summary": rec.get("evidence"),
            "provenance": ["audit.jsonl:%s" % name],
        }

    def final_text(self, key: str) -> Dict[str, Any]:
        name = self.resolve(key)
        if name is None:
            return {"error": "no account named %r" % key}
        res = self.results.get(name) or {}
        return {
            "account": name,
            "final": res.get("final"),
            "stages": {k: v for k, v in res.items() if k.startswith("subagent_")},
            "provenance": ["results.yml:%s.final" % name],
        }

    def flagged_clauses(self, key: str) -> Dict[str, Any]:
        name = self.resolve(key)
        if name is None:
            return {"error": "no account named %r" % key}
        reviews = ((self.results.get(name) or {}).get("agent_4_validation") or {}).get("clause_reviews") or []
        out = []
        for i, r in enumerate(reviews):
            if not isinstance(r, dict) or r.get("supported"):
                continue
            out.append({
                "clause_no": i,
                "clause": r.get("clause"),
                "category": r.get("category"),
                "code": r.get("code"),
                "reason": r.get("reason"),
                "expected": r.get("expected"),
                "amounts": r.get("amounts"),
            })
        return {
            "account": name,
            "clauses_reviewed": len(reviews),
            "flagged": out,
            "provenance": ["results.yml:%s.agent_4_validation.clause_reviews" % name],
        }

    def trace_amount(self, key: str, amount: float) -> Dict[str, Any]:
        """Where does this figure come from? The verifier's own answer."""
        name = self.resolve(key)
        if name is None:
            return {"error": "no account named %r" % key}
        ev = self.evidence.get(name)
        if ev is None:
            return {"error": "no evidence file for %r (older run, or it died before finalize)" % name}
        index = ev.source_index()
        try:
            value = float(amount)
        except Exception:
            return {"error": "amount must be a number, got %r" % (amount,)}
        hit = index.matches(value)
        if hit is not None:
            return {"account": name, "amount": value, "grounded": True,
                    "source": describe_fact(hit), "fact": hit,
                    "provenance": ["evidence/%s.json" % name]}
        miss = index.classify_miss(value)
        nearest = miss.get("nearest")
        return {"account": name, "amount": value, "grounded": False,
                "code": miss.get("code"), "expected": miss.get("expected"),
                "nearest": (describe_fact(nearest) if isinstance(nearest, dict) else None),
                "provenance": ["evidence/%s.json" % name]}

    def what_was_shown(self, key: str) -> Dict[str, Any]:
        name = self.resolve(key)
        if name is None:
            return {"error": "no account named %r" % key}
        ev = self.evidence.get(name)
        if ev is None:
            return {"error": "no evidence file for %r" % name}
        kinds = collections.Counter(f.get("kind") for f in ev.facts)
        return {
            "account": name,
            "shown": ev.shown,
            "pool_size": ev.pool_size,
            "pool_by_kind": dict(kinds),
            "dates_allowed": [f.get("value") for f in ev.dates],
            "provenance": ["evidence/%s.json" % name],
        }

    def series(self, key: str) -> Dict[str, Any]:
        name = self.resolve(key)
        facts = (self.results.get(RUN_STATE_KEY) or {}).get("facts") or {}
        ser = (facts.get("series") or {}).get(name) if name else None
        if not ser:
            return {"error": "no series for %r in __run__.facts" % key}
        return {"account": name, "series": ser,
                "statement_type": (facts.get("statement_types") or {}).get(name),
                "periods": (facts.get("periods") or {}).get(name),
                "provenance": ["results.yml:__run__.facts.series.%s" % name]}

    def links(self, key: str) -> Dict[str, Any]:
        name = self.resolve(key)
        facts = (self.results.get(RUN_STATE_KEY) or {}).get("facts") or {}
        edges = [e for e in (facts.get("links") or [])
                 if isinstance(e, dict) and name in (e.get("source"), e.get("target"))]
        return {"account": name, "links": edges,
                "passed": sum(1 for e in edges if e.get("passed")),
                "failed": sum(1 for e in edges if not e.get("passed")),
                "provenance": ["results.yml:__run__.facts.links"]}

    def tokens(self) -> Dict[str, Any]:
        per_stage: Dict[str, Dict[str, int]] = {}
        for acct in (self.data.get("processing_results") or {}).values():
            for agent, rec in (acct or {}).items():
                if not isinstance(rec, dict):
                    continue
                slot = per_stage.setdefault(agent, {"calls": 0, "prompt": 0, "completion": 0})
                slot["calls"] += 1
                slot["prompt"] += int(rec.get("prompt_tokens") or rec.get("estimated_prompt_tokens") or 0)
                slot["completion"] += int(rec.get("completion_tokens") or rec.get("estimated_output_tokens") or 0)
        return {"per_stage": per_stage, "provenance": ["data.yml:processing_results"]}

    # -- the grounding gate over free text -----------------------------------

    def ground_text(self, text: str, keys: Optional[List[str]] = None) -> Dict[str, Any]:
        """Every amount in `text`, grounded against the union of the named
        accounts' pools (all accounts when none are named). This is what makes
        an answer about the run safe to hand on: the same arithmetic that
        graded the deck grades the answer."""
        names = [self.resolve(k) for k in (keys or [])]
        names = [n for n in names if n] or list(self.evidence.keys())
        facts: List[Dict[str, Any]] = []
        for n in names:
            ev = self.evidence.get(n)
            if ev is not None:
                facts.extend(ev.facts)
        index = SourceIndex(facts)
        report = []
        for value, start, end in extract_amount_spans(text or ""):
            hit = index.matches(value)
            report.append({"amount": value, "span": [start, end], "grounded": hit is not None,
                           "source": (describe_fact(hit) if hit else None)})
        return {"accounts": names, "amounts": report,
                "ungrounded": [r for r in report if not r["grounded"]]}


# -- the registry an agent loop reads ----------------------------------------

TOOLS: Dict[str, Dict[str, Any]] = {
    "run_health": {"args": [], "help": "run-level tally: calls, failures, fallback accounts, safe_to_export, phase tally, defect codes"},
    "account_state": {"args": ["key"], "help": "one account's phase path, defects, repairs, retries, arbiter"},
    "final_text": {"args": ["key"], "help": "the shipped bullet and each stage's output"},
    "flagged_clauses": {"args": ["key"], "help": "the clauses the verifier flagged, with code, reason and expected figure"},
    "trace_amount": {"args": ["key", "amount"], "help": "where a figure (base units, e.g. 6000000 for 0.06亿) comes from, or the nearest source if it does not"},
    "what_was_shown": {"args": ["key"], "help": "what the Generator prompt contained: sections, components kept/dropped, pool size, allowed dates"},
    "series": {"args": ["key"], "help": "the account's total by period"},
    "links": {"args": ["key"], "help": "cross-account relationships touching the account, passed and failed"},
    "tokens": {"args": [], "help": "token use per stage"},
}


def call_tool(memory: RunMemory, name: str, args: Dict[str, Any]) -> Dict[str, Any]:
    spec = TOOLS.get(name)
    if spec is None:
        return {"error": "unknown tool %r; tools: %s" % (name, sorted(TOOLS))}
    fn = getattr(memory, name)
    try:
        return fn(**{k: args.get(k) for k in spec["args"]})
    except TypeError as exc:
        return {"error": "bad arguments for %s: %s (wants %s)" % (name, exc, spec["args"])}
