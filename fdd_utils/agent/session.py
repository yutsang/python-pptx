"""The sources one agent session works against, and the drafts it holds.

A session is a workbook (and its Financials), read once, with an evidence
contract per account. It records the source version it was read at: every tool
result, claim and deck element carries that version, and a session whose
workbook has changed on disk since it was read is stale and says so instead of
answering from figures that are no longer there.

The session holds no model and no prompt. It is the state the tools read and
write; the controller (controller.py) decides what to do with them.
"""
from __future__ import annotations

import hashlib
import json
import os
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

import pandas as pd

from ..ai import evidence_contract as ec
from ..run_identity import build_identity, file_sha256


from ..workbook.visibility import (  # noqa: E402  (re-exported: the agent session's hidden-sheet rule is the shared one)
    HiddenSheetError, filter_visible, find_visible_financials_sheet as find_financials_sheet, refuse_hidden, sheet_states,
)


@dataclass
class Candidate:
    """A draft paragraph for one account, with what verification said of it."""
    account: str
    claims: List[Dict[str, Any]]
    outcome: Dict[str, Any]
    source_version: str
    claims_digest: str
    #: how it came to exist: "edit" (the model wrote it), "promoted" (an answer's verified claims, unchanged)
    origin: str = "edit"
    #: the user has said this goes in the report
    accepted: bool = False
    #: why it no longer stands (a source it rests on changed); empty while it does
    invalid_reason: str = ""
    #: ref -> what it pointed at when the candidate was written: the semantic id, the value, the recipe of a
    #: derived figure, the hash of a note's text. A ref is only a position in this reading of the workbook;
    #: this is how a candidate finds the same fact again, or finds that it is gone, after the sources move.
    cite_ids: Dict[str, Dict[str, Any]] = field(default_factory=dict)
    turn: int = 0
    #: it survived a change of source because nothing it cites moved
    carried: bool = False
    #: what the account looked like, in the periods the draft touches, when it was written (capture_context). Cited
    #: figures are only part of what a paragraph rests on: the other lines of the period, what the checks and
    #: permissions allowed, the remarks on it. Empty: not recorded.
    context: Dict[str, Any] = field(default_factory=dict)
    #: what moved in that context since the user last looked, in words. A draft with entries here still verifies,
    #: but the person who accepted it accepted it against a different account.
    context_changes: List[str] = field(default_factory=list)
    #: the meaning review of this text (controller.py): status not_needed | passed | passed_after_repair | objected | unavailable,
    #: with the reasons it was due and the objections. Empty: no review was configured. Export reads it: an open objection blocks
    #: the deck for this account, a review that did not happen asks for a person.
    review: Dict[str, Any] = field(default_factory=dict)
    #: what a person should check in this text, set by the state when the paragraph was finished (controller._flags): material
    #: notes deferred, figures from a note's free text, causes not proven, basis and detail caveats, what is unresolved
    flags: List[Dict[str, Any]] = field(default_factory=list)

    @property
    def validated(self) -> bool:
        return bool(self.outcome.get("ok")) and not self.invalid_reason


def capture_cite_ids(contract: ec.EvidenceContract, claims: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    """What each cited ref pointed at in this contract (see Candidate.cite_ids)."""
    out: Dict[str, Dict[str, Any]] = {}
    for claim in claims:
        for ref in list(claim.get("cites") or []):
            if ref in out:
                continue
            item = contract.fact(ref)
            if item is None:
                continue
            meta: Dict[str, Any] = {"id": item["id"], "kind": "derived" if item.get("basis") == "derived" else "fact",
                                    "value": item.get("value"), "period": item.get("period"), "line": item.get("line")}
            if meta["kind"] == "derived":
                meta["recipe"] = derived_recipe(item)
                # the facts it was computed from: a figure can keep its value while what it came from moved
                meta["inputs"] = {i: (contract.fact(i) or {}).get("value") for i in item.get("inputs") or []}
            out[ref] = meta
        for ref in list(claim.get("cause_cites") or []):
            note = contract.note(ref)
            if note is not None and ref not in out:
                out[ref] = {"id": note["id"], "kind": "note", "text_sha": hashlib.sha256(str(note.get("text") or "").encode("utf-8")).hexdigest()[:16],
                            "origin": note.get("origin"), "note_kind": note.get("kind"), "period": note.get("period")}
    return out


def derived_recipe(item: Dict[str, Any]) -> Dict[str, Any]:
    """How a derived figure was computed, so it can be computed again in a fresh reading (and keep its D number)."""
    return {"kind": item.get("kind"), "periods": list(item.get("periods") or []), "period": item.get("period"),
            "line": item.get("line"), "inputs": list(item.get("inputs") or [])}


#: most facts of one period kept by value; past it only their hash is kept (the change is still found, just not named)
CONTEXT_FACT_CAP = 400


def _sha16(obj: Any) -> str:
    return hashlib.sha256(json.dumps(obj, sort_keys=True, ensure_ascii=False, default=str).encode("utf-8")).hexdigest()[:16]


def _round2(value: Any) -> Optional[float]:
    try:
        return None if value is None else round(float(value), 2)
    except (TypeError, ValueError):
        return None


def touched_periods(contract: ec.EvidenceContract, claims: List[Dict[str, Any]], cite_ids: Dict[str, Dict[str, Any]]) -> List[str]:
    """The periods a draft is about: the ones its claims declare and the ones its cited figures belong to."""
    found = set()
    for claim in claims:
        found |= {str(p) for p in claim.get("periods") or []}
    for meta in cite_ids.values():
        if meta.get("period"):
            found.add(str(meta["period"]))
        found |= {str(p) for p in (meta.get("recipe") or {}).get("periods") or []}
    return sorted(p for p in found if p in contract.periods)


def _period_state(contract: ec.EvidenceContract, period: str, permissions: Dict[str, Any], matrix: Dict[str, Any]) -> Dict[str, Any]:
    facts = {f["id"]: _round2(f.get("value")) for f in contract.facts if f.get("period") == period}
    notes = sorted(_sha16(n.get("text")) for n in contract.notes if n.get("origin") != "search" and n.get("period") in (period, None))
    return {"facts": facts if len(facts) <= CONTEXT_FACT_CAP else None, "facts_sha": _sha16(sorted(facts.items())), "n_facts": len(facts),
            "checks": dict(matrix.get(period) or {}), "permissions": permissions.get(period), "notes": notes}


def account_profile(session: Any, key: str) -> Dict[str, Any]:
    """Where an account's figures come from: the tab, the unit the tab is printed in (the multiplier to yuan), how the resolver chose
    the tab and any manual override. A draft can keep every figure while the account is read from another tab or another unit."""
    df = (getattr(session, "dfs", None) or {}).get(key)
    attrs = getattr(df, "attrs", None) or {}
    resolved = ((getattr(session, "resolution", None) or {}).get("resolved") or {}).get(key) or {}
    overrides = (getattr(session, "load_params", None) or {}).get("mapping_overrides") or {}
    return {"sheet": attrs.get("source_sheet_name") or resolved.get("sheet_name"), "multiplier": attrs.get("source_multiplier"),
            "mapping_method": resolved.get("resolution_method"), "override": overrides.get(key)}


def capture_context(contract: ec.EvidenceContract, claims: List[Dict[str, Any]], cite_ids: Dict[str, Dict[str, Any]],
                    profile: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """Everything about the account in the periods a draft touches, not only the figures it quotes: every other line
    of the period (a composition claim rests on the whole set, a total on its parts), what each check came out as,
    what each period permits, the remarks that apply, and which period is the latest. A paragraph can keep every
    figure it cites while the thing it says about them stops being true; this is how that is noticed."""
    from ..pptx import lineage as L
    permissions, matrix = L.permission_view(contract), contract.status_matrix()
    out = {"periods": list(contract.periods), "latest_period": contract.periods[-1] if contract.periods else None,
           "annualization_sha": _sha16(contract.annualization or {}),
           "by_period": {p: _period_state(contract, p, permissions, matrix) for p in touched_periods(contract, claims, cite_ids)}}
    if profile is not None:
        out["source"] = dict(profile)
    return out


def _fact_label(contract: ec.EvidenceContract, fact_id: str) -> str:
    item = contract.fact(fact_id)
    if item is None:
        return fact_id.split("|")[-1] or fact_id
    return " ".join(x for x in (str(item.get("line") or item.get("role") or ""), f"[{item.get('basis')}]" if item.get("basis") else "") if x)


def _profile_changes(before: Optional[Dict[str, Any]], now: Optional[Dict[str, Any]]) -> List[str]:
    if not before or not now:
        return []
    out = []
    if (before.get("sheet") or None) != (now.get("sheet") or None):
        out.append(f"the account is now read from tab '{now.get('sheet')}', not '{before.get('sheet')}'")
    if before.get("multiplier") != now.get("multiplier"):
        out.append(f"the unit its tab is printed in changed (x{before.get('multiplier')} -> x{now.get('multiplier')} to yuan)")
    if (before.get("override") or None) != (now.get("override") or None):
        out.append(f"its manual mapping override changed ({before.get('override') or 'none'} -> {now.get('override') or 'none'})")
    elif (before.get("mapping_method") or None) != (now.get("mapping_method") or None):
        out.append(f"the way its tab was chosen changed ({before.get('mapping_method')} -> {now.get('mapping_method')})")
    return out


def context_changes(contract: ec.EvidenceContract, context: Dict[str, Any], *, limit: int = 8,
                    profile_now: Optional[Dict[str, Any]] = None) -> List[str]:
    """What differs, in words, between the account as a draft was written against it and the account now. Empty when
    nothing the draft could rest on moved. `context` empty means it was never recorded: that is reported, not skipped."""
    if not context or "by_period" not in context:
        return ["what this draft rests on was not recorded when it was written, so a change in the account cannot be ruled out"]
    from ..pptx import lineage as L
    permissions, matrix = L.permission_view(contract), contract.status_matrix()
    out: List[str] = _profile_changes(context.get("source"), profile_now)
    if list(context.get("periods") or []) != list(contract.periods):
        added = [p for p in contract.periods if p not in (context.get("periods") or [])]
        gone = [p for p in context.get("periods") or [] if p not in contract.periods]
        out.append("the account's periods changed" + (f": added {', '.join(added)}" if added else "") + (f"; gone {', '.join(gone)}" if gone else "")
                   + (f"; the latest period is now {contract.periods[-1]}" if added and contract.periods else ""))
    if context.get("annualization_sha") != _sha16(contract.annualization or {}):
        out.append("the part-year / annualisation basis of the account changed")
    for period, old in (context.get("by_period") or {}).items():
        if period not in contract.periods:
            out.append(f"{period}: the period is no longer in the account")
            continue
        new = _period_state(contract, period, permissions, matrix)
        if old.get("facts_sha") != new["facts_sha"]:
            before, after = old.get("facts"), new["facts"]
            if before is not None and after is not None:
                moved = [f"{_fact_label(contract, i)} {before.get(i):,.0f} -> {after[i]:,.0f}" for i in after
                         if i in before and before[i] is not None and after[i] is not None and abs(after[i] - before[i]) > 0.5]
                added = [_fact_label(contract, i) for i in after if i not in before]
                removed = [i.split("|")[-1] for i in before if i not in after]
                bits = moved[:3] + ([f"new: {', '.join(added[:3])}"] if added else []) + ([f"gone: {', '.join(removed[:3])}"] if removed else [])
                out.append(f"{period}: the account's other figures moved ({len(moved)} changed, {len(added)} new, {len(removed)} gone): " + "; ".join(bits))
            else:
                out.append(f"{period}: the account's figures differ ({old.get('n_facts')} -> {new['n_facts']} facts; too many to name)")
        for kind in sorted(set(old.get("checks") or {}) | set(new["checks"])):
            a, b = (old.get("checks") or {}).get(kind), new["checks"].get(kind)
            if a != b:
                out.append(f"{period}: check {kind} was {a}, is now {b}")
        for key in L.PERMISSION_FIELDS:
            a, b = (old.get("permissions") or {}).get(key), (new["permissions"] or {}).get(key)
            if a != b:
                out.append(f"{period}: what the period permits changed ({key}: {a} -> {b})")
        if sorted(old.get("notes") or []) != new["notes"]:
            gone_n = len(set(old.get("notes") or []) - set(new["notes"]))
            new_n = len(set(new["notes"]) - set(old.get("notes") or []))
            out.append(f"{period}: remarks on the account changed ({new_n} new or reworded, {gone_n} gone)")
    return out[:limit] + ([f"... and {len(out) - limit} more"] if len(out) > limit else [])


def replay_derived(contract: ec.EvidenceContract, recipes: List[Dict[str, Any]]) -> List[str]:
    """Compute the figures again, in order, so the same D numbers come out. Returns the ids that could not be rebuilt."""
    from ..ai import claims as cl
    missing: List[str] = []
    for r in recipes:
        kind, periods = r.get("kind"), list(r.get("periods") or [])
        try:
            if kind == cl.KIND_MOVEMENT and len(periods) == 2:
                cl.derive_movement(contract, periods[0], periods[1])
            elif kind == cl.KIND_SHARE:
                cl.derive_share(contract, r.get("line") or "", r.get("period") or "")
            elif kind == cl.KIND_SUM:
                cl.derive_sum(contract, list(r.get("inputs") or []))
            elif kind == cl.KIND_DIFFERENCE:
                cl.derive_difference(contract, r.get("period") or "")
            elif kind == cl.KIND_ANNUALIZED:
                cl.derive_annualized(contract, r.get("period") or "")
            else:
                missing.append(str(r))
        except Exception:                                           # noqa: BLE001
            missing.append(str(r))
    return missing


def claims_digest(claims: List[Dict[str, Any]]) -> str:
    blob = json.dumps(claims, ensure_ascii=False, sort_keys=True, default=str)
    return hashlib.sha256(blob.encode("utf-8")).hexdigest()[:16]


def check_base_run(loaded: Dict[str, Any], dfs: Dict[str, pd.DataFrame], identity_now: Dict[str, Any]) -> Dict[str, Any]:
    """Does the commentary of an earlier run describe THIS workbook? The same two proofs the offline re-export asks for:
    the run's identity file (workbook hash, entity, sheet, language) and, per account, a hash of the data it was
    written from. A run with neither proves nothing, and its commentary is not ours to carry into a deck unlabelled.

    status
      verified      the identity matches, or every account's data matches its fingerprint
      edited_since  the workbook was saved after the run, and every account has a fingerprint, so it is known which
                    accounts' data changed (`accounts.changed`): the rest of the commentary still describes the data
      unverified    nothing proves it, nothing contradicts it
      mismatch      something contradicts it: another entity, sheet or language, an account that is gone, or a changed
                    workbook with no fingerprints to say which accounts it touched"""
    from ..ai.evidence import frame_fingerprint
    from ..run_identity import compare_identity
    results = loaded.get("results") or {}
    accounts = [k for k, v in results.items() if not str(k).startswith("__") and isinstance(v, dict)]
    saved, fingerprints = loaded.get("identity"), loaded.get("fingerprints") or {}
    diffs = compare_identity(saved, identity_now) if saved else []
    edits = [d for d in diffs if d.split(":")[0] in ("workbook_sha256", "financials_sha256")]
    hard = [d for d in diffs if d not in edits]
    missing = [k for k in accounts if k not in dfs]
    changed = [k for k in accounts if k in dfs and fingerprints.get(k) and fingerprints[k] != frame_fingerprint(dfs[k])]
    unprinted = [k for k in accounts if k in dfs and not fingerprints.get(k)]
    reasons = [f"identity differs -- {m}" for m in hard] + [f"account not in this workbook: {k}" for k in missing]
    fingerprinted = bool(accounts) and not unprinted
    if edits and not fingerprinted:
        reasons += [f"the workbook was saved after the run -- {m}" for m in edits] + ["and the run holds no data fingerprints to say which accounts that touched"]
    if reasons:
        status = "mismatch"
    elif changed:
        status = "edited_since"
        reasons = [f"account data differs from the run's: {k}" for k in changed]
    elif saved or fingerprinted:
        status = "verified"
    else:
        status = "unverified"
        reasons.append("the run has no identity file and " + (f"{len(unprinted)} of its {len(accounts)} accounts" if len(unprinted) < len(accounts)
                                                              else "none of its accounts") + " carry a data fingerprint: nothing ties it to this workbook")
    return {"status": status, "reasons": reasons,
            "accounts": {"in_run": len(accounts), "matched": len(accounts) - len(missing) - len(changed) - len(unprinted),
                         "changed": changed, "missing": missing, "no_fingerprint": unprinted},
            "identity_file": bool(saved), "workbook_edited_since": bool(edits)}


class SourceSession:
    def __init__(
        self, *, binding: ec.SourceBinding, dfs: Dict[str, pd.DataFrame], contracts: Dict[str, ec.EvidenceContract],
        selection: Any = None, bs_is_results: Optional[Dict[str, Any]] = None, mappings: Optional[Dict[str, Any]] = None,
        language: str = "Chi", digest: Optional[Dict[str, Any]] = None, workbook_path: str = "",
        financials_from: Optional[str] = None, financials_sheet: Optional[str] = None, entity_name: str = "",
        base_results: Optional[Dict[str, Any]] = None, base_run_folder: Optional[str] = None,
        reconciliation: Any = None, resolution: Optional[Dict[str, Any]] = None,
        base_summaries: Optional[Dict[str, str]] = None, financials_by_key: Optional[Dict[str, Dict[str, float]]] = None,
        base_run_check: Optional[Dict[str, Any]] = None,
    ) -> None:
        self.binding = binding
        self.dfs = dfs
        self.contracts = contracts
        # One source version for the whole session: a claim stamped with it must verify
        # against any account's contract, so the contracts take the session's binding.
        for contract in contracts.values():
            contract.binding = binding.to_dict()
        self.selection = selection
        self.bs_is_results = bs_is_results
        self.mappings = mappings or {}
        self.language = language
        self.digest = digest
        self.workbook_path = workbook_path
        self.financials_from = financials_from
        self.financials_sheet = financials_sheet
        self.entity_name = entity_name
        #: commentary of an earlier run (pipeline-shaped results), the base the
        #: candidates replace accounts in
        self.base_results = base_results
        self.base_run_folder = base_run_folder
        self.base_summaries = base_summaries or {}
        #: whether the base run is shown to describe THIS workbook (check_base_run); None when there is no base run
        self.base_run_check = base_run_check
        self.reconciliation = reconciliation
        self.resolution = resolution
        #: {account key: {ISO period: Financials value}}, kept so the cross-account tests can be run on request
        self.financials_by_key = financials_by_key or {}
        #: the durable side (workspace.WorkSession) when there is one: told when the sources are read again
        self.workspace: Any = None
        #: parsed text of the visible sheets, per file reading (search_notes)
        self.text_cache: Dict[str, Any] = {}
        self.candidates: Dict[str, Candidate] = {}
        self.exports: List[Dict[str, Any]] = []
        self.events: List[Dict[str, Any]] = []
        #: what load() was called with, so refresh() can read the same sources again
        self.load_params: Dict[str, Any] = {}
        #: the semantic snapshot of the sources before the last refresh (None: no earlier reading)
        self.baseline: Optional[Dict[str, Dict[str, Any]]] = None

    # -- loading -----------------------------------------------------------------

    @classmethod
    def load(
        cls, workbook: str, *, entity: str = "", sheet: Optional[str] = None, financials_from: Optional[str] = None,
        financials_sheet: Optional[str] = None, base_run: Optional[str] = None, mapping_overrides: Optional[Dict[str, str]] = None,
    ) -> "SourceSession":
        """Read the workbook (and Financials), settle which accounts are in scope,
        and build every account's contract. Deterministic; no model call."""
        from ..ai.facts import build_financials_by_key
        from ..ui.pptx_export import batch_extract_entity_data

        fin_sheet = financials_sheet if financials_from else (sheet or find_financials_sheet(workbook))
        refuse_hidden(financials_from or workbook, fin_sheet)   # raises HiddenSheetError
        overrides = {str(k): str(v) for k, v in (mapping_overrides or {}).items() if k and v}
        extracted = batch_extract_entity_data(
            temp_path=workbook, entity_name=entity or "", selected_sheet=(None if financials_from else fin_sheet),
            financials_from=financials_from, financials_sheet=(financials_sheet if financials_from else None),
            mapping_overrides=overrides or None)
        if extracted.get("status") != "ok":
            raise RuntimeError(extracted.get("error") or "the workbook could not be read")
        internal = extracted["_internal"]
        state, dfs, mappings = internal["raw_state"], internal["dfs"], internal["mappings"]
        bs_is = state.get("bs_is_results") or {}
        identity = build_identity(
            workbook=workbook, entity=entity or "", sheet=fin_sheet, financials_from=financials_from,
            financials_sheet=financials_sheet or fin_sheet, language=str(state.get("language") or ""),
            model_type="", model_name=None, accounts=list(dfs))
        binding = ec.SourceBinding.from_identity(identity, statement_source=str(bs_is.get("statement_source") or ""))
        binding.mapping = _sha16(sorted(overrides.items())) if overrides else ""
        digest = (state.get("resolution") or {}).get("workbook_digest")
        fin: Dict[str, Dict[str, float]] = {}
        for statement in ("balance_sheet", "income_statement"):
            for key, values in build_financials_by_key({statement: bs_is.get(statement)}, dfs, mappings).items():
                fin.setdefault(key, {}).update(values)
        contracts = {}
        for key, df in dfs.items():
            stype = str((df.attrs.get("integrity") or {}).get("statement_type") or "")
            contracts[key] = ec.build_contract(key, df, statement_type=stype, financials=fin.get(key),
                                               binding=binding, digest=digest)
        base_results, base_folder, base_summaries, base_check = None, None, {}, None
        if base_run:
            from ..run_identity import load_run, resolve_run_folder
            base_folder = resolve_run_folder(base_run)
            if base_folder:
                loaded = load_run(base_folder)
                base_results, base_summaries = loaded["results"], loaded.get("summaries") or {}
                # compared the way the CLI wrote it: financials_sheet is only set for a roll-up, not for the workbook's own sheet
                compare_with = build_identity(
                    workbook=workbook, entity=entity or "", sheet=fin_sheet, financials_from=financials_from, financials_sheet=financials_sheet,
                    language=str(state.get("language") or ""), model_type="", model_name=None, accounts=list(dfs))
                base_check = check_base_run(loaded, dfs, compare_with)
            else:
                base_check = {"status": "mismatch", "reasons": [f"no run folder with a results.yml for {base_run!r}"], "accounts": {}}
        session = cls(
            binding=binding, dfs=dfs, contracts=contracts, selection=internal.get("selection"), bs_is_results=bs_is,
            mappings=mappings, language=str(state.get("language") or "Chi"), digest=digest, workbook_path=workbook,
            financials_from=financials_from, financials_sheet=financials_sheet or fin_sheet,
            entity_name=entity or "", base_results=base_results, base_run_folder=base_folder,
            base_summaries=base_summaries, financials_by_key=fin, base_run_check=base_check,
            reconciliation=internal.get("reconciliation"), resolution=internal.get("resolution"))
        session.load_params = {"workbook": workbook, "entity": entity, "sheet": sheet, "financials_from": financials_from,
                               "financials_sheet": financials_sheet, "base_run": base_run, "mapping_overrides": overrides}
        return session

    def snapshot(self) -> Dict[str, Dict[str, Any]]:
        """Fact values, check statuses and periods per account: enough to say, later, what changed."""
        out: Dict[str, Dict[str, Any]] = {}
        for account, contract in self.contracts.items():
            from ..pptx import lineage as L
            out[account] = {"facts": {f["id"]: f["value"] for f in list(contract.facts) + list(contract.derived) if f.get("value") is not None},
                            "checks": {f"{c['period']}|{c['kind']}": c["status"] for c in contract.checks},
                            "periods": list(contract.periods), "permissions": L.permission_view(contract), "notes": L.note_hashes(contract, searched=False)}
        return out

    def refresh(self) -> Dict[str, Dict[str, Any]]:
        """Read the same sources again; returns the snapshot of what was there before. Drafts written
        against the old reading stay on the session but no longer count as validated (their source
        version no longer matches)."""
        if not self.load_params:
            raise RuntimeError("this session was not loaded from a workbook; there is nothing to read again")
        before = self.snapshot()
        fresh = SourceSession.load(**self.load_params)
        for name in ("binding", "dfs", "contracts", "selection", "bs_is_results", "mappings", "language", "digest",
                     "reconciliation", "resolution", "financials_by_key", "base_run_check", "load_params"):
            setattr(self, name, getattr(fresh, name))
        if fresh.base_results is not None:                  # a base run was (re)bound by load_params; otherwise the current one stays
            for name in ("base_results", "base_run_folder", "base_summaries"):
                setattr(self, name, getattr(fresh, name))
        self.text_cache = {}
        self.baseline = before
        return before

    # -- state -----------------------------------------------------------------------

    @property
    def source_version(self) -> str:
        return self.binding.source_version

    def stale_reason(self) -> Optional[str]:
        """Why this session no longer describes the workbook, or None. Compared by
        hash, so a file touched without being changed is not stale."""
        if not self.workbook_path or not self.binding.workbook_sha256:
            return None
        now = file_sha256(self.workbook_path)
        if now != self.binding.workbook_sha256:
            return f"the workbook changed on disk since this session read it ({os.path.basename(self.workbook_path)})"
        if self.financials_from and self.binding.financials_sha256 and file_sha256(self.financials_from) != self.binding.financials_sha256:
            return f"the Financials workbook changed on disk since this session read it ({os.path.basename(self.financials_from)})"
        return None

    def contract(self, account: str) -> Optional[ec.EvidenceContract]:
        return self.contracts.get(self.resolve_account(account) or "")

    def resolve_account(self, name: str) -> Optional[str]:
        """An account key from a name: exact, then case/space-insensitive, then a
        unique substring. Ambiguous or unknown returns None (the caller reports the choices)."""
        want = str(name or "").strip()
        if want in self.contracts:
            return want
        low = re.sub(r"\s+", "", want).lower()
        hits = [k for k in self.contracts if re.sub(r"\s+", "", k).lower() == low]
        if not hits and low:
            hits = [k for k in self.contracts if low in re.sub(r"\s+", "", k).lower()]
        return hits[0] if len(hits) == 1 else None

    def in_scope(self, account: str) -> bool:
        if self.selection is None:
            return True
        decision = self.selection.by_key().get(account)
        return True if decision is None else bool(decision.admitted)

    def log(self, **event: Any) -> None:
        self.events.append(event)
