"""What an account's figures are, where each came from, and what may be said from them.

`AccountEvidence` held the pool a sentence is searched against: a list of values
with a sheet and a row. That answers "does this number appear somewhere". It does
not answer the questions a reviewer asks first:

  * is this the tab's figure, Financials' figure, or the one the deck uses after
    the owner's rule "the deck follows Financials" replaced the tab's total?
  * for which periods was that checked, against what, and with what result?
  * if the total was replaced, does the tab's detail still explain it?

This module is the answer, built once per account from the frame the model reads
and carried on `AccountEvidence.contract`. It adds nothing parallel: the totals
come from the same analysis frame, the Financials figures from the same
`tab_to_financials` edges, the replaced totals from the record
`workbook/authority.py` already leaves on the frame (`financials_authority`).

Three things are kept apart, because folding them is how a replaced total came to
make the detail under it look verified:

  FACTS    one record per figure and BASIS. `raw_schedule` is the tab's own number,
           `financials` the statement's, `authority_adjusted` what the frame holds
           now. They coexist; none overwrites another.
  CHECKS   one record per account x period x kind: what was compared, the
           tolerance, the result. Results are pass / mismatch / ambiguous /
           not_comparable / missing -- never a boolean. "Could not tell" stays
           "could not tell".
  PERMISSIONS  what the commentary may do with the account in that period,
           derived from the checks. A total that was replaced may be described;
           the tab's detail may not be used to explain it; a period the tab never
           updated may not be called "nil" or "first arose".

A hidden sheet is never evidence. An account whose only tab is hidden gets a
contract with no citeable facts and an issue saying why.
"""
from __future__ import annotations

import hashlib
import numbers
import re
from dataclasses import dataclass, field
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd

CONTRACT_VERSION = 1
PARSER_VERSION = "evidence-contract-1"

BASIS_RAW = "raw_schedule"
BASIS_FINANCIALS = "financials"
BASIS_ADJUSTED = "authority_adjusted"
BASIS_DERIVED = "derived"

ROLE_TOTAL = "total"
ROLE_COMPONENT = "component"

CHECK_TOTAL_VS_FINANCIALS = "total_vs_financials"
CHECK_DETAIL_VS_TOTAL = "detail_vs_total"

PASS = "pass"
MISMATCH = "mismatch"
AMBIGUOUS = "ambiguous"
NOT_COMPARABLE = "not_comparable"
MISSING = "missing"
NOT_CHECKABLE = "not_checkable"

#: Reconciliation's own materiality (workbook/reconcile.py): 0.5% of the larger
#: side, never tighter than one yuan.
TIE_RELATIVE = 0.005
TIE_ABSOLUTE = 1.0


# A note longer than this is cut. 400 until 2026-11-09, which cut real notes on a client file mid-sentence (the evidence view
# shows every note, so this is the only limit on what the writer reads of one).
NOTE_CHARS = 1500


def find_nesting(parts: List[Dict[str, Any]], tolerance: float = TIE_ABSOLUTE) -> Dict[str, Any]:
    """Lines that are the exact sum of the run of lines right under them (or right above them, a subtotal), in sheet order.
    {"parents": {parent id: [direct child ids]}, "children": {every id under some parent}}. Needs a row for every line; two
    lines at least per run. Several levels (category -> counterparty -> contract): a line found to be a parent counts as one
    unit, its own sub-lines skipped, when a run is summed for the level above it; passes repeat until nothing new is found.
    Single-level until 2026-11-09, which left a three-level layout 'ambiguous' and its composition blocked."""
    rows = [f for f in parts if f.get("row") is not None]
    if len(rows) != len(parts) or len(rows) < 3:
        return {"parents": {}, "children": set()}
    rows.sort(key=lambda f: f["row"])
    parents: Dict[str, List[str]] = {}
    parent_of: Dict[str, str] = {}

    def top_of(fid: str) -> str:
        while fid in parent_of:
            fid = parent_of[fid]
        return fid

    found = True
    while found:
        found = False
        for i, line in enumerate(rows):
            if line["id"] in parents or line["id"] in parent_of or abs(line["value"]) <= tolerance:
                continue
            for run in (range(i + 1, len(rows)), range(i - 1, -1, -1)):
                total, picked = 0.0, []
                for j in run:
                    rid = rows[j]["id"]
                    if rid in parent_of:
                        if top_of(rid) in picked:
                            continue        # a sub-line of a unit already counted
                        break               # belongs to a block outside this run
                    total += rows[j]["value"]
                    picked.append(rid)
                    if len(picked) >= 2 and abs(total - line["value"]) <= tolerance:
                        parents[line["id"]] = list(picked)
                        for child in picked:
                            parent_of[child] = line["id"]
                        found = True
                        break
                if line["id"] in parents:
                    break
    return {"parents": parents, "children": set(parent_of)}


def tie_tolerance(left: float, right: float) -> float:
    return max(TIE_ABSOLUTE, TIE_RELATIVE * max(abs(left), abs(right)))


# ---------------------------------------------------------------------------
# Records
# ---------------------------------------------------------------------------

@dataclass
class SourceBinding:
    """Which workbook, which Financials, which parser: the version every claim,
    tool result and deck element is tied to."""
    entity: str = ""
    workbook: str = ""
    workbook_sha256: Optional[str] = None
    financials_sha256: Optional[str] = None
    financials_sheet: str = ""
    sheet: str = ""
    statement_source: str = ""          # "financials" | "synthesized" | ""
    language: str = ""
    unit: str = "yuan"
    currency: str = "CNY"
    parser_version: str = PARSER_VERSION
    #: a digest of the manual mapping overrides the accounts were read with ("" when none, which leaves every earlier version as it was)
    mapping: str = ""

    @property
    def source_version(self) -> str:
        """Changes when any input the figures were read from changes (a manual mapping override is one)."""
        parts = [self.workbook_sha256, self.financials_sha256, self.financials_sheet, self.entity, self.sheet,
                 self.statement_source, self.parser_version]
        if self.mapping:
            parts.append(f"mapping={self.mapping}")
        blob = "|".join(str(x or "") for x in parts)
        return hashlib.sha256(blob.encode("utf-8")).hexdigest()[:16]

    def to_dict(self) -> Dict[str, Any]:
        out = dict(self.__dict__)
        out["source_version"] = self.source_version
        return out

    @classmethod
    def from_dict(cls, data: Optional[Dict[str, Any]]) -> "SourceBinding":
        data = dict(data or {})
        data.pop("source_version", None)
        known = {k: v for k, v in data.items() if k in cls.__dataclass_fields__}
        return cls(**known)

    @classmethod
    def from_identity(cls, identity: Dict[str, Any], statement_source: str = "") -> "SourceBinding":
        """From a run_identity.build_identity() record."""
        return cls(
            entity=str(identity.get("entity") or ""), workbook=str(identity.get("workbook") or ""),
            workbook_sha256=identity.get("workbook_sha256"), financials_sha256=identity.get("financials_sha256"),
            financials_sheet=str(identity.get("financials_sheet") or ""), sheet=str(identity.get("sheet") or ""),
            statement_source=statement_source, language=str(identity.get("language") or ""))


@dataclass
class CheckResult:
    id: str
    account: str
    period: str
    kind: str
    status: str
    inputs: List[str] = field(default_factory=list)       # fact ids compared
    left: Optional[float] = None
    right: Optional[float] = None
    diff: Optional[float] = None
    tolerance: Optional[float] = None
    overridden: bool = False        # the deck uses the Financials figure in place of the tab's
    note: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return dict(self.__dict__)


@dataclass
class Permissions:
    """What commentary may do with one account in one period. Separate from the
    check results they come from: a figure that does not tie can still be
    described as not tying."""
    describe_total: bool = False
    describe_difference: bool = False
    explain_with_detail: bool = False
    #: tied (detail adds up) | ambiguous (flat sum misses; nesting unknown) |
    #: untied (hierarchy known / total replaced, and it does not add up) | none
    detail_basis: str = "none"
    state_nil_or_first_arose: bool = False
    compute_movement: bool = False
    #: why a movement into this period is not computed (empty when it is, or when there is no earlier period)
    movement_block: str = ""
    reasons: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return dict(self.__dict__)


@dataclass
class EvidenceContract:
    account: str
    statement_type: str = ""
    binding: Dict[str, Any] = field(default_factory=dict)
    periods: List[str] = field(default_factory=list)
    period_meta: Dict[str, Dict[str, Any]] = field(default_factory=dict)
    facts: List[Dict[str, Any]] = field(default_factory=list)
    checks: List[Dict[str, Any]] = field(default_factory=list)
    permissions: Dict[str, Dict[str, Any]] = field(default_factory=dict)
    issues: List[Dict[str, str]] = field(default_factory=list)
    annualization: Dict[str, Any] = field(default_factory=dict)
    #: text evidence a cause may rest on: the tab's own notes and remarks, and
    #: the adjustment entries behind a replaced total. One record each, with a
    #: short ref (N1...), so a sentence about WHY can cite what it rests on.
    notes: List[Dict[str, Any]] = field(default_factory=list)
    #: figures Python computed from facts (claims.derive_*): movements, shares,
    #: sums, annualised values. Each names its formula and the facts it used.
    derived: List[Dict[str, Any]] = field(default_factory=list)
    version: int = CONTRACT_VERSION

    # -- lookups ----------------------------------------------------------------

    def fact(self, fact_id: str) -> Optional[Dict[str, Any]]:
        """A fact or a derived fact by semantic id or short ref."""
        ref = str(fact_id)
        for item in list(self.facts) + list(self.derived):
            if item["id"] == ref or item.get("ref") == ref:
                return item
        return None

    def note(self, note_id: str) -> Optional[Dict[str, Any]]:
        ref = str(note_id)
        for item in self.notes:
            if item["id"] == ref or item.get("ref") == ref:
                return item
        return None

    def check(self, period: str, kind: str) -> Optional[Dict[str, Any]]:
        for item in self.checks:
            if item["period"] == str(period) and item["kind"] == kind:
                return item
        return None

    def total(self, period: str, basis: str = BASIS_ADJUSTED) -> Optional[float]:
        for item in self.facts:
            if item["period"] == str(period) and item["role"] == ROLE_TOTAL and item["basis"] == basis:
                return item["value"]
        return None

    @property
    def source_version(self) -> str:
        return SourceBinding.from_dict(self.binding).source_version

    def status_matrix(self) -> Dict[str, Dict[str, str]]:
        """{period: {check kind: status}} -- the account x period x basis view."""
        out: Dict[str, Dict[str, str]] = {}
        for item in self.checks:
            out.setdefault(item["period"], {})[item["kind"]] = item["status"]
        return out

    # -- persistence ---------------------------------------------------------------

    def to_dict(self) -> Dict[str, Any]:
        return {
            "version": self.version, "account": self.account, "statement_type": self.statement_type,
            "binding": dict(self.binding), "periods": list(self.periods), "period_meta": dict(self.period_meta),
            "facts": list(self.facts), "checks": list(self.checks), "permissions": dict(self.permissions),
            "issues": list(self.issues), "annualization": dict(self.annualization),
            "notes": list(self.notes), "derived": list(self.derived),
        }

    @classmethod
    def from_dict(cls, data: Optional[Dict[str, Any]]) -> Optional["EvidenceContract"]:
        if not data:
            return None
        return cls(
            account=str(data.get("account") or ""), statement_type=str(data.get("statement_type") or ""),
            binding=dict(data.get("binding") or {}), periods=list(data.get("periods") or []),
            period_meta=dict(data.get("period_meta") or {}), facts=list(data.get("facts") or []),
            checks=list(data.get("checks") or []), permissions=dict(data.get("permissions") or {}),
            issues=list(data.get("issues") or []), annualization=dict(data.get("annualization") or {}),
            notes=list(data.get("notes") or []), derived=list(data.get("derived") or []),
            version=int(data.get("version") or CONTRACT_VERSION))


# ---------------------------------------------------------------------------
# Building
# ---------------------------------------------------------------------------

def _number(value: Any) -> Optional[float]:
    # numbers.Real, not (int, float): a pandas int64 cell is neither, and a frame
    # with an integer column would silently lose every figure in it.
    if isinstance(value, bool) or not isinstance(value, numbers.Real):
        return None
    number = float(value)
    return None if number != number else number


def _fact_id(account: str, period: str, basis: str, line: str, taken: set) -> str:
    base = f"{account}|{period}|{basis}|{line or '#total'}"
    candidate, n = base, 1
    while candidate in taken:
        n += 1
        candidate = f"{base}~{n}"
    taken.add(candidate)
    return candidate


def _financials_from_edges(account: str, graph_facts: Optional[Dict[str, Any]]) -> Dict[str, float]:
    for edge in (graph_facts or {}).get("links") or []:
        if not isinstance(edge, dict) or edge.get("kind") != "tab_to_financials":
            continue
        if str(edge.get("source") or "").split("::", 1)[0] != account:
            continue
        values = (edge.get("evidence") or {}).get("financials_values") or {}
        return {str(p): float(v) for p, v in values.items() if _number(v) is not None}
    return {}


def _sheet_hidden(df: Any, digest: Optional[Dict[str, Any]]) -> Tuple[str, bool]:
    attrs = getattr(df, "attrs", None) or {}
    sheet = str(attrs.get("source_sheet_name") or (attrs.get("integrity") or {}).get("sheet_name") or "")
    entry = ((digest or {}).get("sheets") or {}).get(sheet) or {}   # digest-ok: reads the hidden flag itself
    return sheet, bool(entry.get("hidden"))


def build_contract(
    mapping_key: str,
    df: Any,
    *,
    statement_type: str = "",
    financials: Optional[Dict[str, float]] = None,
    graph_facts: Optional[Dict[str, Any]] = None,
    binding: Optional[SourceBinding] = None,
    digest: Optional[Dict[str, Any]] = None,
    max_components: int = 400,
) -> EvidenceContract:
    """The contract for one account. Pure and deterministic: same frame, same
    Financials, same contract."""
    from .facts import _full_periods, _period_lengths, _total_row_index, _value_columns
    from ..workbook.authority import AUTHORITY_KEY

    account = str(mapping_key)
    contract = EvidenceContract(account=account, statement_type=str(statement_type or ""),
                                binding=(binding or SourceBinding()).to_dict())
    attrs = getattr(df, "attrs", None) or {}
    sheet, hidden = _sheet_hidden(df, digest)
    if hidden:
        contract.issues.append({"code": "SOURCE_SHEET_HIDDEN", "severity": "blocked",
                                "message": f"the tab '{sheet}' is a hidden sheet; its figures are not offered as evidence"})
        return contract

    analysis = attrs.get("prompt_analysis_df")
    frame = analysis if isinstance(analysis, pd.DataFrame) and not analysis.empty else df
    if not isinstance(frame, pd.DataFrame) or frame.empty:
        contract.issues.append({"code": "NO_FRAME", "severity": "blocked", "message": "the account has no data frame"})
        return contract
    periods = _value_columns(frame)
    contract.periods = list(periods)
    period_lengths = _period_lengths(periods) if periods else {}
    full = set(_full_periods(periods)) if periods else set()
    for p in periods:
        contract.period_meta[p] = {"months": period_lengths.get(p), "full_period": p in full}
    ann = attrs.get("annualization") or {}
    contract.annualization = {
        "annualized": bool(ann.get("annualized")), "months": ann.get("months"), "factor": ann.get("factor"),
        "raw_effective_date": ann.get("raw_effective_date"),
        "formula": "value / months_elapsed * 12" if ann.get("annualized") else "",
    }

    row_types = attrs.get("row_types_by_description") or (frame.attrs or {}).get("row_types_by_description") or {}
    desc_col = frame.columns[0]
    total_idx = _total_row_index(frame, row_types, periods)
    record = attrs.get(AUTHORITY_KEY) or (frame.attrs or {}).get(AUTHORITY_KEY) or {}
    source_kind = (binding.statement_source if binding else "") or ""
    fin = {str(p): float(v) for p, v in (financials or {}).items() if _number(v) is not None}
    if not fin:
        fin = _financials_from_edges(account, graph_facts)
    if source_kind == "synthesized":
        fin = {}   # the "Financials" figures of a synthesized statement are these tabs' own totals
    taken: set = set()

    fin_sheet = (binding.financials_sheet if binding else "") or None

    def add(period: str, basis: str, role: str, line: str, value: float, **extra: Any) -> str:
        fid = _fact_id(account, period, basis, line, taken)
        contract.facts.append({
            "id": fid, "account": account, "period": period, "basis": basis, "role": role, "line": line,
            "value": float(value), "unit": "yuan",
            # a Financials-basis figure comes from the Financials sheet, not the account's tab
            "sheet": (fin_sheet if basis == BASIS_FINANCIALS else sheet) or None, "visible": True,
            "parser_version": PARSER_VERSION, **extra})
        return fid

    # -- totals, each basis its own fact ------------------------------------------
    total_ids: Dict[str, Dict[str, str]] = {}
    for p in periods:
        adjusted = _number(frame.at[total_idx, p]) if total_idx is not None and p in frame.columns else None
        ids: Dict[str, str] = {}
        in_record = p in record
        raw = _number((record.get(p) or {}).get("tab")) if in_record else adjusted
        row_idx = _number(frame.at[total_idx, "__source_row_idx"]) if total_idx is not None and "__source_row_idx" in frame.columns else None
        loc = {"row": int(row_idx)} if row_idx is not None else {}
        if raw is not None:
            ids[BASIS_RAW] = add(p, BASIS_RAW, ROLE_TOTAL, "", raw, **loc)
        if p in fin:
            ids[BASIS_FINANCIALS] = add(p, BASIS_FINANCIALS, ROLE_TOTAL, "", fin[p])
        if adjusted is not None:
            ids[BASIS_ADJUSTED] = add(p, BASIS_ADJUSTED, ROLE_TOTAL, "", adjusted, overridden=in_record, **loc)
        total_ids[p] = ids

    # -- components: the tab's own detail rows (raw_schedule, whatever happened to the total)
    components: List[Tuple[str, Any]] = []
    for idx in frame.index:
        label = str(frame.at[idx, desc_col]).strip()
        if str(row_types.get(label, "")).strip().lower() in ("detail", "breakdown") and label:
            components.append((label, idx))
    for label, idx in components[:max_components]:
        row_idx = _number(frame.at[idx, "__source_row_idx"]) if "__source_row_idx" in frame.columns else None
        for p in periods:
            value = _number(frame.at[idx, p]) if p in frame.columns else None
            if value is None:
                continue
            add(p, BASIS_RAW, ROLE_COMPONENT, label, value, **({"row": int(row_idx)} if row_idx is not None else {}))
    if len(components) > max_components:
        contract.issues.append({"code": "COMPONENTS_TRUNCATED", "severity": "info",
                                "message": f"{len(components)} detail rows; the first {max_components} are in the contract"})

    # -- checks ---------------------------------------------------------------------
    absolute = str(statement_type).strip().upper() == "IS"
    for p in periods:
        ids = total_ids.get(p) or {}
        overridden = p in record
        # 1. the tab's own total against Financials -- the RAW total, so a replaced
        #    total cannot make the comparison trivially true
        cid = f"chk|{account}|{p}|{CHECK_TOTAL_VS_FINANCIALS}"
        raw_value, fin_value = contract.total(p, BASIS_RAW), contract.total(p, BASIS_FINANCIALS)
        if source_kind == "synthesized":
            contract.checks.append(CheckResult(cid, account, p, CHECK_TOTAL_VS_FINANCIALS, NOT_COMPARABLE,
                                               note="statements are synthesized from the tabs; no independent Financials").to_dict())
        elif raw_value is None or fin_value is None:
            contract.checks.append(CheckResult(
                cid, account, p, CHECK_TOTAL_VS_FINANCIALS, MISSING,
                note=("no Financials figure for this period" if fin_value is None else "the tab has no total for this period")).to_dict())
        else:
            left, right = (abs(raw_value), abs(fin_value)) if absolute else (raw_value, fin_value)
            tol = tie_tolerance(left, right)
            diff = left - right
            status = PASS if abs(diff) <= tol else MISMATCH
            contract.checks.append(CheckResult(
                cid, account, p, CHECK_TOTAL_VS_FINANCIALS, status,
                inputs=[i for i in (ids.get(BASIS_RAW), ids.get(BASIS_FINANCIALS)) if i],
                left=raw_value, right=fin_value, diff=diff, tolerance=tol, overridden=overridden and status == MISMATCH,
                note=("the deck uses the Financials figure; the tab was not updated for this period"
                      if overridden and status == MISMATCH else "")).to_dict())
        # 2. does the tab's detail explain the total the deck uses?
        did = f"chk|{account}|{p}|{CHECK_DETAIL_VS_TOTAL}"
        parts = [f for f in contract.facts if f["period"] == p and f["role"] == ROLE_COMPONENT]
        adjusted_value = contract.total(p, BASIS_ADJUSTED)
        if not parts:
            contract.checks.append(CheckResult(did, account, p, CHECK_DETAIL_VS_TOTAL, NOT_CHECKABLE,
                                               note="the tab lists no detail rows").to_dict())
        elif adjusted_value is None:
            contract.checks.append(CheckResult(did, account, p, CHECK_DETAIL_VS_TOTAL, MISSING,
                                               note="no total to compare the detail with").to_dict())
        else:
            flat = sum(f["value"] for f in parts)
            tol = tie_tolerance(flat, adjusted_value)
            diff = flat - adjusted_value
            ids_used = [f["id"] for f in parts] + [i for i in (ids.get(BASIS_ADJUSTED),) if i]
            if overridden and abs(diff) > tol:
                status, note = MISMATCH, "total set to Financials; the tab's detail was not updated and does not add up to it"
            elif abs(diff) <= tol:
                status, note = PASS, ""
            else:
                # A flat sum that misses the total is either a real gap or a nested
                # layout (a line listed beside its own sub-lines). Nothing here can
                # tell which, and saying "mismatch" would be a guess -- unless the
                # nesting shows itself: a line that equals, to the yuan, the sum of
                # the lines right under (or right above) it, and a top level that then
                # ties. Found 2026-11-08: every account laid out that way had its
                # composition blocked.
                nested = find_nesting(parts)
                top = sum(f["value"] for f in parts if f["id"] not in nested["children"])
                if nested["parents"] and abs(top - adjusted_value) <= tol:
                    for parent_id, child_ids in nested["parents"].items():
                        for f in parts:
                            if f["id"] in child_ids:
                                f["nested_in"] = parent_id      # replaced by the parent's ref once refs exist
                    status, note = PASS, (f"nested: {len(nested['parents'])} line(s) listed with their own sub-lines; the top level "
                                          f"({top:,.2f}) ties to the total")
                else:
                    status, note = AMBIGUOUS, "detail rows do not sum to the total as a flat list; nesting is not known"
            contract.checks.append(CheckResult(did, account, p, CHECK_DETAIL_VS_TOTAL, status, inputs=ids_used,
                                               left=flat, right=adjusted_value, diff=diff, tolerance=tol,
                                               overridden=overridden, note=note).to_dict())

    # -- permissions, derived from the checks ----------------------------------------
    previous: Optional[str] = None
    for p in periods:
        tf = contract.check(p, CHECK_TOTAL_VS_FINANCIALS) or {}
        dt = contract.check(p, CHECK_DETAIL_VS_TOTAL) or {}
        overridden = p in record
        perm = Permissions()
        have_total = contract.total(p, BASIS_ADJUSTED) is not None
        perm.describe_total = have_total
        if have_total and tf.get("status") == MISMATCH and not overridden:
            # Tab and Financials disagree and nothing said which to follow: the
            # difference can be described, neither figure asserted.
            perm.describe_total = False
            perm.describe_difference = True
            perm.reasons.append("tab and Financials disagree and no authority rule applied")
        if tf.get("status") == MISMATCH and overridden:
            perm.describe_difference = True
        perm.explain_with_detail = dt.get("status") == PASS
        perm.detail_basis = {PASS: "tied", AMBIGUOUS: "ambiguous", MISMATCH: "untied"}.get(dt.get("status"), "none")
        if dt.get("status") in (MISMATCH, AMBIGUOUS, NOT_CHECKABLE, MISSING):
            perm.reasons.append(f"detail check {dt.get('status')}: detail may not be used to explain the total")
        perm.state_nil_or_first_arose = have_total and not overridden and tf.get("status") != MISMATCH
        if overridden:
            perm.reasons.append("total follows Financials; do not call the account nil or first arisen from the tab")
        # A change from the previous period: both totals describable, and for a
        # flow (IS) both periods the same length -- a one-month stub against a
        # year is not a movement. A balance is a point in time and needs neither.
        if previous is not None and perm.describe_total and contract.permissions[previous]["describe_total"]:
            lengths_equal = abs((period_lengths.get(p) or 0) - (period_lengths.get(previous) or 0)) < 0.5
            perm.compute_movement = (not absolute) or lengths_equal
            if not perm.compute_movement:
                perm.movement_block = "income-statement periods differ in length"
                perm.reasons.append("income-statement periods differ in length; no movement is computed")
        contract.permissions[p] = perm.to_dict()
        previous = p

    # -- notes: what a cause may rest on --------------------------------------------------
    seen_notes: set = set()

    def add_note(kind: str, text: str, period: str = "") -> None:
        text = re.sub(r"\s+", " ", str(text or "")).strip()
        if len(re.findall(r"[一-鿿A-Za-z]", text)) < 8 or text in seen_notes:
            return          # a bare number or a stray label is not an explanation
        seen_notes.add(text)
        contract.notes.append({"id": f"{account}|note|{len(contract.notes) + 1}", "kind": kind,
                               "text": text[:NOTE_CHARS], "period": period or None, "sheet": sheet or None})

    for text in attrs.get("supporting_notes") or []:
        add_note("supporting_note", text if isinstance(text, str) else str(text))
    for remark in attrs.get("table_linked_remarks") or []:
        add_note("remark", remark.get("summary") if isinstance(remark, dict) else str(remark))
    for p, pair in record.items():
        for line in (pair or {}).get("adjustments") or []:
            add_note("adjustment", line, p)
    for line in ((attrs.get("adjustment_bridge") or {}).get("lines")) or []:
        add_note("adjustment", line)

    # -- short references the prompt and the model use ---------------------------------
    for i, item in enumerate(sorted(contract.facts, key=lambda f: f["id"]), 1):
        item["ref"] = f"F{i}"
    ref_of = {f["id"]: f["ref"] for f in contract.facts}
    for item in contract.facts:
        if item.get("nested_in") in ref_of:
            item["nested_in"] = ref_of[item["nested_in"]]
    for i, item in enumerate(contract.notes, 1):
        item["ref"] = f"N{i}"
    return contract


def attach_contract(evidence: Any, contract: Optional[EvidenceContract]) -> None:
    """File a contract on an AccountEvidence (kept as plain data)."""
    evidence.contract = contract.to_dict() if contract is not None else None


def trace_fact(contract: EvidenceContract, fact_id: str) -> Dict[str, Any]:
    """Source position -> raw fact -> Financials comparison/adjustment -> the checks
    that touch it. Claims and deck elements extend this chain (claims.py, lineage)."""
    fact = contract.fact(fact_id)
    if fact is None:
        return {"found": False, "fact_id": fact_id}
    checks = [c for c in contract.checks if fact["id"] in (c.get("inputs") or [])]
    return {
        "found": True, "fact": fact,
        "source": {"sheet": fact.get("sheet"), "row": fact.get("row"), "period": fact["period"],
                   "visible": fact.get("visible", True)},
        "binding": {"source_version": contract.source_version, **{k: contract.binding.get(k) for k in ("entity", "statement_source")}},
        "checks": checks,
        "permissions": contract.permissions.get(fact["period"]),
    }
