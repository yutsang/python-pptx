"""A changed source and a returned deck: what moved, what it touches, what may be edited.

Three questions, answered separately because a wrong answer to one must not leak into another:

  1. WHAT CHANGED in the sources?   diff_sources, by semantic fact id. Needs a baseline: the
                                    internal lineage sidecar of an earlier export (old values),
                                    or, from the deck alone, value hashes (changed or not, not
                                    from what to what). With neither, "what changed since last
                                    time" is not answerable and is not answered.
  2. WHAT DID THE USER DO to the deck?   analyze_returned: each exported paragraph is found
                                    again (tag, text, similarity) and compared with its
                                    fingerprint. Copies, reordered slides, edits, deletions,
                                    and a stripped manifest are told apart.
  3. WHAT MAY BE EDITED?            a paragraph the user left alone whose cited facts changed is
                                    updated in place; one the user edited and the sources also
                                    changed is a conflict with a proposal and is not touched;
                                    one the user edited and the sources did not change is kept
                                    and its numbers re-checked; one that cannot be found is
                                    handed to a person (inspect and relink), never guessed at.
"""
from __future__ import annotations

import os
import re
from typing import Any, Dict, List, Optional, Tuple

from pptx import Presentation

from ..ai.claims import text_amounts
from ..pptx import lineage as L

TOLERANCE = 0.5      # yuan: a fact whose value moved by less than this has not changed


def _values(contract: Any) -> Dict[str, float]:
    return {f["id"]: f["value"] for f in list(contract.facts) + list(contract.derived) if f.get("value") is not None}


def _checks(contract: Any) -> Dict[str, str]:
    return {f"{c['period']}|{c['kind']}": c["status"] for c in contract.checks}


# ---------------------------------------------------------------------------
# 1. what changed in the sources
# ---------------------------------------------------------------------------

def diff_sources(baseline: Dict[str, Dict[str, Any]], contracts: Dict[str, Any]) -> Dict[str, Any]:
    """Old values (from an internal baseline) against the current contracts, by semantic fact id."""
    accounts: Dict[str, Any] = {}
    for account in sorted(set(baseline) | set(contracts)):
        old, contract = baseline.get(account), contracts.get(account)
        if old is None:
            accounts[account] = {"status": "added"}
            continue
        if contract is None:
            accounts[account] = {"status": "removed"}
            continue
        new_vals, old_vals = _values(contract), old.get("facts") or {}
        changed = [{"id": i, "old": old_vals[i], "new": new_vals[i], "delta": new_vals[i] - old_vals[i]}
                   for i in old_vals if i in new_vals and abs(new_vals[i] - old_vals[i]) > TOLERANCE]
        added = sorted(i for i in new_vals if i not in old_vals and ec_is_fact(i))
        removed = sorted(i for i in old_vals if i not in new_vals)
        old_checks, new_checks = old.get("checks") or {}, _checks(contract)
        check_changes = [{"key": k, "old": old_checks.get(k), "new": new_checks.get(k)}
                         for k in sorted(set(old_checks) | set(new_checks)) if old_checks.get(k) != new_checks.get(k)]
        periods_added = [p for p in contract.periods if p not in (old.get("periods") or [])]
        periods_removed = [p for p in (old.get("periods") or []) if p not in contract.periods]
        permission_changes, notes_changed = [], {}
        if "permissions" in old:                     # a baseline from before permissions were recorded has none to compare
            now_perm = L.permission_view(contract)
            permission_changes = [{"period": p, "field": f, "old": (old["permissions"].get(p) or {}).get(f), "new": (now_perm.get(p) or {}).get(f)}
                                  for p in sorted(set(now_perm) | set(old["permissions"])) for f in L.PERMISSION_FIELDS
                                  if (old["permissions"].get(p) or {}).get(f) != (now_perm.get(p) or {}).get(f)]
        if "notes" in old:
            before_n, now_n = set(old["notes"]), set(L.note_hashes(contract, searched=False))
            if before_n != now_n:
                notes_changed = {"removed_or_edited": len(before_n - now_n), "added": len(now_n - before_n)}
        if changed or added or removed or check_changes or periods_added or periods_removed or permission_changes or notes_changed:
            accounts[account] = {"status": "changed", "changed": changed, "added": added[:40], "removed": removed[:40],
                                 "check_changes": check_changes, "periods_added": periods_added, "periods_removed": periods_removed,
                                 "permission_changes": permission_changes, "notes_changed": notes_changed}
    return {"mode": "values", "accounts": accounts,
            "changed_accounts": sorted(accounts), "affected_fact_ids": sorted(
                {c["id"] for a in accounts.values() for c in a.get("changed", [])} | {i for a in accounts.values() for i in a.get("removed", [])})}


def ec_is_fact(fact_id: str) -> bool:
    return "|derived|" not in fact_id


def diff_by_hash(public: Dict[str, Any], contracts: Dict[str, Any]) -> Dict[str, Any]:
    """From the deck's embedded manifest alone: which cited facts changed or vanished, not by how much."""
    changed: Dict[str, List[str]] = {}
    for element in public.get("elements", []):
        contract = contracts.get(element["account"])
        values = _values(contract) if contract is not None else {}
        for fact_id, old_hash in (element.get("cite_hashes") or {}).items():
            if fact_id not in values or L.value_hash(values[fact_id]) != old_hash:
                changed.setdefault(element["account"], []).append(fact_id)
    for element in public.get("elements", []):
        contract = contracts.get(element["account"])
        if contract is not None and element.get("deps") and L.dependency_changes(contract, element["deps"], list(element.get("cites") or [])):
            changed.setdefault(element["account"], [])
    digest_changed = []
    for account, old in (public.get("account_digest") or {}).items():
        contract = contracts.get(account)
        if contract is None:
            digest_changed.append(account)
            continue
        import hashlib
        import json
        totals = sorted((f["period"], f["basis"], round(f["value"], 2)) for f in contract.facts if f["role"] == "total")
        if hashlib.sha1(json.dumps(totals).encode()).hexdigest()[:12] != old:
            digest_changed.append(account)
    return {"mode": "hashes", "changed_cited": {a: sorted(set(v)) for a, v in changed.items()}, "changed_accounts": sorted(set(changed) | set(digest_changed)),
            "note": "values are not known, only that they differ; keep the lineage sidecar to see from what to what"}


# ---------------------------------------------------------------------------
# 2./3. a returned deck
# ---------------------------------------------------------------------------

def _source_changed(element: Dict[str, Any], contracts: Dict[str, Any], digest_changed: set) -> Tuple[bool, str]:
    contract = contracts.get(element["account"])
    if contract is None:
        return True, "the account is no longer in the sources"
    values = _values(contract)
    cites = element.get("cite_hashes") or {}
    if cites:
        gone = [i for i in cites if i not in values]
        moved = [i for i, h in cites.items() if i in values and L.value_hash(values[i]) != h]
        if gone or moved:
            return True, f"{len(moved)} cited fact(s) changed, {len(gone)} no longer exist"
        # the figures are the same; what lets them be stated, or the note that explained them, may not be
        why = L.dependency_changes(contract, element["deps"], list(cites)) if element.get("deps") else []
        if why:
            return True, "the cited figures are unchanged but " + "; ".join(why)
        return False, ""
    if element["account"] in digest_changed:
        return True, "no citations were recorded for this paragraph; the account's totals changed"
    return False, ""


def verify_user_text(text: str, contract: Any) -> Dict[str, Any]:
    """What can honestly be said about text a person wrote: its amounts, found or not in the
    account's current figures. Found by value, not cited, so never reported as 'verified'."""
    values = [abs(f["value"]) for f in list(contract.facts) + list(contract.derived) if f.get("value") is not None] if contract else []
    found, missing = [], []
    for span in text_amounts(text):
        (found if any(abs(abs(span["value"]) - v) <= span["tolerance"] for v in values) else missing).append(span["written"])
    return {"state": "not claim-verified (written by a person)", "amounts_found": found, "amounts_not_found": missing}


def analyze_returned(deck_path: str, session: Any, *, public: Optional[Dict[str, Any]] = None,
                     internal: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    public = public or L.read_manifest(deck_path)
    internal = internal if internal is not None else L.read_internal(deck_path)
    report: Dict[str, Any] = {"deck": os.path.basename(deck_path), "current_source_version": session.source_version}
    if public is None:
        report.update(mode="none", history="unknown",
                      note="no manifest: the deck cannot be tied to an export. What follows checks numbers against the sources as they are NOW; "
                           "it says nothing about what changed or who changed it.")
        labels = _labels_from_session(session)
        elements = []
        for para in L.read_paragraphs(deck_path, labels.keys()):
            account = labels.get(para.label or "")
            if account is None:
                continue
            elements.append({"slide": para.slide, "account": account, "label": para.label,
                             "check": verify_user_text(para.body, session.contracts.get(account)), "action": "inspect_and_relink"})
        report["paragraphs"] = elements
        report["actions"] = [{"action": "inspect_and_relink", "reason": "no manifest to relink against"}] if elements else []
        return report

    mode = "full" if internal is not None else "fingerprint_only"
    report.update(mode=mode, history="known" if internal is not None else "partial", manifest_source_version=public.get("source_version"),
                  source_moved=public.get("source_version") != session.source_version)
    labels = [e.get("label") or "" for e in public["elements"]]
    paragraphs = L.read_paragraphs(deck_path, labels)
    located = L.locate(public, internal, paragraphs)
    if internal is not None and internal.get("baseline"):
        diff = diff_sources(internal["baseline"], session.contracts)
        digest_changed = set(diff["changed_accounts"])
    else:
        diff = diff_by_hash(public, session.contracts)
        digest_changed = set(diff["changed_accounts"])
    report["source_diff"] = {k: v for k, v in diff.items() if k != "accounts"} | (
        {"accounts": {a: {k: v for k, v in d.items() if k in ("status", "periods_added", "periods_removed")} for a, d in diff["accounts"].items()}}
        if "accounts" in diff else {})
    actions, rows = [], []
    slides_before = [e["slide"] for e in public["elements"]]
    for element in public["elements"]:
        loc = located["elements"].get(element["id"], {"status": "unlocated"})
        changed, why = _source_changed(element, session.contracts, digest_changed)
        row = {"id": element["id"], "account": element["account"], "slide": element["slide"], "status": loc["status"], "source_changed": changed}
        if loc.get("slide_now") is not None:
            row["slide_now"] = loc["slide_now"]
        if loc["status"] in ("ambiguous", "unlocated"):
            action = {"element": element["id"], "action": "inspect_and_relink", "reason": loc.get("reason"), "candidates": loc.get("candidates")}
        elif loc["status"] in ("unchanged", "moved"):
            action = {"element": element["id"], "action": "auto_update" if changed else "keep",
                      "reason": why or ("moved to another slide, text unchanged" if loc["status"] == "moved" else "")}
        else:   # user_edited
            if changed:
                base = ((internal or {}).get("elements") or {}).get(element["id"], {}).get("text")
                action = {"element": element["id"], "action": "conflict", "reason": f"the user edited the paragraph and {why}",
                          "base_known": base is not None, "user_text": loc["text_now"]}
            else:
                action = {"element": element["id"], "action": "reverify",
                          "check": verify_user_text(loc["text_now"], session.contracts.get(element["account"])),
                          "reason": "kept as the user wrote it; its numbers are re-checked against the current sources"}
        rows.append(row)
        actions.append(action)
    tables = L.locate_tables(public.get("tables") or [], L.read_tables(deck_path))
    report["tables"] = tables
    for table_id, loc in tables.items():
        if loc["status"] == "user_edited":
            actions.append({"element": table_id, "action": "table_edited",
                            "reason": "a person changed this statement table; it is kept as returned and not rewritten here"})
        elif loc["status"] in ("ambiguous", "unlocated"):
            actions.append({"element": table_id, "action": "inspect_and_relink", "reason": loc.get("reason")})
    report["elements"] = rows
    report["located"] = located["elements"]
    report["actions"] = actions
    report["added_paragraphs"] = located["added_paragraphs"]
    report["duplicate_tags"] = located["duplicate_tags"]
    moved = [r for r in rows if r.get("slide_now") not in (None, r["slide"])]
    report["reordered_or_moved"] = bool(moved) and any(r["status"] == "moved" for r in rows)
    report["summary"] = {a: sum(x["action"] == a for x in actions) for a in sorted({x["action"] for x in actions})}
    return report


def _labels_from_session(session: Any) -> Dict[str, str]:
    out: Dict[str, str] = {}
    for key in session.contracts:
        conf = session.mappings.get(key) or {}
        for name in (conf.get("display_name_zh"), conf.get("display_name"), key):
            if name:
                out[str(name)] = key
    return out


# ---------------------------------------------------------------------------
# editing in place
# ---------------------------------------------------------------------------

def replace_body(paragraph: Any, old_body: str, new_body: str) -> bool:
    """Replace the commentary text of a '■ label - body' paragraph, keeping the label's run and
    the formatting of the first body run. False when the body cannot be located in the runs."""
    runs = list(paragraph.runs)
    full = "".join(r.text for r in runs)
    start = full.find(old_body)
    if start < 0 or not runs:
        return False
    position, first_done = 0, False
    for r in runs:
        end = position + len(r.text)
        if end <= start:
            position = end
            continue
        if not first_done:
            r.text = r.text[: max(start - position, 0)] + new_body
            first_done = True
        else:
            r._r.getparent().remove(r._r)
        position = end
    return first_done


def apply_updates(deck_path: str, session: Any, report: Dict[str, Any], new_texts: Dict[str, str], out_path: str) -> Dict[str, Any]:
    """Apply `auto_update` actions that have a new text, in place, to a COPY at out_path.
    Conflicts, relink cases and accounts that span several paragraphs are listed, not touched."""
    public = L.read_manifest(deck_path)
    if public is None:
        return {"applied": [], "skipped": [{"reason": "no manifest: nothing can be edited with confidence"}], "out_path": None}
    prs = Presentation(deck_path)
    by_id = {e["id"]: e for e in public["elements"]}
    parts_of: Dict[str, int] = {}
    for e in public["elements"]:
        parts_of[e["account"]] = parts_of.get(e["account"], 0) + 1
    rows = {r["id"]: r for r in report.get("elements", [])}
    applied, skipped = [], []
    for action in report.get("actions", []):
        eid = action["element"]
        element = by_id.get(eid)
        if action["action"] != "auto_update":
            if action["action"] in ("conflict", "inspect_and_relink"):
                skipped.append({"element": eid, "action": action["action"], "reason": action.get("reason")})
            continue
        if eid not in new_texts:
            skipped.append({"element": eid, "reason": "no new text was supplied for it"})
            continue
        if parts_of[element["account"]] > 1:
            skipped.append({"element": eid, "reason": "the account continues across paragraphs; changing its length needs the deck re-packed (reflow_required)"})
            continue
        row = rows[eid]
        slide = prs.slides[row["slide_now"] - 1 if row.get("slide_now") else element["slide"] - 1]
        shape = next((s for s in slide.shapes if s.shape_id == (report_shape_id(report, eid) or element["shape_id"])), None)
        paragraph = None
        if shape is not None:
            old_text = report_text_now(report, eid)
            for p in shape.text_frame.paragraphs:
                if old_text and old_text in p.text:
                    paragraph = p
                    break
        if paragraph is None or not replace_body(paragraph, report_text_now(report, eid), new_texts[eid]):
            skipped.append({"element": eid, "reason": "the paragraph could not be located in the shape's runs"})
            continue
        applied.append({"element": eid, "account": element["account"], "slide": row.get("slide_now") or element["slide"]})
    prs.save(out_path)
    return {"applied": applied, "skipped": skipped, "out_path": out_path,
            "touched_slides": sorted({a["slide"] for a in applied})}


def report_shape_id(report: Dict[str, Any], eid: str) -> Optional[int]:
    return (report.get("located", {}).get(eid) or {}).get("shape_id")


def report_text_now(report: Dict[str, Any], eid: str) -> str:
    return (report.get("located", {}).get(eid) or {}).get("text_now", "")
