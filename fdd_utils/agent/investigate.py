"""The four reading tools an investigation needs beyond the evidence contract.

    read_source_block      the original cells around a fact, as the workbook holds them
    search_notes           visible sheets' remarks that mention something
    cross_account_evidence what other accounts the numbers tie this account to (and how many tests did not tie)
    adjustment_evidence    why a period's total is not the tab's: the entries, the gap, what is unexplained

Each result names where it came from (file, sheet, cell range), the source version it was read at, and what it
does NOT cover (`limits`). None of them returns a summary the model wrote: cells are cells, hits are the text of
the cell, edges are the tests the cross-account builder ran. A hidden sheet is refused or skipped by the same
rule as everywhere else (workbook/visibility.py).

A remark found by `search_notes` with an account named is filed on that account's contract as a note (an N ref,
kind `searched_remark`), so a cause can cite it the way it cites any other note. The note says where it was
found; whether it supports the cause is still the verifier's question, not the search's.
"""
from __future__ import annotations

import hashlib
import os
import re
import time
from typing import Any, Dict, List, Optional, Tuple

from ..ai import claims as cl
from ..ai import evidence_contract as ec
from ..workbook.visibility import HiddenSheetError, filter_visible, refuse_hidden, sheet_states
from .session import SourceSession

MAX_BLOCK_ROWS = 60
MAX_BLOCK_COLS = 40
LOOK_FOR_FIGURE_COLS = 80
MAX_CELL_CHARS = 200
MAX_NOTES_PER_ACCOUNT = 30

#: where a scan stops, and what it says when it does
SCAN_SHEETS = 80
SCAN_ROWS_PER_SHEET = 4000
SCAN_SECONDS = 30.0


def _tool_error(code: str, message: str, data: Optional[Dict[str, Any]] = None):
    from .tools import ToolError
    return ToolError(code, message, data)


def _account(session: SourceSession, name: Any) -> str:
    from .tools import _account as resolve
    return resolve(session, name)


def _source_stamp(session: SourceSession, path: str) -> Dict[str, Any]:
    return {"file": os.path.basename(path), "sha256": (session.binding.workbook_sha256 or "")[:12], "source_version": session.source_version}


def _col_letter(n: int) -> str:
    out = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out


def _path_for_sheet(session: SourceSession, sheet: str) -> str:
    """The Financials sheet may live in a different file (a roll-up); every other sheet is in the workbook."""
    if session.financials_from and sheet == session.financials_sheet:
        return session.financials_from
    return session.workbook_path


def _plain(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, (int, float, bool)):
        return value
    if hasattr(value, "isoformat"):
        return value.isoformat()
    text = str(value).replace("\n", " ").strip()
    return text if len(text) <= MAX_CELL_CHARS else text[:MAX_CELL_CHARS] + "..."


def _col_number(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _address(address: str) -> Tuple[str, int]:
    m = re.match(r"^\$?([A-Za-z]{1,3})\$?(\d+)$", str(address or "").strip())
    return (m.group(1).upper(), int(m.group(2))) if m else ("A", 1)


# ---------------------------------------------------------------------------
# read_source_block
# ---------------------------------------------------------------------------

def read_source_block(session: SourceSession, account: str = "", ref: str = "", sheet: str = "", row: int = 0,
                      rows_before: int = 2, rows_after: int = 8, max_cols: int = 14) -> Dict[str, Any]:
    """The workbook's own cells around a fact (ref), a searched remark, or a row of a named sheet."""
    from openpyxl import load_workbook

    located = {"how": "", "highlight": None, "value": None}
    if ref:
        key = _account(session, account)
        contract = session.contracts[key]
        item = contract.fact(ref) or contract.note(ref)
        if item is None:
            raise _tool_error("UNKNOWN_REF", f"no fact or note {ref!r} for {key}")
        sheet = str(item.get("sheet") or "")
        if item.get("address"):                                       # a remark found by search_notes
            _col, row = _address(item["address"])
        elif item.get("row") is not None:                             # the extractor's row is 0-based
            row = int(item["row"]) + 1
        else:
            row = 1
        located["how"] = f"{ref} ({item.get('basis') or item.get('kind')}, {item.get('period') or 'no period'})"
        located["highlight"] = row
        if isinstance(item.get("value"), (int, float)) and item.get("basis") != "derived":
            located["value"] = float(item["value"])
    elif account and not sheet:
        key = _account(session, account)
        sheet = next((str(f.get("sheet")) for f in session.contracts[key].facts if f.get("sheet")), "")
        row = row or 1
        located["how"] = f"top of {key}'s schedule"
    if not sheet:
        raise _tool_error("BAD_ARGS", "give a ref (with its account), an account, or a sheet and a row")
    if row < 1:
        row = 1
    path = _path_for_sheet(session, sheet)
    try:
        refuse_hidden(path, sheet)
    except HiddenSheetError as exc:
        raise _tool_error("HIDDEN_SHEET", str(exc))
    rows_before, rows_after = max(0, min(int(rows_before), 20)), max(0, min(int(rows_after), MAX_BLOCK_ROWS))
    max_cols = max(1, min(int(max_cols), MAX_BLOCK_COLS))
    first, last = max(1, row - rows_before), row + rows_after
    truncated = False
    if last - first + 1 > MAX_BLOCK_ROWS:
        last, truncated = first + MAX_BLOCK_ROWS - 1, True
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet not in wb.sheetnames:
            raise _tool_error("NO_SUCH_SHEET", f"no sheet {sheet!r}", {"visible_sheets": filter_visible(path, wb.sheetnames)[:40]})
        ws = wb[sheet]
        # The extractor records a fact's row, not its column. Find the cells of that row that hold the figure (as it is
        # printed: in yuan, thousands or ten-thousands) so the block can include them and say where they are.
        matched: List[Dict[str, Any]] = []
        if located["value"] is not None and located["highlight"]:
            wanted = abs(located["value"])
            for row_cells in ws.iter_rows(min_row=located["highlight"], max_row=located["highlight"], max_col=LOOK_FOR_FIGURE_COLS, values_only=True):
                for c_index, v in enumerate(row_cells, 1):
                    if isinstance(v, (int, float)) and not isinstance(v, bool):
                        for unit in (1.0, 1e3, 1e4):
                            if abs(abs(v) * unit - wanted) <= max(1.0, 1e-9 * wanted):
                                matched.append({"cell": f"{_col_letter(c_index)}{located['highlight']}", "value": v, "printed_in": {1.0: "yuan", 1e3: "thousands", 1e4: "ten-thousands"}[unit]})
                                break
            if matched:
                reach = max(_col_number(_address(m["cell"])[0]) for m in matched)
                max_cols = max(max_cols, min(reach + 1, MAX_BLOCK_COLS))
        top = []
        for r_index, cells in enumerate(ws.iter_rows(min_row=1, max_row=4, max_col=max_cols, values_only=True), 1):
            text = " | ".join(str(_plain(c)) for c in cells if c not in (None, ""))
            if text:
                top.append({"row": r_index, "text": text[:240]})
        rows = []
        for r_index, cells in enumerate(ws.iter_rows(min_row=first, max_row=last, max_col=max_cols, values_only=True), first):
            filled = {_col_letter(c): _plain(v) for c, v in enumerate(cells, 1) if v not in (None, "")}
            rows.append({"row": r_index, "cells": filled})
        sheet_rows = ws.max_row
    finally:
        wb.close()
    last = min(last, sheet_rows) if sheet_rows else last
    span = f"A{first}:{_col_letter(max_cols)}{max(first, last)}"
    return {
        "location": f"{sheet}!{span}", "sheet": sheet, "first_row": first, "last_row": last, "highlight_row": located["highlight"],
        "figure_found_in": matched or None,
        "found_by": located["how"] or "sheet and row", "rows": [r for r in rows if r["row"] <= max(first, last)], "sheet_top": top,
        "source": _source_stamp(session, path),
        "limits": {"max_rows": MAX_BLOCK_ROWS, "max_cols": MAX_BLOCK_COLS, "columns_read": max_cols, "truncated_rows": truncated,
                   "figure_cells": ("cells of the fact's row that hold its figure; several means the same figure sits in several columns "
                                    "(periods or blocks) and which one the extractor read is not recorded") if matched else
                                   ("the fact's figure was not found on its row as printed (yuan, thousands or ten-thousands)" if located["value"] is not None else None),
                   "sheet_rows": sheet_rows,
                   "values": "cached values as last saved by Excel; formulas are neither shown nor recalculated",
                   "scope": "visible sheets only; a block is context, not a citable fact (cite the F or N refs)"},
    }


# ---------------------------------------------------------------------------
# search_notes
# ---------------------------------------------------------------------------

def _terms(query: str) -> List[str]:
    return [t for t in re.split(r"[\s,;，；、/|]+", str(query or "").lower()) if len(t) >= 2 or re.search(r"[一-鿿]", t)]


def _scan_text(session: SourceSession) -> Dict[str, Any]:
    """Every string cell of every visible sheet, once per reading of the file. Bounded, and says when it was."""
    from openpyxl import load_workbook
    cache_key = ("text", session.source_version)
    cached = session.text_cache.get(cache_key)
    if cached is not None:
        return cached
    started = time.perf_counter()
    path = session.workbook_path
    states = sheet_states(path) or {}
    wb = load_workbook(path, read_only=True, data_only=True)
    cells: List[Dict[str, Any]] = []
    scanned, hidden, truncated, notes = 0, [], False, []
    try:
        visible = filter_visible(path, wb.sheetnames)
        hidden = [n for n in wb.sheetnames if n not in visible]
        for name in visible:
            if scanned >= SCAN_SHEETS or time.perf_counter() - started > SCAN_SECONDS:
                truncated = True
                notes.append("the scan stopped before every sheet was read")
                break
            ws = wb[name]
            scanned += 1
            for r_index, cols in enumerate(ws.iter_rows(max_row=SCAN_ROWS_PER_SHEET, max_col=40, values_only=True), 1):
                label = ""
                for c_index, v in enumerate(cols, 1):
                    if isinstance(v, str) and v.strip():
                        text = re.sub(r"\s+", " ", v).strip()
                        if not label and c_index <= 4:
                            label = text[:60]
                        if len(text) >= 4:
                            cells.append({"sheet": name, "address": f"{_col_letter(c_index)}{r_index}", "text": text[:600], "label": label})
            if (ws.max_row or 0) > SCAN_ROWS_PER_SHEET:
                truncated = True
                notes.append(f"{name}: only the first {SCAN_ROWS_PER_SHEET} rows were read")
    finally:
        wb.close()
    out = {"cells": cells, "sheets_scanned": scanned, "hidden": hidden, "truncated": truncated, "notes": notes,
           "seconds": round(time.perf_counter() - started, 2), "states_known": bool(states)}
    session.text_cache[cache_key] = out
    return out


def search_notes(session: SourceSession, query: str, account: str = "", max_hits: int = 12) -> Dict[str, Any]:
    terms = _terms(query)
    if not terms:
        raise _tool_error("BAD_ARGS", "the query has no searchable word")
    scan = _scan_text(session)
    key = _account(session, account) if account else ""
    max_hits = max(1, min(int(max_hits), 30))
    scored = []
    for cell in scan["cells"]:
        low = cell["text"].lower()
        found = [t for t in terms if t in low]
        if not found or (len(terms) > 1 and len(found) / len(terms) < 0.6):
            continue
        # a sentence about the thing beats a bare label that happens to contain it
        scored.append((len(found) / len(terms), min(len(cell["text"]), 160), cell))
    scored.sort(key=lambda x: (-x[0], -x[1], x[2]["sheet"], x[2]["address"]))
    hits = []
    contract = session.contracts[key] if key else None
    for score, _length, cell in scored[:max_hits]:
        hit = {"sheet": cell["sheet"], "address": cell["address"], "text": cell["text"][:300], "row_label": cell["label"] or None,
               "match": round(score, 2), "location": f"{cell['sheet']}!{cell['address']}"}
        if contract is not None:
            existing = next((n for n in contract.notes if (n.get("sheet"), n.get("address")) == (cell["sheet"], cell["address"])
                             or re.sub(r"\s+", " ", n.get("text", "")) == cell["text"][:400]), None)
            if existing is None and len(contract.notes) < MAX_NOTES_PER_ACCOUNT + 20:
                number = len(contract.notes) + 1
                existing = {"id": f"{key}|note|{number}", "kind": "searched_remark", "text": cell["text"][:400], "period": None,
                            "sheet": cell["sheet"], "address": cell["address"], "origin": "search", "query": str(query)[:80],
                            "ref": f"N{number}"}
                contract.notes.append(existing)
            if existing is not None:
                hit["ref"] = existing["ref"]
        hits.append(hit)
    limits = ["a hit is the text of one cell; whether it explains anything is for the verifier and the reader",
              "only visible sheets are read", f"at most {SCAN_ROWS_PER_SHEET} rows and 40 columns per sheet"] + scan["notes"]
    if contract is not None:
        limits.append(f"hits were filed on {key} as notes (N refs) so a cause can cite them")
    return {"query": query, "terms": terms, "hits": hits, "total_matches": len(scored),
            "searched": {"sheets_scanned": scan["sheets_scanned"], "hidden_sheets_skipped": len(scan["hidden"]),
                         "cells_scanned": len(scan["cells"]), "truncated": scan["truncated"], "seconds": scan["seconds"]},
            "source": _source_stamp(session, session.workbook_path), "limits": limits}


# ---------------------------------------------------------------------------
# cross_account_evidence
# ---------------------------------------------------------------------------

def _cross_facts(session: SourceSession) -> Dict[str, Any]:
    from ..ai.facts import build_cross_account_facts
    cache_key = ("cross", session.source_version)
    if cache_key not in session.text_cache:
        session.text_cache[cache_key] = build_cross_account_facts(session.dfs, financials=session.financials_by_key or None)
    return session.text_cache[cache_key]


def _brief(value: Any, depth: int = 0) -> Any:
    """A bounded copy of an edge's evidence: scalars as they are (long strings cut), at most eight keys or items of anything nested."""
    if isinstance(value, (int, float, bool)) or value is None:
        return value
    if isinstance(value, str):
        return value if len(value) <= 100 else value[:100] + "..."
    if depth >= 3:
        return "..."
    if isinstance(value, dict):
        items = list(value.items())
        out = {str(k): _brief(v, depth + 1) for k, v in items[:8]}
        if len(items) > 8:
            out["..."] = f"{len(items) - 8} more"
        return out
    if isinstance(value, (list, tuple)):
        out_list = [_brief(v, depth + 1) for v in list(value)[:8]]
        return out_list + ([f"... {len(value) - 8} more"] if len(value) > 8 else [])
    return str(value)[:100]


def _why_not_tied(edge: Dict[str, Any]) -> str:
    evidence = edge.get("evidence") if isinstance(edge.get("evidence"), dict) else {}
    if evidence.get("ambiguous"):
        return "more than one account fits the same row, so no single account was tied"
    differing = list(evidence.get("differing_periods") or [])
    if differing:
        return "the two sides differ in " + ", ".join(str(p) for p in differing[:6]) + (" ..." if len(differing) > 6 else "")
    if edge.get("periods") == []:
        return "no period was shared by both sides"
    return "the test did not pass for the periods it covers"


def cross_account_evidence(session: SourceSession, account: str, max_failed: int = 12) -> Dict[str, Any]:
    """The relationships other accounts have to this one that the cross-account builder tested. Two lists, kept apart:

    edges            relationships that passed (a tie-out, a shared name, a ratio): the numbers support them
    failed_checks    relationships that were looked for and did NOT hold, each with the test, what it compared and why it did
                     not tie. They say where a link was expected and is absent (a schedule that does not add up to its account,
                     a breakdown that differs from the total in named periods). That is a reason to say "does not tie" or to
                     leave something unresolved. It is never support for a link: nothing in a failed check is citable, and a
                     sentence resting on one stays unverified."""
    from ..ai.facts import cross_account_links_for
    key = _account(session, account)
    facts = _cross_facts(session)
    passed = cross_account_links_for(facts, key, passed_only=True)
    every = cross_account_links_for(facts, key, passed_only=False)
    failed = [e for e in every if not e.get("passed")]
    not_tied: Dict[str, int] = {}
    for edge in failed:
        not_tied[str(edge.get("kind"))] = not_tied.get(str(edge.get("kind")), 0) + 1

    def other(edge: Dict[str, Any]) -> List[str]:
        names = []
        for side in (edge.get("source_ref") or {}, edge.get("target_ref") or {}):
            account_name = side.get("account")
            if account_name and account_name != key and account_name not in names:
                names.append(str(account_name))
        return names

    shown = max(0, min(int(max_failed), 40))
    edges = [{"edge_id": e.get("edge_id"), "kind": e.get("kind"), "with": other(e), "periods": list(e.get("periods") or []),
              "test": e.get("test"), "evidence": str(e.get("evidence"))[:400]} for e in passed]
    failed_checks = [{"edge_id": e.get("edge_id"), "kind": e.get("kind"), "with": other(e), "periods": list(e.get("periods") or []),
                      "test": e.get("test"), "result": "did not tie", "why": _why_not_tied(e), "compared": _brief(e.get("evidence")),
                      "supports_a_claim": False} for e in failed[:shown]]
    return {
        "account": key, "edges": edges,
        "failed_checks": failed_checks, "failed_checks_omitted": max(0, len(failed) - len(failed_checks)),
        "tested_and_not_tied": not_tied,
        "other_accounts": sorted({a for e in edges for a in e["with"]}),
        "coverage": {"accounts_examined": len(session.contracts), "relationships_tested_in_workbook": len(facts.get("links") or []) if isinstance(facts, dict) else None,
                     "touching_this_account": len(every), "passed": len(passed), "failed": len(failed)},
        "source": _source_stamp(session, session.workbook_path),
        "limits": [
            "an edge that passed is a relationship the numbers support (a tie-out, a shared name, a ratio), not a cause",
            "failed_checks are relationships that were looked for and did not hold, with what was compared: they explain a gap or a "
            "'does not tie'; they are not evidence that a relationship exists, cannot be cited, and a sentence resting on one is not verified",
            "to use another account's figure, fetch that account's evidence; a sentence citing another account's fact is "
            "marked unverified by the verifier (CROSS_ACCOUNT_UNBOUND) until a person confirms the figure belongs to it",
            "the tests are the cross-account builder's own (names, sums, ratios); a relationship it has no test for is not listed",
            "evidence is trimmed to eight keys or items per level; failed_checks_omitted counts what max_failed left out"],
    }


# ---------------------------------------------------------------------------
# adjustment_evidence
# ---------------------------------------------------------------------------

def adjustment_evidence(session: SourceSession, account: str) -> Dict[str, Any]:
    key = _account(session, account)
    contract = session.contracts[key]
    periods = []
    for p in contract.periods:
        raw = next((f for f in contract.facts if f["period"] == p and f["role"] == ec.ROLE_TOTAL and f["basis"] == ec.BASIS_RAW), None)
        fin = next((f for f in contract.facts if f["period"] == p and f["role"] == ec.ROLE_TOTAL and f["basis"] == ec.BASIS_FINANCIALS), None)
        used = next((f for f in contract.facts if f["period"] == p and f["role"] == ec.ROLE_TOTAL and f["basis"] == ec.BASIS_ADJUSTED), None)
        check = contract.check(p, ec.CHECK_TOTAL_VS_FINANCIALS) or {}
        replaced = bool(check.get("overridden"))
        difference = (raw["value"] - fin["value"]) if raw is not None and fin is not None else None
        entries = [n for n in contract.notes if n.get("kind") == "adjustment" and n.get("period") == p]
        entries += [n for n in contract.notes if n.get("kind") == "adjustment" and not n.get("period")]
        amounts = []
        for n in entries:
            amounts.append([round(a["value"], 2) for a in cl.text_amounts(n.get("text", ""))])
        listed = None
        if entries and all(len(a) == 1 for a in amounts):
            listed = round(sum(a[0] for a in amounts), 2)
        gap = difference is not None and abs(difference) > 1.0
        periods.append({
            "period": p, "tab_total": raw and {"ref": raw["ref"], "value": raw["value"], "sheet": raw.get("sheet"), "row": raw.get("row")},
            "financials_total": fin and {"ref": fin["ref"], "value": fin["value"]},
            "deck_total": used and {"ref": used["ref"], "value": used["value"]},
            "check": check.get("status"), "replaced_by_financials": replaced, "difference_tab_minus_financials": difference,
            "adjustment_entries": [{"ref": n["ref"], "text": n["text"], "sheet": n.get("sheet"), "period": n.get("period")} for n in entries],
            "amounts_in_entries": amounts,
            "entries_add_up_to_difference": (None if listed is None or difference is None else abs(abs(listed) - abs(difference)) <= max(1.0, 0.005 * abs(difference))),
            "evidence_gap": bool(gap and not entries),
            "permits": {k: contract.permissions.get(p, {}).get(k) for k in ("describe_total", "describe_difference", "explain_with_detail", "detail_basis")},
        })
    return {
        "account": key, "periods": periods,
        "periods_with_a_gap_and_no_entry": [x["period"] for x in periods if x["evidence_gap"]],
        "source": _source_stamp(session, session.workbook_path),
        "limits": [
            "an adjustment entry is text found in the workbook; the amounts in it are read from the text and may be in another unit",
            "entries_add_up_to_difference is null when an entry carries no single amount; nothing is guessed",
            "a gap with no entry means no adjustment note was found on the visible sheets, not that none exists"],
    }
