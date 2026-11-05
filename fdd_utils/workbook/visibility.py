"""One rule for hidden sheets wherever a Financials sheet is chosen or read.

A hidden tab is a working copy, an old version or a scratch sheet. It is never the sheet every
account is reconciled against, and nothing read from it is offered as evidence. The rule lives
here so each place that picks or reads a Financials sheet applies the same one:

  * picking by name (the agent session, the CLI's auto-detection, the CLI's roll-up matching,
    the resolver's reference sheet) keeps only VISIBLE sheets, and keeps nothing when the
    workbook's sheet states cannot be read -- "all of them" is not a fallback
  * naming a sheet explicitly (--sheet, --financials-sheet, a roll-up choice) is refused if the
    sheet is hidden, or if it cannot be shown not to be
  * the one function that reads the Financials sheet (statements.extract_balance_sheet_and_
    income_statement) applies the refusal itself, so a caller that forgot cannot bypass it

Visible sheets behave exactly as before: they are picked in the same order by the same rules.
"""
from __future__ import annotations

import os
import re
from typing import Dict, Iterable, List, Optional, Tuple

_CACHE: Dict[Tuple[str, float, int], Optional[Dict[str, str]]] = {}


class HiddenSheetError(RuntimeError):
    """A hidden sheet was named, or visibility could not be read for a named sheet."""


def sheet_states(path: str) -> Optional[Dict[str, str]]:
    """{sheet name: 'visible' | 'hidden' | 'veryHidden'}, or None when the workbook's own state cannot be read."""
    try:
        stat = os.stat(path)
    except (OSError, TypeError):
        return None
    key = (str(path), stat.st_mtime, stat.st_size)
    if key in _CACHE:
        return _CACHE[key]
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            states: Optional[Dict[str, str]] = {ws.title: str(getattr(ws, "sheet_state", "visible") or "visible") for ws in wb.worksheets}
        finally:
            wb.close()
    except Exception:
        states = None
    _CACHE[key] = states
    return states


def visible_sheet_names(path: str) -> Optional[List[str]]:
    states = sheet_states(path)
    return None if states is None else [name for name, state in states.items() if state == "visible"]


def filter_visible(path: str, names: Iterable[str]) -> List[str]:
    """`names` without the hidden ones, order kept. Nothing at all when visibility cannot be read."""
    visible = visible_sheet_names(path)
    if visible is None:
        return []
    allowed = set(visible)
    return [n for n in names if n in allowed]


def hidden_among(path: str, names: Iterable[str]) -> List[str]:
    states = sheet_states(path) or {}
    return [n for n in names if states.get(n, "visible") != "visible"]


def refuse_hidden(path: str, sheet: Optional[str], what: str = "Financials") -> None:
    """Raise HiddenSheetError if `sheet` is hidden, or its visibility cannot be told. A path that is not
    a file name (an in-memory buffer) is outside what can be checked and is let through."""
    if not sheet or not isinstance(path, (str, os.PathLike)):
        return
    states = sheet_states(os.fspath(path))
    if states is None:
        raise HiddenSheetError(f"the {what} sheet '{sheet}' was named, but the workbook's sheet visibility could not be read, "
                               "so it cannot be shown not to be hidden")
    if sheet in states and states[sheet] != "visible":
        raise HiddenSheetError(f"the {what} sheet '{sheet}' is {states[sheet]}; hidden sheets are not used as evidence")


def find_visible_financials_sheet(path: str) -> Optional[str]:
    """The VISIBLE sheet named Financials (exact first, then one that contains it), or None."""
    pool = visible_sheet_names(path)
    if pool is None:
        return None
    exact = [n for n in pool if n.strip().lower() == "financials"]
    if exact:
        return exact[0]
    loose = [n for n in pool if re.search("financials", n, re.IGNORECASE)]
    return loose[0] if loose else None
