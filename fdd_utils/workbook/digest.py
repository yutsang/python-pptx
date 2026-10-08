"""Every tab of the workbook in memory: what it holds, where, and what reached the model.

"Understand 100% of the databook" has to be a number before it can be a goal.
The digest indexes EVERY non-empty cell of EVERY sheet -- mapped or not --
with its position, the block it sits in and the labels beside and above it,
and, once the AI frames exist, whether it reached the model (mark_reached).
The coverage ledger it yields is that number.

Measured on four local workbooks before this existed: 1-4 tabs per workbook
reached nothing downstream (supporting schedules -- sales by customer,
headcount, a stamp-duty detail tab, a construction ledger), and inside the
mapped tabs roughly 3 in 10 non-zero numbers reached the model.

What it is built from, so it cannot drift from what production read:
  - the raw frames production already loaded (load_workbook_frames, cached);
  - the profile production already computed (title / stage / date rows);
  - for a mapped sheet, the AI frame's own provenance: the sheet rows in its
    __source_row_idx column, the sheet columns of its analysis stage in
    attrs["normalized_columns"], the unit multiplier.

Positions are raw-frame indices (row r, column c of
pd.read_excel(header=None)), the same indices integrity rows and
__source_row_idx use. Plain dicts and lists only: it rides on `resolution`
and is written into a run folder as JSON.
"""
from __future__ import annotations

import bisect
import math
import os
from typing import Any, Dict, Iterable, List, Optional, Tuple

from .inspector import contains_unit_marker

DIGEST_KEY = "workbook_digest"

#: reach codes, per cell
REACH_FRAME = "f"       # inside the AI frame: a source row x an analysis-stage column
REACH_ADJACENT = "a"    # an adjacent detail row x one of its columns
REACH_VALUE = "v"       # not by position: the value appears in an attr the prompt renders
REACH_DESCRIPTION = "d" # text: a row description the model is shown
REACH_NOTE = "n"        # text: inside the notes / remarks / detail table the prompt renders

#: The attrs prompts.py renders that carry figures and text BEYOND the analysis
#: frame (whose cells are matched by position). Deliberately not the verifier's
#: pool: _attr_text_blob walks every attr but one, including
#: projection_original_values_by_description, auxiliary_check_totals_by_date and
#: any_period_nonzero_by_description, none of which prompts.py reads. A cell that
#: matches only those was never shown to the model.
_PROMPT_ATTRS = ("presentation_detail_table", "adjacent_detail_rows", "supporting_notes", "table_linked_remarks")
_INDEX_FIELDS = frozenset({"__source_row_idx", "row_idx", "col_idx", "level", "indent", "depth", "order"})
_REMARK_MIN_CHARS = 15


# -- reading a raw frame ------------------------------------------------------

def _number(value: Any) -> Optional[float]:
    if value is None or isinstance(value, (bool, str)):
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    return number if math.isfinite(number) else None


def _text(value: Any) -> Optional[str]:
    """Stripped text of a non-numeric cell; None for an empty one. A date or
    timestamp header is text here (ISO), not a number."""
    if value is None:
        return None
    if isinstance(value, float) and value != value:
        return None
    iso = getattr(value, "isoformat", None)
    if callable(iso) and not isinstance(value, str):
        try:
            return iso()[:10]
        except Exception:
            pass
    text = str(value).strip()
    return text or None


def _components(points: List[Tuple[int, int]]) -> List[int]:
    """Union-find over non-empty cells. Two cells share a block when they sit
    in the same or an adjacent row and at most two columns apart: a blank ROW
    ends a block (notes under a table are their own block), a single blank
    spacer COLUMN between stage groups does not."""
    index = {p: i for i, p in enumerate(points)}
    parent = list(range(len(points)))

    def find(i: int) -> int:
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    for i, (r, c) in enumerate(points):
        for dr in (0, 1):
            for dc in range(-2, 3):
                if dr == 0 and dc <= 0:
                    continue
                j = index.get((r + dr, c + dc))
                if j is not None:
                    a, b = find(i), find(j)
                    if a != b:
                        parent[b] = a
    return [find(i) for i in range(len(points))]


def _digest_sheet(name: str, frame: Any, profile: Dict[str, Any]) -> Dict[str, Any]:
    num_r: List[int] = []
    num_c: List[int] = []
    num_v: List[float] = []
    texts: List[List[Any]] = []
    blank = 0
    for r, row in enumerate(frame.itertuples(index=False)):
        for c, value in enumerate(row):
            number = _number(value)
            if number is not None:
                num_r.append(r)
                num_c.append(c)
                num_v.append(number)
                continue
            text = _text(value)
            if text is None:
                if value is not None and not (isinstance(value, float) and value != value):
                    blank += 1  # a whitespace-only string: non-empty to pandas, nothing to a reader
                continue
            texts.append([r, c, text])

    points = [(r, c) for r, c in zip(num_r, num_c)] + [(t[0], t[1]) for t in texts]
    roots = _components(points)
    block_of_root: Dict[int, int] = {}
    spans: List[List[int]] = []
    for (r, c), root in zip(points, roots):
        if root not in block_of_root:
            block_of_root[root] = len(spans)
            spans.append([r, r, c, c, 0, 0])
        s = spans[block_of_root[root]]
        s[0], s[1], s[2], s[3] = min(s[0], r), max(s[1], r), min(s[2], c), max(s[3], c)
    n_num = len(num_r)
    num_b = [block_of_root[roots[i]] for i in range(n_num)]
    for i, b in enumerate(num_b):
        spans[b][4] += 1
    for i, t in enumerate(texts):
        b = block_of_root[roots[n_num + i]]
        spans[b][5] += 1
        t.append(b)

    # Labels: for each row holding numbers, the text nearest to the left of its
    # first number; for each column, the text nearest above its first number.
    text_at = {(t[0], t[1]): i for i, t in enumerate(texts)}
    first_c_in_row: Dict[int, int] = {}
    first_r_in_col: Dict[int, int] = {}
    for r, c in zip(num_r, num_c):
        first_c_in_row[r] = min(c, first_c_in_row.get(r, c))
        first_r_in_col[c] = min(r, first_r_in_col.get(c, r))
    label_of_row: Dict[int, int] = {}
    for r, c0 in first_c_in_row.items():
        for c in range(c0 - 1, -1, -1):
            if (r, c) in text_at:
                label_of_row[r] = text_at[(r, c)]
                break
    label_of_col: Dict[int, int] = {}
    for c, r0 in first_r_in_col.items():
        for r in range(r0 - 1, max(-1, r0 - 7), -1):
            if (r, c) in text_at:
                label_of_col[c] = text_at[(r, c)]
                break

    header_rows = {profile.get("stage_row_idx"), profile.get("date_row_idx")} - {None}
    title_row = profile.get("title_row_idx")
    label_ids = set(label_of_row.values())
    header_ids = set(label_of_col.values())
    rows_with_numbers = set(first_c_in_row)
    for i, t in enumerate(texts):
        r, c, text = t[0], t[1], t[2]
        if r == title_row:
            role = "title"
        elif i in header_ids or r in header_rows:
            role = "header"
        elif i in label_ids:
            role = "label"
        elif contains_unit_marker(text):
            role = "unit"
        elif (r in rows_with_numbers and c > first_c_in_row[r]) or len(text) >= _REMARK_MIN_CHARS:
            role = "remark"
        else:
            role = "other"
        t.append(role)

    nonzero = sum(1 for v in num_v if v != 0)
    return {
        "kind": profile.get("sheet_kind"),
        "hidden": bool(profile.get("is_hidden")),
        "title": profile.get("title"),
        "unit_markers": list(profile.get("unit_markers") or []),
        "shape": [int(frame.shape[0]), int(frame.shape[1])],
        "blocks": [
            {"id": i, "rows": [s[0], s[1]], "cols": [s[2], s[3]], "numeric": s[4], "text": s[5]}
            for i, s in enumerate(spans)
        ],
        # Parallel arrays, one entry per numeric cell (zero included).
        "cells": {"r": num_r, "c": num_c, "v": num_v, "block": num_b, "reach": [""] * n_num},
        # One [row, col, text, block, role, reach] per text cell.
        "texts": [t + [""] for t in texts],
        "row_labels": {str(r): texts[i][2] for r, i in sorted(label_of_row.items())},
        "col_labels": {str(c): texts[i][2] for c, i in sorted(label_of_col.items())},
        "coverage": {
            "numeric": n_num, "numeric_nonzero": nonzero, "text": len(texts), "blank_strings": blank,
            "reached_nonzero": 0, "reached_by": {}, "text_reached": 0,
        },
    }


# -- building ----------------------------------------------------------------

def _hidden_names(profiles: Dict[str, Dict[str, Any]]) -> set:
    return {name for name, p in profiles.items() if (p or {}).get("is_hidden")}


def build_workbook_digest(
    workbook_path: str,
    profiles: Dict[str, Dict[str, Any]],
    resolution: Dict[str, Any],
    workbook_frames: Dict[str, Any],
) -> Dict[str, Any]:
    """Every sheet indexed; who (if anyone) reads it. Reach is filled later by
    mark_reached, once the frames the model will read exist."""
    profiles = profiles or {}
    mapped: Dict[str, List[str]] = {}
    for key, resolved in (resolution.get("resolved") or {}).items():
        if isinstance(resolved, dict) and resolved.get("sheet_name"):
            mapped.setdefault(str(resolved["sheet_name"]), []).append(str(key))
    unresolved = {
        str(e.get("sheet_name")): str(e.get("reason") or "")
        for e in (resolution.get("unresolved_sheets") or []) if isinstance(e, dict)
    }
    errors = resolution.get("normalization_errors") or {}
    sheets: Dict[str, Any] = {}
    for name, frame in (workbook_frames or {}).items():
        sheet = _digest_sheet(str(name), frame, profiles.get(name) or {})
        if name in mapped:
            sheet["status"], sheet["accounts"] = "mapped", mapped[name]
        elif sheet["kind"] == "financial_summary":
            sheet["status"], sheet["accounts"] = "financials", []
            sheet["reason"] = "read by reconciliation, never shown to the model"
        else:
            sheet["status"], sheet["accounts"] = "unmapped", []
            sheet["reason"] = unresolved.get(str(name)) or "no mapping resolved to this sheet"
        if name in errors:
            sheet["normalization_error"] = str(errors[name])[:200]
        sheets[str(name)] = sheet
    return {
        "workbook": os.path.basename(str(workbook_path)),
        "hidden_sheets": sorted(_hidden_names(profiles)),
        "reach_marked": False,
        "sheets": sheets,
        "coverage": _rollup(sheets),
    }


def _rows_of(frame: Any, row_key: str) -> set:
    if frame is None or row_key not in getattr(frame, "columns", []):
        return set()
    out = set()
    for v in frame[row_key].tolist():
        number = _number(v)
        if number is not None:
            out.add(int(number))
    return out


def _descriptions(frame: Any) -> set:
    if frame is None or not len(getattr(frame, "columns", [])):
        return set()
    return {str(v).strip() for v in frame.iloc[:, 0].tolist() if _text(v)}


def _prompt_material(attrs: Dict[str, Any]) -> Tuple[List[float], str]:
    """Numbers and text in the attrs the prompt renders. Index fields (row and
    column positions, levels) are skipped: a row number 7 is not a figure 7."""
    from ..ai.validator import _numbers_in_text

    numbers: List[float] = []
    texts: List[str] = []

    def walk(value: Any, depth: int = 0) -> None:
        if depth > 6 or value is None:
            return
        if isinstance(value, str):
            texts.append(value)
            numbers.extend(_numbers_in_text(value))
            return
        if isinstance(value, dict):
            for k, v in value.items():
                if str(k) not in _INDEX_FIELDS:
                    walk(v, depth + 1)
            return
        if isinstance(value, (list, tuple)):
            for v in value:
                walk(v, depth + 1)
            return
        number = _number(value)
        if number is not None:
            numbers.append(number)

    for key in _PROMPT_ATTRS:
        walk(attrs.get(key))
    return sorted(numbers), "\n".join(texts)


def _near(pool: List[float], target: float, tol: float) -> bool:
    j = bisect.bisect_left(pool, target - tol)
    return j < len(pool) and abs(pool[j] - target) <= tol


def mark_reached(digest: Dict[str, Any], dfs: Dict[str, Any]) -> Dict[str, Any]:
    """Fill each cell's reach from the frames the model reads (process_workbook_data's
    dfs). In place; returns the digest. A cell the model never sees keeps ""."""
    if not isinstance(digest, dict) or not digest.get("sheets"):
        return digest
    from . import INTERNAL_ROW_KEY

    by_sheet: Dict[str, List[Tuple[str, Any]]] = {}
    for key, df in (dfs or {}).items():
        sheet = (getattr(df, "attrs", None) or {}).get("source_sheet_name")
        if sheet:
            by_sheet.setdefault(str(sheet), []).append((str(key), df))

    for name, sheet in digest["sheets"].items():
        cells, texts, cov = sheet["cells"], sheet["texts"], sheet["coverage"]
        reach = [""] * len(cells["r"])
        for t in texts:
            t[5] = ""
        main_blocks = []
        for key, df in by_sheet.get(name, []):
            attrs = df.attrs or {}
            nested = attrs.get("prompt_analysis_df")
            rows = _rows_of(df, INTERNAL_ROW_KEY) | _rows_of(nested, INTERNAL_ROW_KEY)
            ncols = [c for c in (attrs.get("normalized_columns") or []) if isinstance(c, dict) and "col_idx" in c]
            stage = attrs.get("prompt_analysis_stage")
            frame_dates = {str(c) for c in list(df.columns) + list(getattr(nested, "columns", []))}
            ai_cols = {int(c["col_idx"]) for c in ncols if (c.get("stage") == stage if stage else str(c.get("date")) in frame_dates)}
            adj_cols = {int(c["col_idx"]) for c in (attrs.get("adjacent_detail_columns") or [])
                        if isinstance(c, dict) and c.get("col_idx") is not None}
            adj_rows = {int(r[INTERNAL_ROW_KEY]) for r in (attrs.get("adjacent_detail_rows") or [])
                        if isinstance(r, dict) and _number(r.get(INTERNAL_ROW_KEY)) is not None}
            mult = _number(attrs.get("source_multiplier")) or 1.0
            try:
                pool, blob = _prompt_material(attrs)
            except Exception:
                pool, blob = [], ""
            integ = attrs.get("integrity") or {}
            main_blocks.append({
                "account": key,
                "rows": [integ.get("block_start_row"), integ.get("block_end_row")],
                "ai_stage": stage,
                "ai_cols": sorted(ai_cols),
                "periods": {str(int(c["col_idx"])): c.get("date") for c in ncols if int(c["col_idx"]) in ai_cols},
                "multiplier": mult,
                "frame_rows": len(rows),
            })
            for i, (r, c, v) in enumerate(zip(cells["r"], cells["c"], cells["v"])):
                if reach[i] or v == 0:
                    continue
                if r in rows and c in ai_cols:
                    reach[i] = REACH_FRAME
                elif r in adj_rows and c in adj_cols:
                    reach[i] = REACH_ADJACENT
                elif pool and (_near(pool, v, max(0.005, abs(v) * 1e-9))
                               or _near(pool, v * mult, max(0.5, abs(v * mult) * 1e-9))):
                    # as written in the sheet, or scaled to base units
                    reach[i] = REACH_VALUE
            described = _descriptions(df) | _descriptions(nested)
            for t in texts:
                if t[5]:
                    continue
                text = t[2]
                if text in described:
                    t[5] = REACH_DESCRIPTION
                elif len(text) >= 2 and blob and text in blob:
                    t[5] = REACH_NOTE
        cells["reach"] = reach
        if main_blocks:
            sheet["main_blocks"] = main_blocks
        by: Dict[str, int] = {}
        for code, v in zip(reach, cells["v"]):
            if code and v != 0:
                by[code] = by.get(code, 0) + 1
        cov["reached_nonzero"] = sum(by.values())
        cov["reached_by"] = by
        cov["text_reached"] = sum(1 for t in texts if t[5])
    digest["reach_marked"] = True
    digest["coverage"] = _rollup(digest["sheets"])
    return digest


def _rollup(sheets: Dict[str, Any]) -> Dict[str, Any]:
    total = {"sheets": len(sheets), "numeric_nonzero": 0, "reached_nonzero": 0, "text": 0, "text_reached": 0,
             "reached_by": {}}
    unread = []
    for name, s in sheets.items():
        cov = s["coverage"]
        for k in ("numeric_nonzero", "reached_nonzero", "text", "text_reached"):
            total[k] += int(cov.get(k) or 0)
        for code, n in (cov.get("reached_by") or {}).items():
            total["reached_by"][code] = total["reached_by"].get(code, 0) + int(n)
        if s.get("status") != "mapped":
            unread.append({"sheet": name, "status": s.get("status"), "reason": s.get("reason"),
                           "numeric_nonzero": cov["numeric_nonzero"], "text": cov["text"]})
    total["unread_sheets"] = unread
    return total


# -- reading it back -----------------------------------------------------------

def coverage_rows(digest: Dict[str, Any]) -> List[Dict[str, Any]]:
    """One row per sheet, mapped first, then the ones nothing reads."""
    rows = []
    for name, s in (digest or {}).get("sheets", {}).items():
        cov = s["coverage"]
        nz = cov["numeric_nonzero"]
        rows.append({
            "sheet": name, "status": s.get("status"), "accounts": s.get("accounts") or [],
            "kind": s.get("kind"), "hidden": bool(s.get("hidden")), "title": s.get("title"),
            "blocks": len(s.get("blocks") or []),
            "numeric_nonzero": nz, "reached_nonzero": cov["reached_nonzero"],
            "reached_pct": (round(100.0 * cov["reached_nonzero"] / nz, 1) if nz else None),
            "reached_by": cov.get("reached_by") or {},
            "text": cov["text"], "text_reached": cov["text_reached"],
            "reason": s.get("reason"),
        })
    rows.sort(key=lambda r: (r["status"] != "mapped", r["sheet"]))
    return rows


def format_coverage(digest: Dict[str, Any]) -> List[str]:
    """Printable coverage ledger. Says when reach was never marked, rather than
    printing a column of zeros that reads as 'nothing reached the model'."""
    if not digest:
        return ["(no workbook digest)"]
    lines = []
    if not digest.get("reach_marked"):
        lines.append("(reach not marked: no AI frames were given; counts below are what the tabs hold)")
    lines.append(f"{'sheet':<28} {'status':<10} {'blocks':>6} {'nonzero#':>9} {'reached':>8} {'%':>6}  "
                 f"{'text':>5} {'text->AI':>8}  by position / by value")
    for row in coverage_rows(digest):
        pct = "-" if row["reached_pct"] is None else f"{row['reached_pct']:.0f}"
        by = row["reached_by"]
        detail = (f"frame {by.get(REACH_FRAME, 0)}, adjacent {by.get(REACH_ADJACENT, 0)}, value {by.get(REACH_VALUE, 0)}"
                  if row["status"] == "mapped" else (row["reason"] or ""))
        lines.append(f"{row['sheet'][:28]:<28} {row['status']:<10} {row['blocks']:>6} {row['numeric_nonzero']:>9} "
                     f"{row['reached_nonzero']:>8} {pct:>6}  {row['text']:>5} {row['text_reached']:>8}  {detail}")
    cov = digest.get("coverage") or {}
    nz = cov.get("numeric_nonzero") or 0
    by = cov.get("reached_by") or {}
    lines.append(f"TOTAL: {cov.get('reached_nonzero', 0)} of {nz} non-zero numbers reach the model"
                 + (f" ({100.0 * cov.get('reached_nonzero', 0) / nz:.0f}%: by position "
                    f"{by.get(REACH_FRAME, 0) + by.get(REACH_ADJACENT, 0)}, by value {by.get(REACH_VALUE, 0)})" if nz else "")
                 + f"; {cov.get('text_reached', 0)} of {cov.get('text', 0)} text cells; "
                 f"{len(cov.get('unread_sheets') or [])} of {cov.get('sheets', 0)} tabs read by no account")
    return lines


def sheet_view(digest: Dict[str, Any], sheet: str, max_items: int = 40) -> Optional[Dict[str, Any]]:
    """What one tab holds, for a reader: status, blocks, labels, remarks, and
    the largest non-zero figures with whether each reached the model."""
    s = (digest or {}).get("sheets", {}).get(sheet)
    if s is None:
        return None
    cells = s["cells"]
    figures = sorted(
        (i for i, v in enumerate(cells["v"]) if v != 0), key=lambda i: -abs(cells["v"][i])
    )[:max_items]
    return {
        "sheet": sheet, "status": s.get("status"), "accounts": s.get("accounts"), "reason": s.get("reason"),
        "kind": s.get("kind"), "title": s.get("title"), "shape": s.get("shape"), "blocks": s.get("blocks"),
        "main_blocks": s.get("main_blocks"), "coverage": s.get("coverage"),
        "row_labels": dict(list(s.get("row_labels", {}).items())[:max_items]),
        "col_labels": s.get("col_labels"),
        "remarks": [t[2] for t in s.get("texts", []) if t[4] == "remark"][:max_items],
        "largest_figures": [
            {"row": cells["r"][i], "col": cells["c"][i], "value": cells["v"][i],
             "row_label": s.get("row_labels", {}).get(str(cells["r"][i])),
             "col_label": s.get("col_labels", {}).get(str(cells["c"][i])),
             "reached": cells["reach"][i] or None}
            for i in figures
        ],
    }
