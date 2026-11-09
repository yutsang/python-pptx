"""The other tables on an account's own tab: the ones beside or below the main schedule.

An account tab carries more than the schedule the pipeline parses. Beside it or under it the preparer lays out what the
report's tables are built from: the balance grouped by nature, an ageing split, the phases of a building, the terms of a
capital agreement, the adjusted view of an expense. The 2026-11-09 evidence locator found the team's report quoting 22 of 95
figures the writer never saw, and six of them sat in such a block on the same tab (none were on a hidden sheet).

This reads them, deterministically and without a model: every rectangle of filled cells on the tab outside the main
schedules, with a label column, a header row and at least three numbers. Hidden sheets are never read (the caller checks).
"""

from __future__ import annotations

import re
from collections import deque
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd

from ..financial_common import cell_text, coerce_numeric

Region = Tuple[int, int, int, int]          # first row, last row, first column, last column (0-based, inclusive)

_WORKING = ("check", "差异", "差額", "差额", "对账", "核对", "diff", "variance to", "tie-out", "勾稽")
_TOTAL = ("total", "合计", "总计", "小计", "合計", "總計", "小計")
_RATE_HEADER = ("%", "率", "rate", "比例", "占比")
_THOUSAND = ("千元", "'000", "’000", "k cny")
# columns that number or date an entry rather than measure it (a voucher listing has them all)
_CODE_HEADER = ("凭证", "编码", "编号", "代码", "序号", "期间", "日期", "code", "no.", "voucher", "date", "account no")
_TEN_THOUSAND = ("万元", "萬元")
_YUAN = ("单位：元", "单位:元", "(元)", "（元）", "人民币元", "rmb", "cny")


def column_letter(index: int) -> str:
    out = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        out = chr(65 + rem) + out
    return out


def where(region: Region) -> str:
    r0, r1, c0, c1 = region
    return f"{column_letter(c0)}{r0 + 1}:{column_letter(c1)}{r1 + 1}"


def _is_working(text: str) -> bool:
    low = str(text or "").strip().lower()
    return bool(low) and any(w in low for w in _WORKING)


def _filled(raw: pd.DataFrame, col_gap: int = 2) -> List[List[bool]]:
    """Which cells hold something. A tie-out row ('Check' and its differences, up to the next gap wider than `col_gap`) is
    left out: it sits right under a schedule or a table, and counted as filled it joins the table below to the row above
    and gives the table a check figure for a title."""
    grid = [[bool(cell_text(v)) for v in raw.iloc[r].tolist()] for r in range(len(raw))]
    for r in range(len(raw)):
        row = raw.iloc[r].tolist()
        c = 0
        while c < len(row):
            if grid[r][c] and coerce_numeric(row[c]) is None and _is_working(cell_text(row[c])):
                blank = 0
                while c < len(row) and blank <= col_gap:
                    blank = 0 if grid[r][c] else blank + 1
                    grid[r][c] = False
                    c += 1
                continue
            c += 1
    return grid


def _inside(r: int, c: int, regions: Iterable[Region]) -> bool:
    return any(r0 <= r <= r1 and c0 <= c <= c1 for r0, r1, c0, c1 in regions)


def _components(grid: List[List[bool]], exclude: Sequence[Region], row_gap: int = 1, col_gap: int = 2) -> List[Region]:
    """Bounding boxes of the filled cells outside `exclude`, two cells joined when they share a row with at most two blank
    columns between them, or a column with at most one blank row between them (a table has blank cells inside it). Never
    diagonally: two tables that only touch at a corner stay two tables."""
    rows = len(grid)
    seen = set()
    boxes: List[Region] = []
    for r in range(rows):
        for c in range(len(grid[r])):
            if not grid[r][c] or (r, c) in seen or _inside(r, c, exclude):
                continue
            queue = deque([(r, c)])
            seen.add((r, c))
            r0 = r1 = r
            c0 = c1 = c
            while queue:
                y, x = queue.popleft()
                r0, r1, c0, c1 = min(r0, y), max(r1, y), min(c0, x), max(c1, x)
                steps = [(0, dx) for dx in range(-col_gap - 1, col_gap + 2) if dx] + \
                        [(dy, 0) for dy in range(-row_gap - 1, row_gap + 2) if dy]
                for dy, dx in steps:
                    ny, nx = y + dy, x + dx
                    if (ny, nx) in seen or not (0 <= ny < rows) or not (0 <= nx < len(grid[ny])):
                        continue
                    if grid[ny][nx] and not _inside(ny, nx, exclude):
                        seen.add((ny, nx))
                        queue.append((ny, nx))
            boxes.append((r0, r1, c0, c1))
    return boxes


def _unit(texts: Iterable[str], default: float) -> float:
    blob = " ".join(texts).lower()
    if any(m in blob for m in _TEN_THOUSAND):
        return 10000.0
    if any(m.lower() in blob for m in _THOUSAND):
        return 1000.0
    if any(m.lower() in blob for m in _YUAN):
        return 1.0
    return default


def _period(text: str, periods: Sequence[str]) -> Optional[str]:
    """The account period a header names, or None. Only headers with date punctuation or a year-month shape count: a bare
    number in a header row is an amount far more often than a date."""
    from .statements import parse_date
    t = str(text or "").strip()
    if not t or not re.search(r"\d{4}\s*[-/.年]|\d{1,2}M\d{2}|FY\s*\d{2,4}", t, re.I):
        return _year_period(t, periods)
    parsed = parse_date(t, allow_bare_number=False)
    if parsed is None or not hasattr(parsed, "strftime"):
        return _year_period(t, periods)
    iso = parsed.strftime("%Y-%m-%d")
    return iso if iso in periods else ""


def _year_period(text: str, periods: Sequence[str]) -> Optional[str]:
    """A header or title that names only a year ('2020 Total', '2022'): that year's period, '' when the account reports no
    such year, None when no year is named. Read as the latest period instead, a pivot of 2020 by tenant was filed as the
    2022 figures on a local databook."""
    years = set(re.findall(r"(?<!\d)(20\d{2})(?!\d)", str(text or "")))
    if len(years) != 1:
        return None
    year = years.pop()
    matches = [p for p in periods if p.startswith(year)]
    return matches[-1] if matches else ""


def _title_period(texts: Sequence[str], periods: Sequence[str]) -> Optional[str]:
    """The period a table's title gives ('固定资产 - 截至2026年3月31日', '2022'), '' when it names one the account does not
    report, None when it names none."""
    for t in texts:
        m = re.search(r"\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日|\d{4}[-/.]\d{1,2}[-/.]\d{1,2}", str(t or ""))
        if m:
            q = _period(m.group(0), periods)
            if q is not None:
                return q
    for t in texts:
        q = _year_period(t, periods)
        if q is not None:
            return q
    return None


def read_block(raw: pd.DataFrame, region: Region, *, periods: Sequence[str], multiplier: float, max_rows: int = 25,
               max_columns: int = 6) -> Optional[Dict[str, Any]]:
    """One rectangle as a table: {title, where, unit, columns: [{header, period, assumed}], rows: [{label, values}]}.
    None when it is not a table of amounts (no label column, fewer than three numbers, a rate grid, a tie-out)."""
    r0, r1, c0, c1 = region
    cells = {(r, c): raw.iat[r, c] for r in range(r0, r1 + 1) for c in range(c0, c1 + 1)}
    text = {k: cell_text(v) for k, v in cells.items()}
    num = {k: coerce_numeric(v) for k, v in cells.items()}
    is_num = {k for k, v in num.items() if v is not None and text[k] and not re.fullmatch(r"\d{4}", text[k].strip())}
    # label column: the column with the most text cells that are not numbers
    label_col, best = None, 0
    for c in range(c0, c1 + 1):
        n = sum(1 for r in range(r0, r1 + 1) if text[(r, c)] and (r, c) not in is_num)
        if n > best:
            label_col, best = c, n
    if label_col is None or best < 2:
        return None
    # header row: the first row with at least two filled cells right of the label column and no more numbers than text
    def unit_only(t: str) -> bool:
        low = t.lower().replace(" ", "")
        return bool(low) and len(low) <= 12 and _unit([low], 0.0) != 0.0 and not re.search(r"\d{4}", low)
    header = None
    for r in range(r0, r1 + 1):
        right = [c for c in range(label_col + 1, c1 + 1) if text[(r, c)] and not unit_only(text[(r, c)])]
        if len(right) >= 1 and sum(1 for c in right if (r, c) in is_num) < len(right) and \
                any(any(cc > label_col and (rr, cc) in is_num for cc in range(label_col + 1, c1 + 1)) for rr in range(r + 1, r1 + 1)):
            header = r
            break
    if header is None:
        return None
    title_bits = [text[(r, c)] for r in range(r0, header) for c in range(c0, c1 + 1)
                  if text[(r, c)] and (r, c) not in is_num and not unit_only(text[(r, c)]) and not _is_working(text[(r, c)])]
    unit_bits = [text[(r, c)] for r in range(r0, header + 1) for c in range(c0, c1 + 1) if unit_only(text[(r, c)])]
    above = [t for t in (cell_text(raw.iat[r, c]) for r in range(max(0, r0 - 2), r0) for c in range(c0, min(c1 + 1, raw.shape[1])))
             if t and coerce_numeric(t) is None and not _is_working(t)]       # a check row's difference is not a title
    title = " ".join(above[-1:] + title_bits)[:80].strip() or text[(header, label_col)]
    unit_texts = above + title_bits + unit_bits + [text[(header, c)] for c in range(c0, c1 + 1)]
    unit = _unit(unit_texts, multiplier)
    unit_marked = _unit(unit_texts, 0.0) != 0.0
    from .inspector import canonical_stage_label
    # A second header row (the dates under a stage row) when the row under the header carries no number to the right.
    head_rows = [header]
    if header + 1 <= r1 and not any((header + 1, c) in is_num for c in range(label_col + 1, c1 + 1)) and \
            any(text[(header + 1, c)] for c in range(label_col + 1, c1 + 1)):
        head_rows.append(header + 1)
    first_data = head_rows[-1] + 1
    carried: Dict[int, str] = {}
    for h in head_rows[:-1] if len(head_rows) > 1 else []:
        last = ""
        for c in range(label_col + 1, c1 + 1):             # a merged stage cell names every column under it
            last = text[(h, c)] or last
            carried[c] = last
    columns: List[Dict[str, Any]] = []
    titled = _title_period(above[-1:] + title_bits + [text[(header, label_col)]], periods)   # a column naming no period takes the title's, else the latest
    for c in range(label_col + 1, c1 + 1):
        parts = [carried.get(c, "")] if len(head_rows) > 1 else []
        parts += [text[(head_rows[-1], c)] or (text[(header - 1, c)] if len(head_rows) == 1 and header - 1 >= r0 else "")]
        head = " ".join(t for t in parts if t).strip()
        if not text[(head_rows[-1], c)] and not (len(head_rows) == 1 and head):
            continue                                   # a column nobody named is a working calculation, not a measure
        if not head or any(h in head.lower() for h in _RATE_HEADER + _CODE_HEADER + _WORKING):
            continue
        if coerce_numeric(text[(head_rows[-1], c)]) is not None and not re.search(r"[A-Za-z一-鿿]", text[(head_rows[-1], c)]):
            continue                                   # a figure in the header row (a check's difference) names nothing
        values = [num[(r, c)] for r in range(first_data, r1 + 1) if (r, c) in is_num]
        nonzero = [v for v in values if v]
        if not values or (nonzero and all(abs(v) <= 1.0 for v in nonzero)):
            continue                                   # empty, or a column of rates / ratios (a nil year is kept: it says nil)
        p = None
        for t in parts:
            q = _period(t, periods)
            if q is not None:
                p = q
                break
        if p is None and titled is not None:
            p = titled
        if p == "":
            continue                                   # a date this account does not report (a month, an older year)
        columns.append({"col": c, "header": head, "period": p or (periods[-1] if periods else ""), "assumed": p is None,
                        "stage": next((canonical_stage_label(t) for t in parts if canonical_stage_label(t)), None)})
    if any(col["stage"] for col in columns):
        # Stage columns (audited, management, adjusted) are the same date read three ways: only the adjusted one is the
        # deck's figure. Without an adjusted column the stages are kept, each named by its header.
        adjusted = [col for col in columns if col["stage"] == "Indicative adjusted"]
        if adjusted:
            columns = adjusted
    if not columns:
        return None
    columns = columns[:max_columns]
    rows: List[Dict[str, Any]] = []
    for r in range(first_data, r1 + 1):
        label = text[(r, label_col)]
        if not label or any(w in label.lower() for w in _WORKING):
            continue
        values = {}
        for col in columns:
            v = num[(r, col["col"])]
            if v is not None and (r, col["col"]) in is_num:
                values[col["col"]] = round(v * unit, 2)
        if values:
            rows.append({"label": label, "row": r, "values": values, "total": any(t in label.lower() for t in _TOTAL)})
    if sum(len(x["values"]) for x in rows) < 3 or len(rows) < 2:
        return None
    if not any(abs(v) > 0.005 for x in rows for v in x["values"].values()):
        return None                                    # every figure nil: nothing to say
    entry_like = sum(1 for x in rows if re.fullmatch(r"\s*(\d{4}[/.-]\d{1,2}[/.-]\d{1,2}.*|\d{4,}\s*)", x["label"]))
    if entry_like * 2 >= len(rows):
        return None                                    # labelled by dates or codes: a listing of entries, not a breakdown
    if len(rows) > max_rows:
        keep = {id(x) for x in sorted(rows, key=lambda x: -max(abs(v) for v in x["values"].values()))[:max_rows]}
        keep |= {id(x) for x in rows if x["total"]}
        rows = [x for x in rows if id(x) in keep]
    return {"title": title, "where": where(region), "unit": unit, "unit_marked": unit_marked, "label_col": label_col, "header_row": header,
            "columns": [{k: v for k, v in col.items()} for col in columns], "rows": rows}


def _names_another_entity(title: str, entity: str) -> bool:
    """A title that ends ' - <entity>' naming a different company: a tab holding several entities' blocks side by side."""
    m = re.search(r"\s[-–—]\s*([^-–—]+)$", str(title or ""))
    if not m or not entity:
        return False
    name, own = m.group(1).strip().lower(), str(entity).strip().lower()

    def marks(text: str) -> set:
        cjk = re.findall(r"[一-鿿]+", text)
        return {run[i:i + 2] for run in cjk for i in range(len(run) - 1)} | set(re.findall(r"[a-z]{4,}", text))
    return bool(marks(name)) and not (marks(name) & marks(own))


def _is_template(block: Dict[str, Any]) -> bool:
    """A layout left unfilled: a placeholder date in the title ("2023年ZZ月ZZ日") or placeholder rows (客商1, 客商2). On the
    client file a model wrote "账龄分析显示…" from an empty ageing template (2026-11-09)."""
    if re.search(r"(?:ZZ|XX|xx)\s*[年月日]|\b20(?:ZZ|XX|xx)\b", str(block.get("title") or "")):
        return True
    labels = [str(x.get("label") or "").strip() for x in block.get("rows") or []]
    return sum(1 for label in labels if re.fullmatch(r"(?:客商|供应商|客户|公司|项目|租户)\s*\d+", label)) >= 2


def find_supporting_blocks(raw: pd.DataFrame, exclude: Sequence[Region], *, periods: Sequence[str], multiplier: float = 1.0,
                           entity: str = "", max_blocks: int = 4, scale: float = 0.0) -> List[Dict[str, Any]]:
    """The tables on a tab outside `exclude` (the main schedules), largest first, at most `max_blocks`. A table titled for
    another company (' - <name>') is left out.

    `scale` is the account's largest balance in yuan. A table that names no unit is read in the schedule's unit; when that
    makes its largest figure more than ten times the account, the table was kept in yuan (a working block under a 千元
    schedule on the DEMO databook read as 4,037亿 for a 4亿 transfer), so it is read in yuan, and left out if even that is
    too large."""
    if raw is None or raw.empty:
        return []
    grid = _filled(raw)
    out = []
    for region in _components(grid, exclude):
        r0, r1, c0, c1 = region
        if r1 - r0 < 2 or c1 - c0 < 1:
            continue
        block = read_block(raw, region, periods=periods, multiplier=multiplier)
        if block is None or _names_another_entity(block["title"], entity) or _is_template(block):
            continue
        block["size"] = max((abs(v) for x in block["rows"] for v in x["values"].values()), default=0.0)
        if scale > 0 and block["size"] > 10 * scale:
            if block.get("unit_marked") or multiplier == 1 or block["size"] / multiplier > 10 * scale:
                continue
            for x in block["rows"]:
                x["values"] = {c: round(v / multiplier, 2) for c, v in x["values"].items()}
            block.update(unit=1.0, size=block["size"] / multiplier, unit_inferred=True)
        out.append(block)
    out.sort(key=lambda b: -b["size"])
    return out[:max_blocks]


def main_region(attrs: Dict[str, Any], width: int, last_row: Optional[int] = None) -> Optional[Region]:
    """Where an account's main schedule sits on its tab: its rows (with the title and header rows above the data) and every
    column up to its rightmost value column, plus the remark columns beside it.

    `last_row` is the last sheet row the schedule took a figure from. The extractor's block end runs on to the last row with
    any number in the schedule's columns, so a table laid out under the schedule falls inside it: on the client file's
    2026-11-09 locator the fixed-asset block ended at row 61 for a schedule that ends at row 34, and the phase table under it
    (B40:F50) was excluded as part of the schedule; so were the payables table at W41:Y57 and the receivables block at
    B84:D87."""
    integrity = attrs.get("integrity") or {}
    start, end = integrity.get("block_start_row"), integrity.get("block_end_row")
    if start is None or end is None:
        return None
    if last_row is not None and int(start) <= int(last_row) < int(end):
        end = int(last_row)
    cols = [c.get("col_idx") for c in attrs.get("normalized_columns") or [] if c.get("col_idx") is not None]
    cols += [c.get("col_idx") for c in attrs.get("adjacent_detail_columns") or [] if c.get("col_idx") is not None]
    right = max(cols) if cols else width - 1
    return (max(0, int(start) - 6), int(end) + 1, 0, int(right))
