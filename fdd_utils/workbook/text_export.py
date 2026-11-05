from __future__ import annotations

"""
Utilities for exporting selected Excel tabs into compact plain text.
"""

from .preflight import _is_empty_value


from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable, List, Sequence, Tuple

import pandas as pd


@dataclass(frozen=True)
class TrimmedSheet:
    sheet_name: str
    dataframe: pd.DataFrame
    start_row: int  # zero-based
    start_col: int  # zero-based

    @property
    def end_row(self) -> int:
        return self.start_row + len(self.dataframe.index) - 1

    @property
    def end_col(self) -> int:
        return self.start_col + len(self.dataframe.columns) - 1

    @property
    def used_range(self) -> str:
        return (
            f"{_column_letter(self.start_col + 1)}{self.start_row + 1}:"
            f"{_column_letter(self.end_col + 1)}{self.end_row + 1}"
        )


def _column_letter(column_number: int) -> str:
    result = []
    current = column_number
    while current > 0:
        current, remainder = divmod(current - 1, 26)
        result.append(chr(65 + remainder))
    return "".join(reversed(result))


def _is_blank_or_na_value(value) -> bool:
    if pd.isna(value):
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def _normalize_cell_value(value) -> str:
    if _is_blank_or_na_value(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:.6f}".rstrip("0").rstrip(".")
    return str(value).strip()


def _trim_sheet(sheet_name: str, df: pd.DataFrame) -> TrimmedSheet:
    if df is None or df.empty:
        raise ValueError(f"Sheet '{sheet_name}' is empty.")

    non_empty_rows = [
        idx for idx in range(len(df.index)) if not df.iloc[idx].map(_is_empty_value).all()
    ]
    non_empty_cols = [
        idx for idx in range(len(df.columns)) if not df.iloc[:, idx].map(_is_empty_value).all()
    ]

    if not non_empty_rows or not non_empty_cols:
        raise ValueError(f"Sheet '{sheet_name}' is empty after trimming blank borders.")

    start_row = min(non_empty_rows)
    end_row = max(non_empty_rows)
    start_col = min(non_empty_cols)
    end_col = max(non_empty_cols)

    trimmed = df.iloc[start_row : end_row + 1, start_col : end_col + 1].copy()
    trimmed = trimmed.fillna("")

    return TrimmedSheet(
        sheet_name=sheet_name,
        dataframe=trimmed,
        start_row=start_row,
        start_col=start_col,
    )


def _validate_selected_tabs(all_sheet_names: Sequence[str], selected_tabs: Iterable[str]) -> List[str]:
    normalized_tabs = [str(tab).strip() for tab in selected_tabs if str(tab).strip()]
    if not normalized_tabs:
        raise ValueError("At least one sheet tab must be selected.")

    missing_tabs = [tab for tab in normalized_tabs if tab not in all_sheet_names]
    if missing_tabs:
        raise ValueError(
            f"Requested sheet tabs were not found: {', '.join(missing_tabs)}. "
            f"Available sheets: {', '.join(all_sheet_names)}"
        )

    return normalized_tabs


def _render_sheet_table(trimmed_sheet: TrimmedSheet) -> str:
    column_headers = ["ExcelRow"] + [
        f"Col{_column_letter(index + 1)}" for index in range(len(trimmed_sheet.dataframe.columns))
    ]
    lines = [
        f"===== SHEET: {trimmed_sheet.sheet_name} =====",
        f"USED_RANGE: {trimmed_sheet.used_range}",
        "| " + " | ".join(column_headers) + " |",
    ]

    for row_offset, (_, row) in enumerate(trimmed_sheet.dataframe.iterrows()):
        excel_row = trimmed_sheet.start_row + row_offset + 1
        cells = [_normalize_cell_value(value) for value in row.tolist()]
        lines.append("| " + " | ".join([str(excel_row), *cells]) + " |")

    return "\n".join(lines)


def render_selected_tabs_text(workbook_path: str, selected_tabs: Sequence[str]) -> str:
    workbook = Path(workbook_path)
    all_sheets = pd.read_excel(workbook, sheet_name=None, header=None, engine="openpyxl")
    tabs_to_export = _validate_selected_tabs(list(all_sheets.keys()), selected_tabs)

    rendered_sections = []
    for sheet_name in tabs_to_export:
        trimmed_sheet = _trim_sheet(sheet_name, all_sheets[sheet_name])
        rendered_sections.append(_render_sheet_table(trimmed_sheet))

    header = [
        f"WORKBOOK: {workbook.name}",
        f"EXPORTED_AT: {datetime.now().isoformat(timespec='seconds')}",
        f"SHEETS: {', '.join(tabs_to_export)}",
        "",
    ]
    return "\n\n".join(["\n".join(header).rstrip(), *rendered_sections]).strip() + "\n"


def export_selected_tabs_to_file(
    workbook_path: str,
    selected_tabs: Sequence[str],
    output_path: str | None = None,
) -> str:
    workbook = Path(workbook_path)
    destination = (
        Path(output_path)
        if output_path
        else workbook.with_name(f"{workbook.stem}_selected_tabs.txt")
    )
    rendered = render_selected_tabs_text(str(workbook), selected_tabs)
    destination.write_text(rendered, encoding="utf-8")
    return str(destination)
# --- end workbook/text_export.py ---
