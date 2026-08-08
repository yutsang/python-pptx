from __future__ import annotations

# re-added: bound by an import in another section of the pre-split module
from ..keyword_registry import (
    BS_END_KEYWORDS,
    BS_HEADER_KEYWORDS,
    INDICATIVE_KEYWORDS,
    IS_END_KEYWORDS,
    IS_HEADER_KEYWORDS,
    REMARK_KEYWORDS,
    SUBTOTAL_KEYWORDS,
    SUMMARY_ACCOUNT_SKIP_KEYWORDS,
    TABLE_END_KEYWORDS,
    contains_thousand_unit_marker,
)

"""
FDD Table Inspector - Understand databook table structure and number flow.
Adapts HR-style island/section detection for financial databooks.
Focus: 示意性調整後 / Indicative adjusted section, header hierarchy, multiplier, calculation flow.
"""

from .inspector import load_workbook_frames, profile_workbook
from .preflight import _build_workbook_preflight_cached

import pandas as pd
import re
from functools import lru_cache
from typing import Dict, List, Tuple, Optional, Any
from dataclasses import dataclass, field



@dataclass
class TableSection:
    """A detected table section (BS or IS) with metadata."""
    name: str
    start_row: int
    end_row: int
    header_row: int
    marker_row: int
    date_row: int
    desc_col_idx: int
    indicative_cols: List[Tuple[int, str, Optional[Any]]]  # (col_idx, date_str, parsed_date)
    multiply_by_1000: bool
    project_name: Optional[str] = None


@dataclass
class RowClassification:
    """Classification of a data row."""
    is_detail: bool
    is_subtotal: bool
    is_total: bool
    description: str
    values: Dict[str, float]
    row_idx: int


@dataclass
class TableInspection:
    """Full inspection result for a sheet."""
    sheet_name: str
    sections: List[TableSection] = field(default_factory=list)
    row_classifications: List[RowClassification] = field(default_factory=list)
    header_hierarchy: Dict[str, Any] = field(default_factory=dict)
    multiplier_note: str = ""






def _is_subtotal_or_total(desc: str) -> Tuple[bool, bool]:
    """Return (is_subtotal, is_total)."""
    if not desc or pd.isna(desc):
        return False, False
    d = str(desc).strip().lower()
    total_kw = ['总计', '合计', 'total', '小计', 'subtotal', '小計', '合計', '總計']
    is_total = any(k in d for k in ['总计', '合计', 'total', '總計', '合計'])
    is_subtotal = any(k in d for k in ['小计', 'subtotal', 'sub-total', '小計'])
    return is_subtotal, is_total


def _find_header_rows(df: pd.DataFrame) -> List[Tuple[int, str]]:
    """Find all header rows (Indicative adjusted) and their section type. Returns [(row_idx, 'BS'|'IS'|'')]."""
    found = []
    for idx, row in df.iterrows():
        row_str = ' '.join(row.astype(str).values)
        if 'Indicative adjusted' not in row_str and '示意性调整后' not in row_str and '示意性調整後' not in row_str:
            continue
        row_lower = row_str.lower()
        if any(kw.lower() in row_lower for kw in BS_HEADER_KEYWORDS):
            found.append((idx, 'BS'))
        elif any(kw.lower() in row_lower for kw in IS_HEADER_KEYWORDS):
            found.append((idx, 'IS'))
        else:
            found.append((idx, ''))  # Generic indicative row (e.g. first occurrence)
    return sorted(found, key=lambda x: x[0])  # Ensure row order


def inspect_sheet(df: pd.DataFrame, sheet_name: str = "Sheet") -> TableInspection:
    """
    Inspect a financial sheet: detect BS/IS sections, 示意性調整後 columns, header hierarchy, multiplier.
    Uses logic aligned with financial_extraction.py. Handles both BS and IS in same sheet.
    """
    from .statements import _find_description_column, get_valid_financial_columns  # local: table_debug imports the later statements; module-level would be a cycle
    inspection = TableInspection(sheet_name=sheet_name)
    
    if df is None or df.empty:
        return inspection
    
    # ── Find all header rows (may have BS and IS in same sheet) ──
    header_rows = _find_header_rows(df)
    if not header_rows:
        # Fallback: any row with Indicative adjusted
        for idx, row in df.iterrows():
            row_str = ' '.join(row.astype(str).values)
            if 'Indicative adjusted' in row_str or '示意性调整后' in row_str:
                header_rows = [(idx, '')]
                break
    
    if not header_rows:
        inspection.multiplier_note = "No 示意性調整後 / Indicative adjusted header found."
        return inspection
    
    # Use first header row for shared structure (desc col, multiplier)
    header_row_idx = header_rows[0][0]
    
    # ── For each header row, build section ──
    desc_col_idx = _find_description_column(df)
    if desc_col_idx is None:
        inspection.multiplier_note = "No description column found."
        return inspection
    
    all_classifications = []
    header_hierarchies = []
    
    for h_idx, (header_row_idx, section_type) in enumerate(header_rows):
        marker_row_idx = header_row_idx + 2
        date_row_idx = header_row_idx + 1
        
        if marker_row_idx >= len(df) or date_row_idx >= len(df):
            continue
        
        date_row = df.iloc[date_row_idx]
        
        indicative_cols = []
        validated_columns = get_valid_financial_columns(
            df=df,
            desc_col_idx=desc_col_idx,
            header_row_idx=header_row_idx,
        )
        for col_idx, parsed, date_str in validated_columns:
            date_display = parsed.strftime('%Y-%m-%d') if parsed else str(date_str)
            indicative_cols.append((col_idx, date_display, None))
        
        date_row_str = ' '.join(date_row.astype(str).values)
        multiply_by_1000 = "CNY'000" in date_row_str or "人民币千元" in date_row_str
        if not inspection.multiplier_note:
            inspection.multiplier_note = "×1000 (CNY'000 / 人民币千元)" if multiply_by_1000 else "No multiplier"
        
        section_name = "Balance Sheet" if section_type == 'BS' else ("Income Statement" if section_type == 'IS' else "Financial Table")
        data_start = date_row_idx + 1
        data_end = len(df)
        end_keywords = BS_END_KEYWORDS + IS_END_KEYWORDS
        
        # For IS, end at 净利润; for BS, end at 负债及所有者权益总计 or at next section start
        next_section_start = header_rows[h_idx + 1][0] if h_idx + 1 < len(header_rows) else len(df)
        for row_idx in range(data_start, min(next_section_start, len(df))):
            desc = df.iloc[row_idx, desc_col_idx]
            desc_str = str(desc).strip() if pd.notna(desc) else ""
            if any(kw.lower() in desc_str.lower() for kw in end_keywords):
                data_end = row_idx + 1
                break
        
        section = TableSection(
            name=section_name,
            start_row=int(header_row_idx),
            end_row=int(data_end),
            header_row=int(header_row_idx),
            marker_row=int(marker_row_idx),
            date_row=int(date_row_idx),
            desc_col_idx=int(desc_col_idx),
            indicative_cols=indicative_cols,
            multiply_by_1000=multiply_by_1000,
        )
        inspection.sections.append(section)
        
        header_hierarchies.append({
            "section": section_name,
            "header_row": int(header_row_idx),
            "date_row": int(date_row_idx),
            "marker_row": int(marker_row_idx),
            "indicative_columns": [{"col_idx": c[0], "date": c[1]} for c in indicative_cols],
        })
        
        for row_idx in range(data_start, data_end):
            if row_idx >= len(df):
                break
            row = df.iloc[row_idx]
            desc = row.iloc[desc_col_idx]
            desc_str = str(desc).strip() if pd.notna(desc) else ""
            
            values = {}
            for col_idx, date_disp, _ in indicative_cols:
                val = row.iloc[col_idx]
                try:
                    v = float(val)
                    if multiply_by_1000:
                        v *= 1000
                    values[date_disp] = round(v, 0)
                except (ValueError, TypeError):
                    values[date_disp] = 0.0
            
            is_subtotal, is_total = _is_subtotal_or_total(desc_str)
            all_classifications.append(RowClassification(
                is_detail=not (is_subtotal or is_total),
                is_subtotal=is_subtotal,
                is_total=is_total,
                description=desc_str,
                values=values,
                row_idx=int(row_idx),
            ))
    
    inspection.header_hierarchy = {"sections": header_hierarchies} if header_hierarchies else {}
    if header_hierarchies:
        inspection.header_hierarchy["description_col"] = desc_col_idx
        inspection.header_hierarchy["multiply_by_1000"] = inspection.sections[0].multiply_by_1000 if inspection.sections else False
    inspection.row_classifications = all_classifications
    
    return inspection


def inspect_workbook(workbook_path: str, sheet_name: Optional[str] = None) -> Dict[str, TableInspection]:
    """
    Inspect all relevant sheets in a workbook.
    Returns dict of sheet_name -> TableInspection.
    """
    try:
        xls = pd.ExcelFile(workbook_path, engine='openpyxl')
        sheets = [sheet_name] if sheet_name else xls.sheet_names
        results = {}
        for sh in sheets:
            if sh not in xls.sheet_names:
                continue
            df = pd.read_excel(workbook_path, sheet_name=sh, header=None)
            results[sh] = inspect_sheet(df, sh)
        return results
    except Exception as e:
        return {"_error": TableInspection(sheet_name="_error")}  # Placeholder; multiplier_note can hold error




@lru_cache(maxsize=32)
def get_table_inspection(workbook_path: str, sheet_name: str) -> Any:
    inspections = inspect_workbook(workbook_path, sheet_name)
    return inspections.get(sheet_name)


def clear_table_inspection_cache():
    get_table_inspection.cache_clear()


def clear_workbook_caches():
    """Clear all lru_cache entries for workbook profiling and loading."""
    load_workbook_frames.cache_clear()
    profile_workbook.cache_clear()
    _build_workbook_preflight_cached.cache_clear()
    get_table_inspection.cache_clear()
# --- end workbook/table_debug.py ---
