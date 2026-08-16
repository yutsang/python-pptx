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
Standalone Financial Extraction Helper Module.

Extracts Balance Sheet and Income Statement data from Excel workbooks.
"""

from .mapping import find_mapping_key
from .inspector import _cell_text, _coerce_numeric

import pandas as pd
import re
import logging
from datetime import datetime
from typing import Any, Dict, Tuple, Optional, List
import warnings

from ..financial_common import cell_text, coerce_numeric, normalize_financial_date_label

warnings.simplefilter(action='ignore', category=UserWarning)

logger = logging.getLogger(__name__)


def _contains_indicative_marker(text: str) -> bool:
    """Check whether text marks an indicative-adjusted value column."""
    lowered = text.lower()
    # Handle both "Indicative adjusted" (with space) and "Indicativeadjusted" (no space)
    normalised = lowered.replace(' ', '')
    return '示意性调整后' in lowered or '示意性調整後' in lowered or 'indicativeadjusted' in normalised


def _looks_like_remark_text(text: str) -> bool:
    """Detect note-like text that should not be treated as a value column."""
    lowered = text.lower()
    return any(keyword in lowered for keyword in REMARK_KEYWORDS)


def _column_numeric_profile(series: pd.Series) -> Dict[str, float]:
    """Measure how numeric a candidate value column looks."""
    non_empty_count = 0
    numeric_count = 0
    text_count = 0

    for value in series:
        text = _cell_text(value)
        if not text:
            continue
        non_empty_count += 1
        if _coerce_numeric(value) is not None:
            numeric_count += 1
        else:
            text_count += 1

    numeric_ratio = (numeric_count / non_empty_count) if non_empty_count else 0.0
    return {
        'non_empty_count': non_empty_count,
        'numeric_count': numeric_count,
        'text_count': text_count,
        'numeric_ratio': numeric_ratio,
    }

def parse_date(date_str, debug=False, allow_bare_number=True):
    """
    Parse date string in various formats including xMxx and Chinese formats.
    Uses the same preprocessing logic as process_databook.py for consistency.

    Args:
        date_str: Date string in various formats
        debug: If True, print debug info
        allow_bare_number: whether a value carrying no date punctuation at all --
            a raw Excel serial, or a bare four-digit number -- may become a date.
            True preserves the long-standing behaviour and is right once you
            already KNOW the cell is a date. Pass False while still deciding
            WHICH row holds the dates: there, a number that happens to look like
            a date is not a weak signal, it is a wrong answer for the whole tab.

            A real client run is what this is for. date_row_index scores a row by
            how many of its cells parse, and 20000..60000 is precisely the 千元
            range of a warehouse entity's balances, so a row of AMOUNTS outscored
            the real date row on 7 of 7 entities. The corrupted headers went
            straight into the deck: 无形资产 净值3,307.2万 + 累计摊销436.8万 =
            3,744.0万 = 37,440 became "2002-07-03"; 货币资金 4,624.4万 = 46,244
            became "2026-08-10"; 长期借款 3,915.1万 = 39,151 became "2007-03-10".
            Four-digit balances took the other branch and became bare years --
            应收账款 181.8万 = 1,818 became "1818-01-01". The model then quoted
            those headers faithfully, so it read as commentary inventing dates
            when nothing had been invented.

    Returns:
        datetime object or None if parsing fails
    """
    if not date_str or pd.isna(date_str):
        return None

    if not allow_bare_number:
        # Anything that is a number, or spells one with no date punctuation,
        # is refused outright. A genuinely date-formatted Excel cell arrives
        # here as a datetime/Timestamp via openpyxl and is unaffected.
        if isinstance(date_str, (int, float)) and not isinstance(date_str, bool):
            return None
        if str(date_str).strip().replace(",", "").replace(".", "").isdigit():
            return None

    if isinstance(date_str, (int, float)) and not isinstance(date_str, bool):
        serial_value = float(date_str)
        if 20000 <= serial_value <= 60000:
            try:
                result = pd.to_datetime(serial_value, unit='D', origin='1899-12-30', errors='coerce')
                if pd.notna(result):
                    if debug:
                        print(f"      [parse_date]   ✅ Excel serial success: {result.strftime('%Y-%m-%d')}")
                    return result.to_pydatetime(warn=False)
            except Exception:
                pass

    original_str = str(date_str).strip()

    if original_str.isdigit():
        try:
            serial_value = float(original_str)
            if 20000 <= serial_value <= 60000:
                result = pd.to_datetime(serial_value, unit='D', origin='1899-12-30', errors='coerce')
                if pd.notna(result):
                    if debug:
                        print(f"      [parse_date]   ✅ Excel serial string success: {result.strftime('%Y-%m-%d')}")
                    return result.to_pydatetime(warn=False)
        except Exception:
            pass
    
    if debug:
        print(f"      [parse_date] Parsing: '{original_str}'")
    
    # Preprocess the date
    preprocessed = normalize_financial_date_label(original_str)
    
    if debug and preprocessed != original_str:
        print(f"      [parse_date]   Preprocessed: '{preprocessed}'")
    
    # Try to convert preprocessed date to datetime using pandas
    try:
        result = pd.to_datetime(preprocessed, errors='coerce')
        if pd.notna(result):
            if debug:
                print(f"      [parse_date]   ✅ Success: {result.strftime('%Y-%m-%d')}")
            return result.to_pydatetime(warn=False)
    except (TypeError, ValueError, OverflowError):
        pass
    
    if debug:
        print(f"      [parse_date]   ❌ Failed to parse")
    
    return None


def _find_description_column(df: pd.DataFrame, debug: bool = False) -> Optional[int]:
    """Find the most likely description column for a financial table."""
    for col_idx in range(len(df.columns)):
        try:
            col_str = df.iloc[:, col_idx].astype(str)
            if col_str.str.contains(r"CNY'000|人民币千元", case=False, na=False, regex=True).any():
                if debug:
                    print(f"[DEBUG] ✅ Description column found by unit marker at index: {col_idx}")
                return col_idx
        except Exception:
            continue

    tick_symbols = ['✓', '√', '✔', '☑', '■', '□', '●', '○', '★', '☆']
    best_idx = None
    best_score = -1
    rows_to_scan = min(12, len(df))

    for col_idx in range(len(df.columns)):
        sample_values = [_cell_text(df.iloc[row_idx, col_idx]) for row_idx in range(rows_to_scan)]
        non_empty_values = [value for value in sample_values if value]
        if not non_empty_values:
            continue

        has_tick_symbols = sum(any(tick in value for tick in tick_symbols) for value in non_empty_values)
        text_like_count = 0
        remark_count = 0
        date_like_count = 0

        for value in non_empty_values:
            if _looks_like_remark_text(value):
                remark_count += 1
                continue
            if parse_date(value):
                date_like_count += 1
                continue
            if _coerce_numeric(value) is not None:
                continue
            if len(value) > 1:
                text_like_count += 1

        score = (text_like_count * 2) - (has_tick_symbols * 2) - remark_count - date_like_count
        if text_like_count > 0 and score > best_score:
            best_idx = col_idx
            best_score = score

    if best_idx is not None:
        if debug:
            print(f"[DEBUG] ✅ Description column found by text heuristic at index: {best_idx}")
        return best_idx

    if len(df.columns) > 0:
        if debug:
            print("[DEBUG] ⚠️ Falling back to first column as description")
        return 0

    return None


def _get_valid_financial_columns_for_rows(
    df: pd.DataFrame,
    desc_col_idx: int,
    date_row_idx: int,
    marker_row_idx: Optional[int] = None,
    data_start_row: Optional[int] = None,
    data_end_row: Optional[int] = None,
    debug: bool = False,
) -> List[Tuple[int, datetime, str]]:
    """Validate financial columns given explicit date and marker rows."""
    if date_row_idx < 0 or date_row_idx >= len(df):
        return []

    if data_end_row is None:
        data_end_row = len(df)

    marker_row = df.iloc[marker_row_idx] if marker_row_idx is not None and 0 <= marker_row_idx < len(df) else None
    if data_start_row is None:
        last_header_row = max(date_row_idx, marker_row_idx if marker_row_idx is not None else date_row_idx)
        data_start_row = min(last_header_row + 1, len(df))

    date_row = df.iloc[date_row_idx]
    validated_columns: List[Tuple[int, datetime, str]] = []

    for col_idx in range(desc_col_idx + 1, len(df.columns)):
        date_text = _cell_text(date_row.iloc[col_idx])
        marker_text = _cell_text(marker_row.iloc[col_idx]) if marker_row is not None else ''
        parsed_date = parse_date(date_text)
        has_marker = _contains_indicative_marker(marker_text)
        header_has_remark = _looks_like_remark_text(date_text) or _looks_like_remark_text(marker_text)
        body_profile = _column_numeric_profile(df.iloc[data_start_row:data_end_row, col_idx])

        is_numeric_enough = (
            body_profile['numeric_count'] >= 1 and
            body_profile['numeric_ratio'] >= 0.5 and
            body_profile['text_count'] <= 2
        )

        is_valid = bool(
            parsed_date and
            not header_has_remark and
            (has_marker or is_numeric_enough or body_profile['non_empty_count'] == 0)
        )

        if debug:
            print(
                "[DEBUG]   Column {col_idx}: date_row={date_row_idx}, marker_row={marker_row_idx}, "
                "date='{date_text}', marker='{marker_text}', parsed_date={parsed_date}, "
                "numeric_count={numeric_count}, text_count={text_count}, numeric_ratio={numeric_ratio:.2f}, "
                "valid={is_valid}".format(
                    col_idx=col_idx,
                    date_row_idx=date_row_idx,
                    marker_row_idx=marker_row_idx,
                    date_text=date_text,
                    marker_text=marker_text,
                    parsed_date=parsed_date.strftime('%Y-%m-%d') if parsed_date else None,
                    numeric_count=body_profile['numeric_count'],
                    text_count=body_profile['text_count'],
                    numeric_ratio=body_profile['numeric_ratio'],
                    is_valid=is_valid,
                )
            )

        if is_valid:
            validated_columns.append((col_idx, parsed_date, date_text))

    return validated_columns


def _find_best_columns_from_header(
    df: pd.DataFrame,
    desc_col_idx: int,
    header_row_idx: int,
    data_end_row: Optional[int] = None,
    debug: bool = False,
) -> Tuple[Optional[int], Optional[int], List[Tuple[int, datetime, str]]]:
    """Try nearby date/marker layouts and return the best match for a header row."""
    best_date_row_idx = None
    best_marker_row_idx = None
    best_columns: List[Tuple[int, datetime, str]] = []
    best_score = (-1, -1, -9999, -1)

    start_date_row = max(0, header_row_idx - 2)
    end_date_row = min(len(df), header_row_idx + 4)
    for date_row_idx in range(start_date_row, end_date_row):
        for marker_row_idx in (header_row_idx, date_row_idx - 1, date_row_idx + 1):
            if marker_row_idx < 0 or marker_row_idx >= len(df):
                continue
            candidate_columns = _get_valid_financial_columns_for_rows(
                df=df,
                desc_col_idx=desc_col_idx,
                date_row_idx=date_row_idx,
                marker_row_idx=marker_row_idx,
                data_end_row=data_end_row,
                debug=debug,
            )
            marker_row = df.iloc[marker_row_idx] if 0 <= marker_row_idx < len(df) else None
            indicative_count = 0
            if marker_row is not None:
                indicative_count = sum(
                    _contains_indicative_marker(_cell_text(marker_row.iloc[col_idx]))
                    for col_idx, _, _ in candidate_columns
                )
            unique_dates = {
                parsed_date.strftime('%Y-%m-%d')
                for _, parsed_date, _ in candidate_columns
            }
            duplicate_count = len(candidate_columns) - len(unique_dates)
            candidate_score = (
                indicative_count,
                len(unique_dates),
                -duplicate_count,
                len(candidate_columns),
            )
            if candidate_score > best_score:
                best_score = candidate_score
                best_columns = candidate_columns
                best_date_row_idx = date_row_idx
                best_marker_row_idx = marker_row_idx

    return best_date_row_idx, best_marker_row_idx, best_columns


def _select_indicative_cluster(
    date_columns: List[Tuple[int, datetime, str]],
    marker_row: Optional[pd.Series],
) -> List[Tuple[int, datetime, str]]:
    """Prefer the contiguous date block anchored by indicative-adjusted markers."""
    if marker_row is None or not date_columns:
        return date_columns

    indicative_columns = [
        column
        for column in date_columns
        if _contains_indicative_marker(_cell_text(marker_row.iloc[column[0]]))
    ]
    if not indicative_columns:
        return date_columns

    min_col = min(column[0] for column in indicative_columns)
    max_col = max(column[0] for column in indicative_columns)
    clustered_columns = [
        column
        for column in date_columns
        if min_col <= column[0] <= max_col
    ]
    return clustered_columns or indicative_columns


def _dedupe_date_columns(
    date_columns: List[Tuple[int, datetime, str]],
) -> List[Tuple[int, datetime, str]]:
    """Keep the first column for each output date to avoid later overwrites."""
    deduped_columns: List[Tuple[int, datetime, str]] = []
    seen_dates = set()
    for column in date_columns:
        date_key = column[1].strftime('%Y-%m-%d')
        if date_key in seen_dates:
            continue
        seen_dates.add(date_key)
        deduped_columns.append(column)
    return deduped_columns


def _find_relaxed_date_columns(
    df: pd.DataFrame,
    desc_col_idx: int,
    debug: bool = False,
) -> Tuple[Optional[int], List[Tuple[int, datetime, str]]]:
    """Fallback for layouts where headers are messy but date columns are still parseable."""
    return _scan_relaxed_date_columns(
        df=df,
        desc_col_idx=desc_col_idx,
        max_scan_rows=8,
        min_numeric_ratio=0.4,
        max_text_count=2,
        debug=debug,
        label="Relaxed date-column fallback",
    )


def _find_extended_relaxed_date_columns(
    df: pd.DataFrame,
    desc_col_idx: int,
    max_scan_rows: int = 20,
    debug: bool = False,
) -> Tuple[Optional[int], List[Tuple[int, datetime, str]]]:
    """Broader fallback for messy IS layouts with the date row farther down."""
    return _scan_relaxed_date_columns(
        df=df,
        desc_col_idx=desc_col_idx,
        max_scan_rows=max_scan_rows,
        min_numeric_ratio=0.3,
        max_text_count=3,
        debug=debug,
        label="Extended relaxed date fallback",
    )


def _scan_relaxed_date_columns(
    df: pd.DataFrame,
    desc_col_idx: int,
    max_scan_rows: int,
    min_numeric_ratio: float,
    max_text_count: int,
    debug: bool,
    label: str,
) -> Tuple[Optional[int], List[Tuple[int, datetime, str]]]:
    """Shared relaxed date-row scan used by the primary and extended fallbacks."""
    best_date_row_idx = None
    best_columns: List[Tuple[int, datetime, str]] = []
    max_rows = min(max_scan_rows, len(df))

    for date_row_idx in range(max_rows):
        candidate_columns: List[Tuple[int, datetime, str]] = []
        for col_idx in range(desc_col_idx + 1, len(df.columns)):
            date_text = _cell_text(df.iloc[date_row_idx, col_idx])
            parsed_date = parse_date(date_text)
            if not parsed_date or _looks_like_remark_text(date_text):
                continue

            body_profile = _column_numeric_profile(df.iloc[date_row_idx + 1:, col_idx])
            if (
                body_profile['numeric_count'] >= 1
                and body_profile['numeric_ratio'] >= min_numeric_ratio
                and body_profile['text_count'] <= max_text_count
            ):
                candidate_columns.append((col_idx, parsed_date, date_text))

        if len(candidate_columns) > len(best_columns):
            best_columns = candidate_columns
            best_date_row_idx = date_row_idx

    if debug and best_columns:
        print(f"[DEBUG] ✅ {label} selected row {best_date_row_idx} with columns {[col for col, _, _ in best_columns]}")

    return best_date_row_idx, best_columns


def _table_end_keywords(table_name: str) -> List[str]:
    return TABLE_END_KEYWORDS.get(table_name, TABLE_END_KEYWORDS["Income Statement"])


def _find_table_end_row(df: pd.DataFrame, desc_col_idx: int, data_start_row: int, table_name: str) -> int:
    """Locate the logical end of a BS/IS table."""
    data_end_row = len(df)
    end_keywords = _table_end_keywords(table_name)

    for row_idx in range(data_start_row, len(df)):
        desc = _cell_text(df.iloc[row_idx, desc_col_idx])
        if any(keyword.lower() in desc.lower() for keyword in end_keywords):
            return row_idx + 1

    for row_idx in range(data_start_row, len(df)):
        if df.iloc[row_idx].isna().all():
            return row_idx

    return data_end_row


def _build_financial_result(
    df: pd.DataFrame,
    table_name: str,
    desc_col_idx: int,
    date_row_idx: int,
    date_columns: List[Tuple[int, datetime, str]],
    multiply_values: bool = True,
    debug: bool = False,
) -> Optional[pd.DataFrame]:
    """Build output DataFrame once description/date columns are known."""
    if date_row_idx >= len(df) or not date_columns:
        return None

    date_row = df.iloc[date_row_idx]
    header_texts = [' '.join(date_row.astype(str).values)]
    if date_row_idx - 1 >= 0:
        header_texts.append(' '.join(df.iloc[date_row_idx - 1].astype(str).values))
    if date_row_idx + 1 < len(df):
        header_texts.append(' '.join(df.iloc[date_row_idx + 1].astype(str).values))
    header_blob = ' '.join(header_texts)
    multiply_by_1000 = multiply_values and contains_thousand_unit_marker(header_blob)

    data_start_row = date_row_idx + 1
    if data_start_row < len(df):
        row_text = ' '.join(_cell_text(value).lower() for value in df.iloc[data_start_row].values)
        if _contains_indicative_marker(row_text):
            data_start_row += 1

    data_end_row = _find_table_end_row(df, desc_col_idx, data_start_row, table_name)

    result_rows = []
    for row_idx in range(data_start_row, data_end_row):
        description = _cell_text(df.iloc[row_idx, desc_col_idx])
        if not description or _contains_indicative_marker(description):
            continue

        row_dict = {'Description': description}
        for col_idx, parsed_date, _ in date_columns:
            value = df.iloc[row_idx, col_idx]
            numeric_value = _coerce_numeric(value)
            if numeric_value is None:
                numeric_value = 0
            if multiply_by_1000:
                numeric_value *= 1000
            row_dict[parsed_date.strftime('%Y-%m-%d')] = int(round(numeric_value, 0))

        result_rows.append(row_dict)

    if not result_rows:
        if debug:
            print("[DEBUG] ❌ No rows built in financial result helper")
        return None

    result_df = pd.DataFrame(result_rows)
    date_cols = [col for col in result_df.columns if col != 'Description']
    if date_cols:
        result_df = result_df[result_df[date_cols].ne(0).any(axis=1)]

    if result_df.empty:
        if debug:
            print("[DEBUG] ❌ Financial result helper produced only zero rows")
        return None

    return result_df.reset_index(drop=True)


def _extract_income_statement_directly(
    df: pd.DataFrame,
    debug: bool = False,
    multiply_values: bool = True,
) -> Optional[pd.DataFrame]:
    """Fallback IS extractor for layouts the primary indicative-header logic misses."""
    working_df = df.dropna(how='all').dropna(axis=1, how='all').reset_index(drop=True)
    if working_df.empty:
        return None

    if debug:
        print(f"[DEBUG] 🔄 Direct IS fallback on shape {working_df.shape}")

    desc_col_idx = _find_description_column(working_df, debug=debug)
    if desc_col_idx is None:
        return None

    date_row_idx, date_columns = _find_extended_relaxed_date_columns(
        working_df,
        desc_col_idx,
        debug=debug,
    )
    if not date_columns or date_row_idx is None:
        return None

    return _build_financial_result(
        df=working_df,
        table_name="Income Statement",
        desc_col_idx=desc_col_idx,
        date_row_idx=date_row_idx,
        date_columns=date_columns,
        multiply_values=multiply_values,
        debug=debug,
    )


def get_valid_financial_columns(
    df: pd.DataFrame,
    desc_col_idx: int,
    header_row_idx: int,
    data_start_row: Optional[int] = None,
    data_end_row: Optional[int] = None,
    debug: bool = False,
) -> List[Tuple[int, datetime, str]]:
    """
    Return validated financial value columns to the right of the description column.

    A valid financial column should look like a reporting-date column and should not
    behave like a free-text remark column.
    """
    if data_start_row is not None:
        date_row_idx = max(header_row_idx + 1, data_start_row - 1)
        marker_row_idx = date_row_idx + 1 if date_row_idx + 1 < len(df) else None
        return _get_valid_financial_columns_for_rows(
            df=df,
            desc_col_idx=desc_col_idx,
            date_row_idx=date_row_idx,
            marker_row_idx=marker_row_idx,
            data_start_row=data_start_row,
            data_end_row=data_end_row,
            debug=debug,
        )

    _, _, validated_columns = _find_best_columns_from_header(
        df=df,
        desc_col_idx=desc_col_idx,
        header_row_idx=header_row_idx,
        data_end_row=data_end_row,
        debug=debug,
    )
    return validated_columns






def extract_financial_table(
    df: pd.DataFrame,
    table_name: str,
    entity_keywords: Optional[List[str]] = None,
    debug: bool = False,
    multiply_values: bool = True
) -> Optional[pd.DataFrame]:
    """
    Extract financial table (Balance Sheet or Income Statement) from a worksheet.
    Gets ALL columns with "示意性调整后" or "Indicative adjusted".
    
    Args:
        df: DataFrame containing the financial data
        table_name: Name of the table (e.g., "Balance Sheet", "Income Statement")
        entity_keywords: Optional list of entity name components to search for
        debug: If True, print debugging information
        multiply_values: If True, multiply by 1000 if CNY'000 detected
        
    Returns:
        Cleaned DataFrame with Description column and ALL adjusted columns
    """
    if debug:
        print(f"\n[DEBUG] Extracting {table_name}...")
        print(f"[DEBUG] DataFrame shape: {df.shape}")
    
    # Detect header row candidates with "Indicative adjusted" or "示意性调整后"
    header_row_candidates = []
    for idx, row in df.iterrows():
        row_str = ' '.join(row.astype(str).values)
        if _contains_indicative_marker(row_str):
            header_row_candidates.append(idx)

    if debug:
        if header_row_candidates:
            print(f"[DEBUG] ✅ Header row candidates found: {header_row_candidates}")
        else:
            print(f"[DEBUG] ⚠️ No indicative-adjusted header row candidates found")

    # Find description column
    if debug:
        print(f"[DEBUG] Searching for description column...")

    desc_col_idx = _find_description_column(df, debug=debug)

    if desc_col_idx is None:
        if debug:
            print(f"[DEBUG] ❌ No description column found")
        return None
    
    if debug:
        print(f"[DEBUG] Showing first 5 rows of ALL columns (to understand structure):")
        for row_num in range(min(5, len(df))):
            print(f"[DEBUG] Row {row_num}: {df.iloc[row_num].values[:20]}")  # Show first 20 cols
    
    if debug:
        print(f"\n[DEBUG] ========== FILTERING FINANCIAL VALUE COLUMNS ==========")

    best_header_row_idx = None
    best_date_row_idx = None
    best_marker_row_idx = None
    date_columns: List[Tuple[int, datetime, str]] = []

    for candidate_header_row in header_row_candidates:
        candidate_date_row_idx, candidate_marker_row_idx, candidate_columns = _find_best_columns_from_header(
            df=df,
            desc_col_idx=desc_col_idx,
            header_row_idx=candidate_header_row,
            debug=debug,
        )
        if len(candidate_columns) > len(date_columns):
            best_header_row_idx = candidate_header_row
            best_date_row_idx = candidate_date_row_idx
            best_marker_row_idx = candidate_marker_row_idx
            date_columns = candidate_columns

    if not date_columns and table_name == "Income Statement":
        best_date_row_idx, date_columns = _find_relaxed_date_columns(df, desc_col_idx, debug=debug)
        best_header_row_idx = best_date_row_idx - 1 if best_date_row_idx is not None and best_date_row_idx > 0 else best_date_row_idx
        best_marker_row_idx = None

    if not date_columns:
        if debug:
            print(f"[DEBUG] ❌ No columns found!")
        return None

    marker_row = None
    if best_marker_row_idx is not None and 0 <= best_marker_row_idx < len(df):
        marker_row = df.iloc[best_marker_row_idx]
        date_columns = _select_indicative_cluster(date_columns, marker_row)

    date_columns = _dedupe_date_columns(date_columns)

    header_row_idx = best_header_row_idx if best_header_row_idx is not None else 0
    date_row_idx = best_date_row_idx if best_date_row_idx is not None else header_row_idx + 1

    if debug:
        print(f"\n[DEBUG] ========== FINAL SELECTED COLUMNS ==========")
        print(f"[DEBUG] Header row index: {header_row_idx}")
        print(f"[DEBUG] Date row index: {date_row_idx}")
        print(f"[DEBUG] Marker row index: {best_marker_row_idx}")
        print(f"[DEBUG] Total columns selected: {len(date_columns)}")
        print(f"[DEBUG] Column indices: {[col_idx for col_idx, _, _ in date_columns]}")
        print(f"=" * 80)

    # Get date row
    if date_row_idx >= len(df):
        if debug:
            print(f"[DEBUG] ❌ Date row index {date_row_idx} is out of bounds")
        return None
    
    date_row = df.iloc[date_row_idx]
    
    if debug:
        print(f"\n[DEBUG] ========== COLUMNS SELECTED FOR OUTPUT ==========")
        print(f"[DEBUG] Description column: {desc_col_idx}")
        print(f"[DEBUG] Date columns found: {len(date_columns)}")
        for col_idx, parsed_date, date_str in date_columns:
            print(f"[DEBUG]   ✅ Column {col_idx}: '{date_str}' → will be named '{parsed_date.strftime('%Y-%m-%d')}'")
        
        print(f"\n[DEBUG] Final output will have columns:")
        output_cols = ['Description'] + [parsed_date.strftime('%Y-%m-%d') for _, parsed_date, _ in date_columns]
        print(f"[DEBUG]   {output_cols}")
        print(f"[DEBUG] Total: {len(output_cols)} columns ({len(output_cols)-1} date columns)")
        print(f"=" * 80)
    
    # Check if CNY'000 multiplier needed
    header_texts = [' '.join(date_row.astype(str).values)]
    if date_row_idx - 1 >= 0:
        header_texts.append(' '.join(df.iloc[date_row_idx - 1].astype(str).values))
    if date_row_idx + 1 < len(df):
        header_texts.append(' '.join(df.iloc[date_row_idx + 1].astype(str).values))
    header_blob = ' '.join(header_texts)
    multiply_by_1000 = multiply_values and contains_thousand_unit_marker(header_blob)

    if debug and multiply_by_1000:
        print(f"[DEBUG] Will multiply values by 1000 (CNY'000 detected)")
    elif debug and not multiply_values:
        print(f"[DEBUG] Multiplication disabled by parameter")
    
    # Determine end row based on table type
    data_start_row = date_row_idx + 1
    data_end_row = len(df)
    
    # For Balance Sheet: end at liabilities/equity total.
    # For Income Statement: end at net profit/(loss)-style rows.
    end_keywords = _table_end_keywords(table_name)
    
    if debug:
        print(f"[DEBUG] Looking for end markers: {end_keywords}")
        print(f"[DEBUG] Searching from row {data_start_row} to {len(df)}")
    
    end_marker_found = False
    for row_idx in range(data_start_row, len(df)):
        row = df.iloc[row_idx]
        desc = str(row.iloc[desc_col_idx]).strip()
        
        if debug and row_idx < data_start_row + 10:  # Show first 10 rows
            print(f"[DEBUG]   Row {row_idx}: '{desc}'")
        
        if any(keyword.lower() in desc.lower() for keyword in end_keywords):
            data_end_row = row_idx + 1  # Include this row
            end_marker_found = True
            if debug:
                print(f"[DEBUG] ✅ Found end marker at row {row_idx}: '{desc}'")
            break
    
    if debug:
        if not end_marker_found:
            print(f"[DEBUG] ⚠️  No end marker found! Will extract to end of dataframe (row {len(df)})")
        print(f"[DEBUG] Data extraction range: rows {data_start_row} to {data_end_row} ({data_end_row - data_start_row} rows)")
        print(f"[DEBUG] Preview of extraction range:")
        for row_idx in range(data_start_row, min(data_start_row + 5, data_end_row)):
            if row_idx < len(df):
                desc = str(df.iloc[row_idx].iloc[desc_col_idx]).strip()
                print(f"[DEBUG]   Row {row_idx}: '{desc}'")
    
    # Build result dataframe with Description + ALL adjusted columns
    if debug:
        print(f"\n[DEBUG] Extracting data from rows {data_start_row} to {data_end_row}...")
        print(f"[DEBUG] Will extract from description column {desc_col_idx} and date columns: {[col for col, _, _ in date_columns]}")
    
    result_rows = []
    skipped_empty_desc = 0
    
    for row_idx in range(data_start_row, data_end_row):
        row = df.iloc[row_idx]
        
        description = row.iloc[desc_col_idx]
        
        # Skip if description is null or empty
        if pd.isna(description) or str(description).strip() == '':
            skipped_empty_desc += 1
            continue
        
        # Build row dict with description and all date values
        row_dict = {'Description': str(description).strip()}
        
        has_any_nonzero_value = False
        conversion_errors = 0
        
        for col_idx, parsed_date, date_str in date_columns:
            value = row.iloc[col_idx]
            col_name = parsed_date.strftime('%Y-%m-%d')
            
            numeric_value = _coerce_numeric(value)
            try:
                if numeric_value is None:
                    raise ValueError("not numeric")
                if multiply_by_1000:
                    numeric_value *= 1000
                numeric_value = round(numeric_value, 0)
                
                row_dict[col_name] = int(numeric_value)
                
                if numeric_value != 0:
                    has_any_nonzero_value = True
                    
            except (ValueError, TypeError) as e:
                conversion_errors += 1
                row_dict[col_name] = 0
        
        # Debug only problematic rows or first few
        if debug and (len(result_rows) < 3 or conversion_errors > 0 or not has_any_nonzero_value):
            values_str = ", ".join([f"{col}: {row_dict[col]}" for col in row_dict if col != 'Description'])
            status = "⚠️ ALL ZEROS" if not has_any_nonzero_value else ("⚠️ ERRORS" if conversion_errors > 0 else "✅")
            print(f"[DEBUG]   Row {row_idx} {status}: '{row_dict['Description'][:50]}' → {values_str}")
            print(f"[DEBUG]     row_dict keys: {list(row_dict.keys())}")
            print(f"[DEBUG]     row_dict values: {list(row_dict.values())}")
        
        # Add row (even if all zeros, we'll filter later)
        result_rows.append(row_dict)
    
    if debug:
        print(f"\n[DEBUG] Extraction complete:")
        print(f"[DEBUG]   - Total rows processed: {data_end_row - data_start_row}")
        print(f"[DEBUG]   - Rows with empty descriptions: {skipped_empty_desc}")
        print(f"[DEBUG]   - Rows extracted: {len(result_rows)}")
        
        # Show which columns had most conversion errors
        if result_rows:
            temp_df = pd.DataFrame(result_rows)
            for col in temp_df.columns:
                if col != 'Description':
                    zero_count = (temp_df[col] == 0).sum()
                    nonzero_count = (temp_df[col] != 0).sum()
                    print(f"[DEBUG]   Column '{col}': {nonzero_count} non-zero, {zero_count} zeros")
    
    if not result_rows:
        if debug:
            print(f"[DEBUG] ❌ No valid data rows found!")
            print(f"[DEBUG] Processed {data_end_row - data_start_row} rows but none had valid data")
        return None
    
    if debug:
        print(f"\n[DEBUG] ========== CREATING DATAFRAME ==========")
        print(f"[DEBUG] Creating DataFrame from {len(result_rows)} rows")
        
        # Show first 3 result_rows as dict
        for i, row_dict in enumerate(result_rows[:3]):
            print(f"[DEBUG] result_rows[{i}]:")
            for k, v in row_dict.items():
                print(f"[DEBUG]   {k}: {v} (type: {type(v).__name__})")
    
    result_df = pd.DataFrame(result_rows)
    
    if debug:
        print(f"\n[DEBUG] DataFrame created successfully!")
        print(f"[DEBUG]   Shape: {result_df.shape}")
        print(f"[DEBUG]   Columns: {list(result_df.columns)}")
        print(f"[DEBUG]   Dtypes: {result_df.dtypes.to_dict()}")
        print(f"\n[DEBUG] DataFrame content (first 3 rows):")
        print(result_df.head(3).to_string())
        
        # Check for any issues with column values
        for col in result_df.columns:
            if col != 'Description':
                print(f"[DEBUG]   Column '{col}' stats: min={result_df[col].min()}, max={result_df[col].max()}, mean={result_df[col].mean():.0f}")
    
    # Remove rows where ALL date column values are 0
    date_cols = [col for col in result_df.columns if col != 'Description']
    
    if debug:
        print(f"\n[DEBUG] ========== FILTERING ZERO ROWS ==========")
        print(f"[DEBUG] Date columns to check: {date_cols}")
    
    if date_cols:
        # Keep rows where at least one date column is non-zero
        rows_before = len(result_df)
        mask = result_df[date_cols].ne(0).any(axis=1)
        
        if debug:
            print(f"[DEBUG] Rows before filtering: {rows_before}")
            print(f"[DEBUG] Mask (True = keep, False = remove):")
            print(f"[DEBUG]   {mask.values[:20]}")  # Show first 20
            
            # Show which rows will be removed
            removed_indices = result_df[~mask].index
            if len(removed_indices) > 0:
                print(f"[DEBUG] Rows to be removed ({len(removed_indices)} total):")
                for idx in list(removed_indices)[:5]:
                    desc = result_df.loc[idx, 'Description']
                    vals = [result_df.loc[idx, col] for col in date_cols]
                    print(f"[DEBUG]   Row {idx}: '{desc}' → {vals}")
        
        result_df = result_df[mask]
        rows_after = len(result_df)
        
        if debug:
            print(f"[DEBUG] Rows after filtering: {rows_after}")
            print(f"[DEBUG] Removed {rows_before - rows_after} rows with all zeros")
    
    if result_df.empty:
        if debug:
            print(f"[DEBUG] ❌ DataFrame is empty after removing zero rows!")
        return None
    
    if debug:
        print(f"\n[DEBUG] ========== FINAL RESULT ==========")
        print(f"[DEBUG] ✅ Final DataFrame: {len(result_df)} rows × {len(result_df.columns)} columns")
        print(f"[DEBUG] Columns: {list(result_df.columns)}")
        
        # Show value statistics for each date column
        for col in result_df.columns:
            if col != 'Description':
                non_zero = (result_df[col] != 0).sum()
                zero = (result_df[col] == 0).sum()
                print(f"[DEBUG]   '{col}': {non_zero} non-zero, {zero} zeros (max: {result_df[col].max():,.0f})")
        
        print(f"\n[DEBUG] First 5 rows:")
        print(result_df.head(5).to_string())
        
        print(f"\n[DEBUG] Last 5 rows:")
        print(result_df.tail(5).to_string())
        
        # Check if there are any rows with all values as 0
        if len(date_cols) > 0:
            all_zero_mask = (result_df[date_cols] == 0).all(axis=1)
            all_zero_count = all_zero_mask.sum()
            if all_zero_count > 0:
                print(f"\n[DEBUG] ⚠️ WARNING: {all_zero_count} rows still have ALL zeros!")
                print(f"[DEBUG] These rows:")
                print(result_df[all_zero_mask].to_string())
    
    return result_df


# Headers that mark the start of a NON-financial-statement section on a
# Financials sheet -- operating KPIs, ratio analysis, per-unit metrics.
# These sit below the P&L on real sheets and must not be read as accounts:
# their values are rates and percentages, not amounts.
_POST_IS_SECTION_MARKERS = (
    "经营指标", "經營指標", "主要经营指标", "主要經營指標", "运营指标", "營運指標",
    "关键指标", "關鍵指標", "财务比率", "財務比率", "比率分析", "指标分析",
    "operating metric", "operating indicator", "key metric", "key indicator",
    "kpi", "ratio analysis", "financial ratio", "key performance",
)
# Ratio/per-unit row labels. Used only as a FALLBACK when no section header
# exists -- several of these appearing consecutively is itself the boundary.
_RATIO_ROW_MARKERS = (
    "出租率", "单位租金", "單位租金", "单位物管费", "單位物管費",
    "毛利率", "净利率", "淨利率", "营业利润率", "營業利潤率", "增长率", "增長率",
    "ebitda%", "占收入比", "每平方米", "元/平方米",
)


def _find_section_end_row(df: pd.DataFrame, start_row: int) -> int:
    """Row index at which the section beginning at start_row ends.

    Prefers an explicit section header (see _POST_IS_SECTION_MARKERS). Where
    a sheet has none, falls back to the first run of consecutive ratio/
    per-unit rows -- one such label can legitimately appear inside a P&L
    (e.g. a margin shown as a memo line), so a single hit is not treated as
    the boundary, but a run of them is a different section.
    """
    n = len(df)
    run_start = None
    consecutive = 0
    for idx in range(start_row + 1, n):
        row_str = " ".join(df.iloc[idx].astype(str).values).lower()
        if any(marker in row_str for marker in _POST_IS_SECTION_MARKERS):
            return idx
        if any(marker in row_str for marker in _RATIO_ROW_MARKERS):
            if consecutive == 0:
                run_start = idx
            consecutive += 1
            if consecutive >= 3:
                return run_start
        elif row_str.strip() and row_str.replace("nan", "").strip():
            consecutive = 0
            run_start = None
    return n


def extract_balance_sheet_and_income_statement(
    workbook_path: str,
    sheet_name: str,
    debug: bool = False,
    multiply_values: bool = True
) -> Dict[str, Any]:
    """
    Extract Balance Sheet and Income Statement from a SINGLE Excel worksheet.
    Both BS and IS are in the same sheet, separated by header rows.
    
    Args:
        workbook_path: Path to Excel workbook
        sheet_name: Worksheet name containing both BS and IS
        debug: If True, print debugging information
        multiply_values: If True, multiply by 1000 if CNY'000 detected
        
    Returns:
        Dictionary with keys:
        - 'balance_sheet': DataFrame or None
        - 'income_statement': DataFrame or None  
        - 'project_name': String (extracted from headers) or None
        
    Example:
        >>> results = extract_balance_sheet_and_income_statement(
        ...     workbook_path="databook.xlsx",
        ...     sheet_name="Financial Statements",
        ...     debug=True
        ... )
        >>> print(results['balance_sheet'])
        >>> print(results['income_statement'])
        >>> print(results['project_name'])
    """
    results = {
        'balance_sheet': None,
        'income_statement': None,
        'project_name': None
    }
    
    if debug:
        print("=" * 80)
        print("FINANCIAL EXTRACTION - DEBUG MODE")
        print("=" * 80)
        print(f"Workbook: {workbook_path}")
        print(f"Sheet: {sheet_name}")
    
    try:
        # Load Excel file
        df = pd.read_excel(workbook_path, sheet_name=sheet_name, header=None, engine='openpyxl')
        
        if debug:
            print(f"\n[DEBUG] ✅ Sheet loaded: {df.shape}")
        
        # Find Balance Sheet section
        bs_start_row = None
        bs_keywords = [
            "示意性调整后资产负债表",
            "示意性調整後資產負債表",
            "Indicative adjusted balance sheet",
            "Indicative Adjusted Balance Sheet",
            "Balance sheet",
        ]
        
        for idx, row in df.iterrows():
            row_str = ' '.join(row.astype(str).values).lower()
            if any(kw.lower() in row_str for kw in bs_keywords):
                bs_start_row = idx
                if debug:
                    print(f"[DEBUG] ✅ Balance Sheet starts at row {idx}: {df.iloc[idx].values[0]}")
                break
        
        # Find Income Statement section  
        is_start_row = None
        is_keywords = [
            "示意性调整后利润表",
            "示意性調整後利潤表",
            "Indicative adjusted income statement",
            "Indicative Adjusted Income Statement",
            "Income statement",
            "profit and loss",
            "statement of comprehensive income",
        ]
        
        for idx, row in df.iterrows():
            row_str = ' '.join(row.astype(str).values).lower()
            if any(kw.lower() in row_str for kw in is_keywords):
                is_start_row = idx
                if debug:
                    print(f"[DEBUG] ✅ Income Statement starts at row {idx}: {df.iloc[idx].values[0]}")
                break

        if bs_start_row is None:
            # Mirror the relaxed IS fallback below: lowercase the row so the English
            # keyword matches any case ("BALANCE SHEET" / "Balance Sheet").
            relaxed_bs_keywords = ["资产负债表", "資產負債表", "balance sheet"]
            for idx, row in df.iterrows():
                row_str = ' '.join(row.astype(str).values).lower()
                if any(keyword in row_str for keyword in relaxed_bs_keywords):
                    bs_start_row = idx
                    if debug:
                        print(f"[DEBUG] ✅ Balance Sheet starts at row {idx} using relaxed detection")
                    break

        if is_start_row is None:
            relaxed_is_keywords = ["利润表", "利潤表", "income statement", "profit and loss"]
            for idx, row in df.iterrows():
                row_str = ' '.join(row.astype(str).values).lower()
                if any(keyword in row_str for keyword in relaxed_is_keywords):
                    is_start_row = idx
                    if debug:
                        print(f"[DEBUG] ✅ Income Statement starts at row {idx} using relaxed detection")
                    break
        
        # Extract project name (from header row pattern)
        # Pattern: "示意性调整后资产负债表 - <Project Name>" or "Balance Sheet - Project Name"
        # Should appear in both BS and IS headers
        project_name_bs = None
        project_name_is = None
        
        if bs_start_row is not None:
            # Check all cells in BS header row for the pattern
            bs_row = df.iloc[bs_start_row]
            for val in bs_row:
                val_str = str(val)
                if _contains_indicative_marker(val_str) or 'balance sheet' in val_str.lower():
                    if ' - ' in val_str:
                        project_name_bs = val_str.split(' - ', 1)[1].strip()
                    elif '-' in val_str and '调整后' not in val_str.split('-')[-1]:
                        project_name_bs = val_str.split('-')[-1].strip()
                    break
            
            if debug:
                print(f"[DEBUG] BS header project name: '{project_name_bs}'")
        
        if is_start_row is not None:
            # Check all cells in IS header row for the pattern
            is_row = df.iloc[is_start_row]
            for val in is_row:
                val_str = str(val)
                if _contains_indicative_marker(val_str) or 'income statement' in val_str.lower():
                    if ' - ' in val_str:
                        project_name_is = val_str.split(' - ', 1)[1].strip()
                    elif '-' in val_str and '调整后' not in val_str.split('-')[-1]:
                        project_name_is = val_str.split('-')[-1].strip()
                    break
            
            if debug:
                print(f"[DEBUG] IS header project name: '{project_name_is}'")
        
        # Use project name if it appears in both headers (or if only one is found)
        if project_name_bs and project_name_is:
            if project_name_bs == project_name_is:
                project_name = project_name_bs
                if debug:
                    print(f"[DEBUG] ✅ Project name confirmed in both headers: '{project_name}'")
            else:
                if debug:
                    print(f"[DEBUG] ⚠️  Project names don't match! BS: '{project_name_bs}', IS: '{project_name_is}'")
                project_name = project_name_bs  # Use BS name as default
        elif project_name_bs:
            project_name = project_name_bs
        elif project_name_is:
            project_name = project_name_is
        else:
            project_name = None
            if debug:
                print(f"[DEBUG] ❌ No project name found in headers")
        
        results['project_name'] = project_name
        
        # Extract Balance Sheet
        if bs_start_row is not None:
            # Determine end row (either IS start or end of sheet)
            bs_end_row = is_start_row if is_start_row else len(df)
            df_bs = df.iloc[bs_start_row:bs_end_row].copy().reset_index(drop=True)
            
            results['balance_sheet'] = extract_financial_table(
                df_bs, "Balance Sheet", None, debug, multiply_values
            )
        
        # Extract Income Statement
        if is_start_row is not None:
            # The IS used to run to the end of the sheet, unlike the BS which
            # is bounded by is_start_row. On a real Financials sheet that
            # carries an operating-KPI block below the P&L, every one of those
            # rows was swallowed into the income statement -- a reconciliation
            # page came back listing EBITDA%, 出租率, 单位租金 and 毛利率 as if
            # they were accounts, with ratio values (587, 866, -347) read as
            # amounts, and none of the actual P&L lines. Bound the IS at the
            # next section header the same way the BS is bounded.
            is_end_row = _find_section_end_row(df, is_start_row)
            df_is = df.iloc[is_start_row:is_end_row].copy().reset_index(drop=True)
            if debug and is_end_row < len(df):
                print(f"[DEBUG] Income Statement bounded at row {is_end_row}: "
                      f"{str(df.iloc[is_end_row].values[0])[:60]!r}")
            
            results['income_statement'] = extract_financial_table(
                df_is, "Income Statement", None, debug, multiply_values
            )

            if results['income_statement'] is None:
                if debug:
                    print("[DEBUG] ⚠️ Primary Income Statement extraction returned None; trying direct fallback")
                results['income_statement'] = _extract_income_statement_directly(
                    df_is,
                    debug=debug,
                    multiply_values=multiply_values,
                )
                if debug:
                    status = "✅ succeeded" if results['income_statement'] is not None else "❌ failed"
                    print(f"[DEBUG] Direct Income Statement fallback {status}")
        
        # Post-processing: Remove date columns with all zeros in Income Statement
        if results['income_statement'] is not None:
            is_df = results['income_statement']
            date_cols = [col for col in is_df.columns if col != 'Description']
            
            # Find columns with all zeros in IS
            cols_to_remove = []
            for col in date_cols:
                if (is_df[col] == 0).all():
                    cols_to_remove.append(col)
            
            if cols_to_remove:
                remaining_date_cols = [col for col in date_cols if col not in cols_to_remove]
                if not remaining_date_cols:
                    if debug:
                        print("[DEBUG] ⚠️ Skipping zero-column removal because it would remove all Income Statement date columns")
                    cols_to_remove = []

            if cols_to_remove:
                if debug:
                    print(f"\n[DEBUG] ========== REMOVING ZERO COLUMNS ==========")
                    print(f"[DEBUG] Found {len(cols_to_remove)} date columns with ALL zeros in Income Statement:")
                    print(f"[DEBUG]   {cols_to_remove}")
                    print(f"[DEBUG] Removing these columns from BOTH Balance Sheet and Income Statement...")
                
                # Remove from Income Statement
                results['income_statement'] = is_df.drop(columns=cols_to_remove)
                
                # Remove from Balance Sheet ONLY where the column is ALSO all-zero in
                # the BS. A period can have a zero Income Statement but real balances
                # (e.g. no P&L activity yet existing assets/liabilities); dropping such
                # a BS column silently loses data and shifts the reconciliation date.
                if results['balance_sheet'] is not None:
                    bs_df = results['balance_sheet']
                    cols_to_remove_from_bs = [
                        col for col in cols_to_remove
                        if col in bs_df.columns and (bs_df[col] == 0).all()
                    ]
                    if cols_to_remove_from_bs:
                        results['balance_sheet'] = bs_df.drop(columns=cols_to_remove_from_bs)
                        if debug:
                            print(f"[DEBUG]   Removed {len(cols_to_remove_from_bs)} all-zero columns from Balance Sheet")
                    elif debug:
                        print("[DEBUG]   Kept Balance Sheet columns (non-zero in BS even though IS was zero)")
                
                if debug:
                    print(f"[DEBUG] ✅ Columns removed successfully")
        
        if debug:
            print("\n" + "=" * 80)
            print("EXTRACTION RESULTS:")
            print("=" * 80)
            print(f"Project Name: {results['project_name'] or '❌ Not found'}")
            print(f"Balance Sheet: {'✅ Extracted' if results['balance_sheet'] is not None else '❌ None'}")
            print(f"Income Statement: {'✅ Extracted' if results['income_statement'] is not None else '❌ None'}")
            if results['balance_sheet'] is not None:
                print(f"  - Balance Sheet: {len(results['balance_sheet'])} rows × {len(results['balance_sheet'].columns)} cols")
                print(f"  - Columns: {list(results['balance_sheet'].columns)}")
            if results['income_statement'] is not None:
                print(f"  - Income Statement: {len(results['income_statement'])} rows × {len(results['income_statement'].columns)} cols")
                print(f"  - Columns: {list(results['income_statement'].columns)}")
        
    except Exception as e:
        logger.error("Error extracting financial data: %s", e)
        if debug:
            import traceback
            logger.debug("Full traceback for financial extraction error:", exc_info=True)
    
    return results


_SYNTHETIC_BS_CATEGORY_ORDER = [
    "Current assets", "Non-current assets", "Current liabilities", "Non-current liabilities", "Equity",
]
_SYNTHETIC_BS_GRAND_TOTAL_GROUPS = [
    ("Total assets", ["Current assets", "Non-current assets"]),
    ("Total liabilities", ["Current liabilities", "Non-current liabilities"]),
    ("Total owners' equity", ["Equity"]),
]
_SYNTHETIC_IS_CATEGORY_ORDER = ["Revenue", "Expenses"]


def _synthetic_account_total_row(df: pd.DataFrame) -> Optional[Dict[str, Optional[float]]]:
    """Pulls a single schedule tab's own Total/合计 row values out, keyed by
    its date-string columns -- the building block for
    synthesize_balance_sheet_and_income_statement."""
    if df is None or df.empty:
        return None
    desc_col = df.columns[0]
    date_cols = [
        c for c in df.columns
        if c != desc_col and not str(c).endswith("_formatted") and not str(c).startswith("__")
    ]
    if not date_cols:
        return None
    row_types = df.attrs.get("row_types_by_description") or {}
    total_idx = None
    for idx, desc in df[desc_col].items():
        if row_types.get(str(desc)) == "total":
            total_idx = idx
    if total_idx is None:
        total_idx = df.index[-1]  # convention: the Total/合计 row is always last
    row = df.loc[total_idx]
    values: Dict[str, Optional[float]] = {}
    for col in date_cols:
        val = row[col]
        values[col] = None if val is None or (isinstance(val, float) and pd.isna(val)) else float(val)
    return values


def _synthesize_statement(
    dfs: Dict[str, pd.DataFrame],
    mappings: Dict[str, Any],
    statement_type: str,
    category_order: List[str],
    grand_total_groups: List[Tuple[str, List[str]]],
) -> Optional[pd.DataFrame]:
    # dfs is keyed by each account's resolved DISPLAY name (e.g. "COGS",
    # "Paid-in capital"), which is NOT always the same string as
    # mappings.yml's own top-level key (e.g. "OC", "Capital") -- resolve via
    # find_mapping_key (alias lookup), same as split_accounts_by_type does,
    # instead of checking `mapping_key in dfs` directly (which silently
    # dropped almost every account with a non-identical display name).
    dfs_key_by_mapping_key: Dict[str, str] = {}
    for dfs_key in dfs:
        mapping_key = find_mapping_key(dfs_key, mappings)
        if mapping_key:
            dfs_key_by_mapping_key[mapping_key] = dfs_key

    per_account_values: List[Tuple[str, Dict[str, Any], Dict[str, Optional[float]]]] = []
    for mapping_key, meta in mappings.items():
        if not isinstance(meta, dict) or meta.get("type") != statement_type:
            continue
        dfs_key = dfs_key_by_mapping_key.get(mapping_key)
        if not dfs_key:
            continue
        values = _synthetic_account_total_row(dfs[dfs_key])
        if values:
            per_account_values.append((dfs_key, meta, values))
    if not per_account_values:
        return None

    # A Balance Sheet/Income Statement is a snapshot as of ONE date -- but
    # each schedule tab's own "latest available" column can legitimately
    # differ (e.g. one tab genuinely hasn't been updated past an earlier
    # period). Pick whichever date the MOST accounts actually have data for
    # ("as of" date the databook is mostly current to) and only use that
    # column, rather than unioning every column across every account (which
    # would silently sum mismatched-period figures into a meaningless total).
    date_vote: Dict[str, int] = {}
    for _, _, values in per_account_values:
        for col, val in values.items():
            if val is not None:
                date_vote[col] = date_vote.get(col, 0) + 1
    if not date_vote:
        return None
    target_date = max(date_vote.items(), key=lambda kv: (kv[1], parse_date(kv[0]) or kv[0]))[0]

    category_rows: Dict[str, List[Dict[str, Any]]] = {cat: [] for cat in category_order}
    other_rows: List[Dict[str, Any]] = []
    for dfs_key, meta, values in per_account_values:
        if target_date not in values or values[target_date] is None:
            continue
        row = {"Description": dfs_key, target_date: values[target_date]}
        if meta.get("category") in category_rows:
            category_rows[meta["category"]].append(row)
        else:
            other_rows.append(row)

    def _zero_row(label: str) -> Dict[str, Any]:
        return {"Description": label, target_date: 0.0}

    def _sum_rows(label: str, group_rows: List[Dict[str, Any]]) -> Dict[str, Any]:
        total = _zero_row(label)
        for r in group_rows:
            total[target_date] += r.get(target_date) or 0.0
        return total

    rows: List[Dict[str, Any]] = []
    category_subtotals: Dict[str, Dict[str, Any]] = {}
    sole_member_categories = {
        cats[0] for _, cats in grand_total_groups if len(cats) == 1
    }
    for category in category_order:
        group = category_rows.get(category) or []
        if not group:
            continue
        rows.extend(group)
        label = f"Total {category[0].lower()}{category[1:]}"
        subtotal = _sum_rows(label, group)
        category_subtotals[category] = subtotal
        # Skip the redundant per-category subtotal when this category is the
        # SOLE member of a grand-total group below (e.g. Equity) -- the grand
        # total row would just duplicate it under a different label.
        if category not in sole_member_categories:
            rows.append(subtotal)
    if other_rows:
        rows.extend(other_rows)

    for grand_label, member_categories in grand_total_groups:
        members = [category_subtotals[c] for c in member_categories if c in category_subtotals]
        if not members:
            continue
        rows.append(_sum_rows(grand_label, members))

    if not rows:
        return None
    return pd.DataFrame(rows)[["Description", target_date]]


def synthesize_balance_sheet_and_income_statement(
    dfs: Dict[str, pd.DataFrame],
    mappings: Dict[str, Any],
) -> Dict[str, Any]:
    """Builds a BS/IS summary purely from already-extracted schedule tabs
    (dfs), for workbooks with no literal "Financials"-style sheet to read
    one from. Each account's own Total/合计 row becomes its BS/IS line,
    grouped by mappings.yml's `category` field with category subtotals and
    a grand total appended -- matching
    extract_balance_sheet_and_income_statement's output shape
    ({'balance_sheet', 'income_statement', 'project_name'}) so downstream
    consumers (pptx.py's embed_financial_tables, cover-page account
    ordering) need no changes.

    Deliberately NOT wired into reconcile_financial_statements -- comparing
    each schedule tab's own total against itself would be a trivial
    self-match, not a real cross-check. This is purely to give the PPTX
    embedded BS/IS summary table something to render when no source
    Financials sheet exists.

    IS is intentionally coarse (Total revenue - Total expenses = Net
    profit only, no Gross profit/Operating profit tiers) since
    mappings.yml's IS accounts only carry Revenue/Expenses/Others
    categories, not finer COGS/SG&A/non-operating groupings.
    """
    balance_sheet = _synthesize_statement(
        dfs, mappings, "BS", _SYNTHETIC_BS_CATEGORY_ORDER, _SYNTHETIC_BS_GRAND_TOTAL_GROUPS,
    )
    income_statement = _synthesize_statement(
        dfs, mappings, "IS", _SYNTHETIC_IS_CATEGORY_ORDER, [],
    )
    if income_statement is not None:
        date_cols = [c for c in income_statement.columns if c != "Description"]
        rev = income_statement[income_statement["Description"] == "Total revenue"]
        exp = income_statement[income_statement["Description"] == "Total expenses"]
        if not rev.empty and not exp.empty:
            net_profit = {"Description": "Net profit"}
            for c in date_cols:
                net_profit[c] = float(rev.iloc[0][c] or 0.0) - float(exp.iloc[0][c] or 0.0)
            income_statement = pd.concat([income_statement, pd.DataFrame([net_profit])], ignore_index=True)

    return {
        "balance_sheet": balance_sheet,
        "income_statement": income_statement,
        "project_name": None,
    }


# Example usage and testing
if __name__ == "__main__":
    # Example: Extract BS and IS from single sheet
    print("="*80)
    print("Example: Extract Balance Sheet and Income Statement from Single Sheet")
    print("="*80)
    
    workbook_path = "databook.xlsx"
    sheet_name = "Financial Statements"  # Sheet containing both BS and IS
    
    results = extract_balance_sheet_and_income_statement(
        workbook_path=workbook_path,
        sheet_name=sheet_name,
        debug=True  # Enable debugging
    )
    
    print(f"\n{'='*80}")
    print("EXTRACTION SUMMARY")
    print(f"{'='*80}")
    
    # Show project name
    if results['project_name']:
        print(f"✅ Project Name: {results['project_name']}")
    else:
        print("❌ Project Name: Not found")
    
    # Show Balance Sheet
    if results['balance_sheet'] is not None:
        print(f"\n✅ Balance Sheet Extracted:")
        print(f"   Total rows: {len(results['balance_sheet'])}")
        print(f"   Columns: {list(results['balance_sheet'].columns)}")
        print(f"   Sample data:")
        print(results['balance_sheet'].head(5))
    else:
        print("\n❌ Balance Sheet: Not found")
    
    # Show Income Statement
    if results['income_statement'] is not None:
        print(f"\n✅ Income Statement Extracted:")
        print(f"   Total rows: {len(results['income_statement'])}")
        print(f"   Columns: {list(results['income_statement'].columns)}")
        print(f"   Sample data:")
        print(results['income_statement'].head(5))
    else:
        print("\n❌ Income Statement: Not found")
    
    # Example: Access specific account data
    if results['balance_sheet'] is not None:
        print(f"\n{'='*80}")
        print("Example: Access specific account")
        print(f"{'='*80}")
        
        # Find account with description containing keyword
        cash_data = results['balance_sheet'][
            results['balance_sheet']['Description'].str.contains('货币资金', na=False)
        ]
        
        if not cash_data.empty:
            print("货币资金 (Cash) data:")
            print(cash_data.to_string())
            
            # Get values for each date
            for col in cash_data.columns:
                if col != 'Description':
                    value = cash_data.iloc[0][col]
                    print(f"  {col}: {value:,.0f}")
# --- end workbook/statements.py ---
