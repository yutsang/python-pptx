from __future__ import annotations


from typing import Any, Dict, List, Optional

import pandas as pd

_TOTAL_KEYWORDS = [
    "total",
    "合计",
    "总计",
    "subtotal",
    "sub-total",
    "sub total",
]


def _period_analysis_columns(analysis_df: pd.DataFrame) -> List[Any]:
    """The genuine PERIOD columns of a prompt-analysis df, in order.

    Column 0 is the description/block title, and column 1 is ALWAYS the
    INTERNAL_ROW_KEY bookkeeping column (`__source_row_idx`, the row's
    position in the source sheet) -- see _build_prompt_analysis_df, which
    constructs every row as {block_title, INTERNAL_ROW_KEY, *dates}. A bare
    `columns[1:]` therefore treats that row index as if it were the FIRST
    reporting period: confirmed live on real data, where an account's
    trend_summary came back with `start_period="__source_row_idx"` and
    `start_value=14.0` (literally the sheet row number), and every derived
    figure -- net_change, series_direction, the first delta -- was computed
    against that non-financial baseline before being handed to the LLM as
    trend evidence. Same class of bug as the one fixed in
    filter_zero_value_rows (which had assumed columns[1] was a value column).
    """
    from .schedules import INTERNAL_ROW_KEY  # local: breaks the analysis<->schedules import cycle
    return [col for col in analysis_df.columns[1:] if str(col) != INTERNAL_ROW_KEY]


def _trend_direction(values: List[float]) -> str:
    deltas = [curr - prev for prev, curr in zip(values, values[1:])]
    positive = any(delta > 0 for delta in deltas)
    negative = any(delta < 0 for delta in deltas)
    if positive and negative:
        return "volatile"
    if positive:
        return "rising"
    if negative:
        return "falling"
    return "flat"


def _select_trend_focus_row(analysis_df: pd.DataFrame) -> Optional[pd.Series]:
    if analysis_df is None or analysis_df.empty:
        return None
    desc_col = str(analysis_df.columns[0])
    descriptions = analysis_df[desc_col].astype(str)
    total_mask = descriptions.str.lower().str.contains("|".join(_TOTAL_KEYWORDS), regex=True)
    if total_mask.any():
        return analysis_df[total_mask].iloc[0]
    non_zero_mask = analysis_df.iloc[:, 1:].fillna(0).abs().sum(axis=1) > 0
    if non_zero_mask.any():
        return analysis_df[non_zero_mask].iloc[-1]
    return analysis_df.iloc[-1]


def build_trend_summary(analysis_df: pd.DataFrame) -> Dict[str, Any]:
    focus_row = _select_trend_focus_row(analysis_df)
    if focus_row is None or len(analysis_df.columns) < 3:
        return {}

    period_cols = _period_analysis_columns(analysis_df)
    if len(period_cols) < 2:
        return {}
    periods = [str(col) for col in period_cols]
    values = [float(focus_row[col] or 0) for col in period_cols]
    deltas = [
        {
            "from_period": periods[idx],
            "to_period": periods[idx + 1],
            "delta": values[idx + 1] - values[idx],
        }
        for idx in range(len(values) - 1)
    ]

    largest_increase = max(deltas, key=lambda item: item["delta"]) if deltas else None
    if largest_increase and largest_increase["delta"] <= 0:
        largest_increase = None
    largest_decrease = min(deltas, key=lambda item: item["delta"]) if deltas else None
    if largest_decrease and largest_decrease["delta"] >= 0:
        largest_decrease = None

    return {
        "focus_description": str(focus_row.iloc[0]),
        "series_direction": _trend_direction(values),
        "start_period": periods[0],
        "end_period": periods[-1],
        "start_value": values[0],
        "end_value": values[-1],
        "net_change": values[-1] - values[0],
        "largest_increase": largest_increase,
        "largest_decrease": largest_decrease,
    }


def _change_direction(prev_value: float, curr_value: float) -> str:
    if prev_value == 0 and curr_value > 0:
        return "new_increase"
    if prev_value == 0 and curr_value < 0:
        return "new_decrease"
    if curr_value > prev_value:
        return "increase"
    if curr_value < prev_value:
        return "decrease"
    return "flat"


# A period-on-period swing at or above this, relative to the line's own
# opening value, is worth a sentence even when the amount is small. 30% is
# the project team's own stated threshold.
MATERIAL_PCT_CHANGE = 30.0


def build_significant_movements(analysis_df: pd.DataFrame, max_items: int = 3) -> List[Dict[str, Any]]:
    if analysis_df is None or analysis_df.empty or len(analysis_df.columns) < 3:
        return []

    period_cols = _period_analysis_columns(analysis_df)
    if len(period_cols) < 2:
        return []
    periods = [str(col) for col in period_cols]
    movement_candidates: List[Dict[str, Any]] = []
    for _, row in analysis_df.iterrows():
        description = str(row.iloc[0]).strip()
        row_values = [float(row[col] or 0) for col in period_cols]
        best_movement = None
        for idx in range(len(row_values) - 1):
            prev_value = row_values[idx]
            curr_value = row_values[idx + 1]
            delta = curr_value - prev_value
            candidate = {
                "description": description,
                "from_period": periods[idx],
                "to_period": periods[idx + 1],
                "from_value": prev_value,
                "to_value": curr_value,
                "delta": delta,
                "abs_delta": abs(delta),
                "direction": _change_direction(prev_value, curr_value),
            }
            if best_movement is None or candidate["abs_delta"] > best_movement["abs_delta"]:
                best_movement = candidate
        if best_movement and best_movement["abs_delta"] > 0:
            movement_candidates.append(best_movement)

    if not movement_candidates:
        return []

    total_change = sum(item["abs_delta"] for item in movement_candidates)
    if total_change <= 0:
        return []

    def _pct_change(item: Dict[str, Any]) -> Optional[float]:
        base = abs(item["from_value"])
        if base < 0.01:
            return None  # from nil -- a percentage would be meaningless
        return round((item["delta"] / base) * 100, 1)

    # Ranked by share of the account's total movement, as before, but a line
    # that swung hard RELATIVE TO ITSELF now also qualifies even when its
    # absolute delta is small. Ranking on absolute size alone hid exactly the
    # analytically interesting cases -- a minor line doubling is a finding, a
    # 3% drift in the largest line is not -- and the model had no way to tell
    # them apart because it was given raw values and had to do the arithmetic
    # itself, which is where it invents numbers.
    significant = []
    for item in sorted(movement_candidates, key=lambda entry: entry["abs_delta"], reverse=True):
        percent_of_total_change = (item["abs_delta"] / total_change) * 100
        pct_change = _pct_change(item)
        material = (
            percent_of_total_change >= 25
            or (pct_change is not None and abs(pct_change) >= MATERIAL_PCT_CHANGE)
        )
        if not material:
            continue
        significant.append(
            {
                "description": item["description"],
                "from_period": item["from_period"],
                "to_period": item["to_period"],
                "from_value": item["from_value"],
                "to_value": item["to_value"],
                "delta": item["delta"],
                "direction": item["direction"],
                "percent_of_total_change": round(percent_of_total_change, 1),
                # Stated rather than left to be derived, so a commentary that
                # cites a percentage is quoting a supplied figure instead of
                # constructing one -- the same rule that fixed the AR
                # quantified-recommendation hallucination.
                "percent_change": pct_change,
                "exceeds_materiality": bool(
                    pct_change is not None and abs(pct_change) >= MATERIAL_PCT_CHANGE
                ),
            }
        )
        if len(significant) >= max_items:
            break
    return significant
# --- end workbook/analysis.py ---
