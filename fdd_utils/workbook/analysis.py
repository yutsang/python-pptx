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


#: Suffix a part-year column carries once it has been scaled to twelve months.
#: The same wording schedules.py gives the annualised projection column, so a
#: reader meets one name for the one idea.
ANNUALISED_SUFFIX = " annualised"


def partial_tail_months(analysis_df: pd.DataFrame, statement_type: Optional[str] = None) -> Optional[int]:
    """Months the LAST period column covers when it is a part-year, else None.

    Only an income-statement frame can have one: a balance is a point in time
    and a March balance is as good as a December one, while three months of
    revenue against twelve is not a comparison at all.

    Read from the frame's own date columns, NOT attrs["annualization_months"].
    That attribute describes the selected projection column: an account whose
    latest period is nil projects an EARLIER year, so it carries no months even
    though its analysis frame ends on the same part-year column as every other
    account's.
    """
    if analysis_df is None or getattr(analysis_df, "empty", True):
        return None
    if statement_type is None:
        statement_type = ((getattr(analysis_df, "attrs", None) or {}).get("integrity") or {}).get("statement_type")
    if str(statement_type or "").strip().upper() != "IS":
        return None
    periods = [str(col) for col in _period_analysis_columns(analysis_df) if not str(col).endswith("_formatted")]
    if len(periods) < 2:
        return None
    from .schedules import infer_partial_year_annualization  # local: analysis<->schedules cycle
    try:
        inferred = infer_partial_year_annualization("IS", periods, periods[-1])
    except Exception:
        return None
    months = inferred.get("months")
    if inferred.get("annualized") and isinstance(months, (int, float)) and 0 < months < 12:
        return int(months)
    return None


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


def build_trend_summary(analysis_df: pd.DataFrame, statement_type: Optional[str] = None) -> Dict[str, Any]:
    """The account's own line across every period, as figures the model quotes.

    A part-year tail is put on a twelve-month basis before anything is derived
    from it. Left raw, three months of an income-statement line sat beside
    three full years and the summary reported the step down to it as the
    series' "largest_decrease" -- a fall that is nothing but the shorter
    period. The raw figure is kept alongside, named, because the period's own
    amount is what a sentence about that period states.
    """
    focus_row = _select_trend_focus_row(analysis_df)
    if focus_row is None or len(analysis_df.columns) < 3:
        return {}

    period_cols = _period_analysis_columns(analysis_df)
    if len(period_cols) < 2:
        return {}
    periods = [str(col) for col in period_cols]
    values = [float(focus_row[col] or 0) for col in period_cols]
    tail_months = partial_tail_months(analysis_df, statement_type)
    raw_end_value = values[-1]
    if tail_months:
        values[-1] = raw_end_value * 12.0 / tail_months
        periods[-1] = periods[-1] + ANNUALISED_SUFFIX
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

    summary = {
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
    if tail_months:
        # Only on a part-year frame, so every other account's summary -- and
        # the prompt rendered from it -- is byte-for-byte what it was.
        summary["end_period_months"] = tail_months
        summary["raw_end_value"] = raw_end_value
        summary["basis"] = (
            f"end_value and every change into {periods[-1]} are on a 12-month basis "
            f"(x12/{tail_months}); raw_end_value is the {tail_months}-month figure as reported"
        )
    return summary


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


def build_significant_movements(analysis_df: pd.DataFrame, max_items: int = 3,
                                statement_type: Optional[str] = None) -> List[Dict[str, Any]]:
    """The lines that moved most between adjacent periods, percentage supplied.

    The step into a part-year tail is measured on a twelve-month basis. It was
    measured raw, and the prompt requires every material entry here to appear
    in the commentary with percent_change quoted as supplied -- so a real deck
    said revenue "annualised to 3,675.8万元, down 67.0% on 2,787.2万元" (the
    annualised figure is the larger of the two; -67.0% is three months over
    twelve), with finance costs "down 75.0%" and administrative expenses "down
    72.9%" the same way, each followed by a cause for a fall that never
    happened. The model was quoting what it had been handed.
    """
    if analysis_df is None or analysis_df.empty or len(analysis_df.columns) < 3:
        return []

    period_cols = _period_analysis_columns(analysis_df)
    if len(period_cols) < 2:
        return []
    periods = [str(col) for col in period_cols]
    tail_months = partial_tail_months(analysis_df, statement_type)
    if tail_months:
        periods[-1] = periods[-1] + ANNUALISED_SUFFIX
    movement_candidates: List[Dict[str, Any]] = []
    for _, row in analysis_df.iterrows():
        description = str(row.iloc[0]).strip()
        row_values = [float(row[col] or 0) for col in period_cols]
        if tail_months:
            row_values[-1] = row_values[-1] * 12.0 / tail_months
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
        if tail_months:
            # Present on every row of a part-year frame, absent everywhere
            # else: a column that exists on some rows only prints as NaN.
            significant[-1]["basis"] = (
                f"annualised x12/{tail_months}" if item["to_period"] == periods[-1] else "as reported"
            )
        if len(significant) >= max_items:
            break
    return significant
# --- end workbook/analysis.py ---
