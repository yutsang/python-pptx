from __future__ import annotations

"""Display-focused formatting helpers for localized financial tables."""

from typing import Iterable, Union

import pandas as pd


def format_number_chinese(value: Union[float, int], language: str = "Chi") -> str:
    if pd.isna(value) or value == 0:
        return "0"

    is_negative = value < 0
    abs_value = abs(value)

    if language == "Chi":
        if abs_value < 10000:
            result = f"{abs_value:,.0f}元"
        else:
            wan_value = abs_value / 10000
            if wan_value >= 10000:
                yi_value = wan_value / 10000
                result = f"{yi_value:.1f}亿元"
            else:
                result = f"{wan_value:.1f}万元"
        return f"人民币-{result}" if is_negative else f"人民币{result}"

    if abs_value < 10000:
        result = f"{abs_value:,.0f}"
    elif abs_value < 1000000:
        result = f"{abs_value / 1000:.1f}K"
    else:
        result = f"{abs_value / 1000000:.2f} million"
    return f"CNY -{result}" if is_negative else f"CNY {result}"


def convert_scientific_to_normal(value: Union[str, float, int]) -> float:
    if isinstance(value, str):
        try:
            return float(value)
        except ValueError:
            return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    return 0.0


def detect_and_format_amount(value: Union[str, float, int], language: str = "Chi") -> str:
    return format_number_chinese(convert_scientific_to_normal(value), language)


def format_retained_earnings(value: Union[float, int], language: str = "Chi") -> tuple[str, str]:
    if language == "Chi":
        if value < 0:
            return "未弥补亏损", format_number_chinese(abs(value), language)
        return "未分配利润", format_number_chinese(value, language)

    if value < 0:
        return "Accumulated Losses", format_number_chinese(abs(value), language)
    return "Retained Earnings", format_number_chinese(value, language)


def format_dataframe_values(df: pd.DataFrame, language: str = "Chi") -> pd.DataFrame:
    if df is None or df.empty:
        return df

    formatted_df = df.copy()
    for col in formatted_df.columns[1:]:
        if pd.api.types.is_numeric_dtype(formatted_df[col]):
            formatted_df[col] = formatted_df[col].apply(
                lambda x: detect_and_format_amount(x, language)
            )
    return formatted_df


def format_value_by_language(value, language: str, _account_name=None) -> str:
    if pd.isna(value):
        return ""
    if value == 0:
        return "0"

    is_negative = value < 0
    abs_value = abs(value)

    if language == "Chi":
        if abs_value >= 100000000:
            formatted = f"{abs_value / 100000000:.2f}亿"
        elif abs_value >= 10000:
            formatted = f"{abs_value / 10000:.1f}万"
        else:
            formatted = f"{abs_value:.0f}"
        return f"-{formatted}" if is_negative else formatted

    if abs_value >= 1000000:
        formatted = f"{abs_value / 1000000:.2f} million"
    elif abs_value >= 1000:
        formatted = f"{abs_value / 1000:.1f}K"
    else:
        formatted = f"{abs_value:,.0f}"
    return f"-{formatted}" if is_negative else formatted


def handle_retained_earnings(df: pd.DataFrame, language: str) -> pd.DataFrame:
    if df is None or df.empty:
        return df

    df_modified = df.copy()
    desc_col = df_modified.columns[0]
    rename_map: dict[str, str] = {}
    retained_earnings_keywords_chi = ["未分配利润", "留存收益", "盈余公积"]
    retained_earnings_keywords_eng = [
        "Retained Earnings",
        "Retained earnings",
        "retained earnings",
        "Accumulated Profit",
        "Surplus Reserve",
    ]

    for idx, row in df_modified.iterrows():
        description = str(row[desc_col])
        if language == "Chi":
            is_retained_earnings = any(keyword in description for keyword in retained_earnings_keywords_chi)
        else:
            is_retained_earnings = any(keyword in description for keyword in retained_earnings_keywords_eng)

        if not is_retained_earnings:
            continue

        value_col = df_modified.columns[1]
        value = row[value_col]
        if pd.isna(value) or value >= 0:
            continue

        if language == "Chi":
            if "未分配利润" in description:
                df_modified.at[idx, desc_col] = description.replace("未分配利润", "未弥补亏损")
            elif "留存收益" in description:
                df_modified.at[idx, desc_col] = description.replace("留存收益", "累计亏损")
            elif "盈余公积" in description:
                df_modified.at[idx, desc_col] = description.replace("盈余公积", "亏损")
        else:
            if "Retained Earnings" in description:
                df_modified.at[idx, desc_col] = description.replace("Retained Earnings", "Accumulated Losses")
            elif "retained earnings" in description:
                df_modified.at[idx, desc_col] = description.replace("retained earnings", "accumulated losses")
            elif "Accumulated Profit" in description:
                df_modified.at[idx, desc_col] = description.replace("Accumulated Profit", "Accumulated Losses")

        rename_map[description] = str(df_modified.at[idx, desc_col]).strip()
        formatted_col = f"{value_col}_formatted"
        if formatted_col in df_modified.columns:
            df_modified.at[idx, formatted_col] = format_value_by_language(abs(value), language)

    if rename_map:
        existing = dict(df_modified.attrs.get("display_description_map") or {})
        existing.update(rename_map)
        for renamed in rename_map.values():
            existing[renamed] = renamed
        df_modified.attrs["display_description_map"] = existing
    return df_modified


def add_language_display_columns(df: pd.DataFrame, language: str) -> pd.DataFrame:
    if df is None or df.empty or not language:
        return df

    df_modified = df.copy()
    for col in df_modified.columns[1:]:
        if str(col).endswith("_formatted"):
            continue
        if not pd.api.types.is_numeric_dtype(df_modified[col]):
            continue
        df_modified[f"{col}_formatted"] = df_modified[col].apply(
            lambda x: format_value_by_language(x, language)
        )
    return handle_retained_earnings(df_modified, language)


def prepare_display_dataframe(
    df: pd.DataFrame,
    *,
    drop_columns: Iterable[str] = (),
) -> pd.DataFrame:
    display_df = df.copy()
    drop_columns = {str(column) for column in drop_columns}
    if drop_columns:
        existing_drop_columns = [column for column in display_df.columns if str(column) in drop_columns]
        if existing_drop_columns:
            display_df = display_df.drop(columns=existing_drop_columns)

    if not display_df.empty and len(display_df.columns) > 1:
        original_columns = list(display_df.columns)
        column_lookup = {str(column): column for column in original_columns}
        display_columns = [original_columns[0]]
        rename_map = {}

        for column in original_columns[1:]:
            column_name = str(column)
            if column_name.endswith("_formatted"):
                base_name = column_name.replace("_formatted", "")
                if base_name in column_lookup:
                    continue
                display_columns.append(column)
                continue

            formatted_column = column_lookup.get(f"{column_name}_formatted")
            if formatted_column is not None:
                display_columns.append(formatted_column)
                rename_map[formatted_column] = column_name
            else:
                display_columns.append(column)

        display_df = display_df[display_columns]
        if rename_map:
            display_df = display_df.rename(columns=rename_map)
    return display_df


def stringify_display_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    display_df = df.copy()
    for col in display_df.columns:
        if pd.api.types.is_numeric_dtype(display_df[col]):
            display_df[col] = display_df[col].apply(lambda x: f"{x:,.0f}" if pd.notna(x) else "")
        else:
            display_df[col] = display_df[col].fillna("").astype(str)
    return display_df
