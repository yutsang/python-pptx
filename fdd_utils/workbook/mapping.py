from __future__ import annotations

from typing import Any, Dict, Iterable, List, Optional

import pandas as pd

from ..financial_common import load_yaml_file, package_file_path
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


def load_mappings(mappings_path: Optional[str] = None) -> Dict[str, Any]:
    return load_yaml_file(mappings_path or package_file_path("mappings.yml"))


def get_effective_mappings(
    base_mappings: Dict[str, Any],
    resolution: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    effective = dict(base_mappings or {})
    for mapping_key, config in ((resolution or {}).get("dynamic_mappings") or {}).items():
        if mapping_key and isinstance(config, dict):
            effective[mapping_key] = config
    return effective


def should_skip_account_label(account_name: str) -> bool:
    text = str(account_name or "").strip()
    if not text:
        return False

    lowered = text.lower()
    if text.endswith(("合计", "总计", "小计")):
        return True
    if lowered.startswith("total "):
        return True
    return any(keyword in lowered for keyword in SUMMARY_ACCOUNT_SKIP_KEYWORDS)


def normalize_mapping_label(account_name: str) -> str:
    normalized = (account_name or "").strip().lower()
    for suffix in ["：", ":", "（", "）", "(", ")"]:
        normalized = normalized.replace(suffix, "")
    return " ".join(normalized.split())


def iter_account_mappings(mappings: Dict[str, Any]) -> Iterable[tuple[str, Dict[str, Any]]]:
    for mapping_key, config in (mappings or {}).items():
        if str(mapping_key).startswith("_") or not isinstance(config, dict):
            continue
        yield str(mapping_key), config


def find_mapping_key(account_name: str, mappings: Dict[str, Any]) -> str | None:
    """Find the canonical mapping key for an account name or alias."""
    if account_name in mappings:
        return account_name

    normalized_account = normalize_mapping_label(account_name)
    for mapping_key, config in iter_account_mappings(mappings):
        aliases = config.get("aliases", [])
        if account_name in aliases:
            return mapping_key
        normalized_aliases = {normalize_mapping_label(alias) for alias in aliases}
        if normalized_account and normalized_account in normalized_aliases:
            return mapping_key
    return None


def split_accounts_by_type(
    account_names: List[str],
    mappings: Dict[str, Any],
    dfs: Dict[str, "pd.DataFrame"] | None = None,
) -> tuple[List[str], List[str], List[str]]:
    """Preserve account order while grouping by mapping type.

    Falls back to ``df.attrs["integrity"]["statement_type"]`` when the
    account cannot be found in *mappings*, so dynamically-resolved
    accounts are still classified as BS or IS instead of "other".
    """
    bs_accounts: List[str] = []
    is_accounts: List[str] = []
    other_accounts: List[str] = []

    for account_name in account_names:
        account_type = ""
        mapping_key = find_mapping_key(account_name, mappings)
        if mapping_key:
            account_type = mappings[mapping_key].get("type", "")

        # Fallback: read statement_type from the DataFrame attrs
        if account_type not in ("BS", "IS") and dfs and account_name in dfs:
            df = dfs[account_name]
            integrity = getattr(df, "attrs", {}).get("integrity") or {}
            account_type = integrity.get("statement_type", "")

        if account_type == "BS":
            bs_accounts.append(account_name)
        elif account_type == "IS":
            is_accounts.append(account_name)
        else:
            other_accounts.append(account_name)

    return bs_accounts, is_accounts, other_accounts


def build_account_mapping_diagnostics(
    account_names: Iterable[str],
    mappings: Dict[str, Any],
) -> pd.DataFrame:
    rows: List[Dict[str, str]] = []

    for account_name in account_names:
        mapping_key = find_mapping_key(account_name, mappings)
        if not mapping_key:
            rows.append(
                {
                    "account_name": str(account_name),
                    "mapping_key": "",
                    "account_type": "",
                    "classification": "other",
                    "reason": "No mappings.yml key or alias matched this account/tab name.",
                }
            )
            continue

        mapping_config = mappings.get(mapping_key, {})
        account_type = str(mapping_config.get("type", "") or "")
        if account_type in {"BS", "IS"}:
            classification = account_type
            reason = f"Mapped to {account_type} via '{mapping_key}'."
        else:
            classification = "other"
            reason = f"mapped type '{account_type or 'blank'}' is not classified as BS or IS."

        rows.append(
            {
                "account_name": str(account_name),
                "mapping_key": mapping_key,
                "account_type": account_type,
                "classification": classification,
                "reason": reason,
            }
        )

    return pd.DataFrame(
        rows,
        columns=["account_name", "mapping_key", "account_type", "classification", "reason"],
    )
# --- end workbook/mapping.py ---
