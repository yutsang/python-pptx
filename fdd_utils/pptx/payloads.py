from __future__ import annotations

# re-added: bound by an import in another section of the pre-split module
import re


from .text import clean_content_quotes
from typing import Any, Dict, Iterable, List, Optional

import pandas as pd

from ..financial_common import (
    contains_chinese_text,
    contains_predominantly_chinese_text,
    get_pipeline_result_text,
    load_yaml_file,
    package_file_path,
)
from ..keyword_registry import (
    STATEMENT_ORDER_SKIP_KEYWORDS,
    SUMMARY_ACCOUNT_SKIP_KEYWORDS,
    translate_category_to_chinese,
    translate_statement_line_to_chinese,
)
from ..workbook import find_mapping_key


PPTX_DEFAULT_SETTINGS: Dict[str, Any] = {
    "max_commentary_slides_per_statement": 4,
    "executive_summary": {
        "target_words_eng": 110,
        "target_chars_chi": 144,
        "max_sentences_eng": 4,
        "max_sentences_chi": 4,
        "max_tokens": 240,
        "validation_max_tokens": 180,
        "max_input_chars": 1400,
        "max_numeric_sentences": 2,
        "max_workers": 2,
        "enable_validation": True,
        "generation_temperature": 0.2,
        "validation_temperature": 0.1,
    },
    "commentary_packing": {
        "use_pillow_text_fitting": True,
        # Repeatedly bumped up (1.08 -> 1.15 -> 1.25, and the BS override
        # further to 1.13 -> 1.47) in response to page fill plateauing too
        # low -- but the real cause turned out to be _real_para_gap_pt/
        # _real_line_spacing assuming a 6-9pt inter-paragraph gap and 0.9
        # Chinese line spacing that _fill_text_main_bullets_with_category_
        # and_key never actually applies (it hardcodes a flat 3pt gap and
        # 1.0 spacing for every paragraph, any language) plus a capacity
        # formula that floored away up to a full extra line on every box.
        # Once those are fixed at the source (real capacity ~40-50% higher
        # than before), this relax tier goes back to being a small genuine
        # second-chance buffer instead of a compensating hack for an
        # undersized first tier.
        "shape_height_utilization": 1.00,
        "minimum_slot_lines": 22,
        "split_min_remaining_lines": 3,
        "split_min_content_lines": 5,
        # Lowered: pull a whole bullet forward into a slot even when the slot
        # is already 50% full (was 74%). This stops the first IS slot from
        # sitting at 35% fill while later slots are full.
        "move_whole_min_fill_ratio": 0.50,
        "target_fill_min_ratio": 0.95,
        "target_fill_max_ratio": 1.00,
        "ppt_length_ratio": 0.84,
        "ppt_min_chars_eng": 190,
        "ppt_min_chars_chi": 110,
        "ppt_max_sentences_eng": 6,
        "ppt_max_sentences_chi": 5,
        "ppt_max_numeric_sentences": 2,
        "category_line_cost": 0.95,
        "key_line_cost": 1.0,
        "continuation_spacing_penalty": 0.15,
        "line_height_padding_pt": 1.6,
        "split_slot_height_penalty": 1.02,
        "width_scale_min": 0.9,
        "width_scale_max": 1.22,
        "chars_per_line": {
            "single": {"eng": 100, "chi": 50},
            "L": {"eng": 56, "chi": 30},
            "R": {"eng": 56, "chi": 30},
            "default": {"eng": 66, "chi": 36},
        },
        "statement_overrides": {
            "BS": {
                "shape_height_utilization": 1.00,
                "line_height_padding_pt": 1.3,
                "chars_per_line": {
                    "single": {"eng": 106},
                    "L": {"eng": 59},
                    "R": {"eng": 59},
                    "default": {"eng": 69},
                },
            },
        },
    },
}


def _merge_nested_dict(base: Dict[str, Any], overrides: Dict[str, Any]) -> Dict[str, Any]:
    merged = dict(base or {})
    for key, value in (overrides or {}).items():
        if isinstance(value, dict) and isinstance(merged.get(key), dict):
            merged[key] = _merge_nested_dict(merged[key], value)
        else:
            merged[key] = value
    return merged


def _load_pptx_settings(config_path: Optional[str] = None) -> Dict[str, Any]:
    config = load_yaml_file(config_path or package_file_path("config.yml"))
    return _merge_nested_dict(PPTX_DEFAULT_SETTINGS, (config or {}).get("pptx") or {})


def _split_text_sentences(text: str, is_chinese: bool) -> List[str]:
    normalized = str(text or "").strip()
    if not normalized:
        return []
    if is_chinese:
        parts = re.split(r"(?<=[。！？；])", normalized)
    else:
        parts = re.split(r"(?<=[.!?;])\s+", normalized)
    return [part.strip() for part in parts if part and part.strip()]


def _join_text_sentences(sentences: List[str], is_chinese: bool) -> str:
    cleaned = [str(sentence or "").strip() for sentence in sentences if str(sentence or "").strip()]
    if not cleaned:
        return ""
    return "".join(cleaned) if is_chinese else " ".join(cleaned)


def _sentence_is_numeric_heavy(sentence: str) -> bool:
    text = str(sentence or "")
    numeric_tokens = re.findall(r"\d[\d,.\-]*%?|USD|HKD|RMB|CNY|EUR|JPY|\$", text, flags=re.IGNORECASE)
    return len(numeric_tokens) >= 2




#: Legal-form tails and place-of-registration brackets carry NO identifying
#: information -- they say what kind of company it is and where it was
#: registered, not which company. The reference deck strips both throughout
#: (浙江卓圣, not 浙江卓圣物业管理有限公司) and only writes a full legal name
#: when first naming a specific contract counterparty.
_LEGAL_FORM_TAILS = (
    "股份有限责任公司", "股份有限公司", "有限责任公司", "私人有限公司",
    "有限合伙企业", "会计师事务所", "律师事务所", "有限公司",
)
#: Removed only when the bracket holds a short place name. A long bracket is
#: usually a real qualifier and gets left alone.
_REGISTRATION_BRACKET = re.compile(r"[（(][一-鿿]{2,4}[）)](?=[一-鿿]*(?:%s))"
                                   % "|".join(_LEGAL_FORM_TAILS))


def shorten_company_names(text: str) -> str:
    """Strip legal-form tails and registration brackets from company names.

    Done deterministically, NOT by prompt. Two separate prompt attempts
    failed outright: the databook's own row labels ARE the full legal names
    (e.g. a row literally reading *某某咨询管理有限公司), and the prompts
    carry a much older, much stronger instruction to reproduce entity names
    from the data exactly. The model correctly followed the stronger rule.
    Shortening a name is a mechanical string operation, so it belongs here
    rather than in an instruction that has to win an argument.

    Deliberately conservative: only the legal form and the registration
    bracket are removed, because neither identifies the party. Trailing
    business descriptors (物业管理 / 企业管理) are NOT stripped -- that would
    reach the reference deck's brevity but risks making two counterparties
    read the same.
    """
    body = str(text or "")
    if not body:
        return body
    body = _REGISTRATION_BRACKET.sub("", body)
    for tail in _LEGAL_FORM_TAILS:
        body = body.replace(tail, "")
    return body


def _normalize_slide_commentary_text(text: str) -> str:
    normalized = clean_content_quotes(str(text or ""))
    if not normalized:
        return ""
    normalized = normalized.replace("\r\n", "\n")
    normalized = re.sub(r"[ \t]+", " ", normalized)
    normalized = re.sub(r"\n{3,}", "\n\n", normalized)
    normalized = shorten_company_names(normalized)
    return normalized.strip()


def _extract_summary(content):
    text = str(content or "").strip()
    if not text:
        return ""
    if _looks_like_blocked_ai_content(text):
        return ""
    return text


def _looks_like_blocked_ai_content(text: str) -> bool:
    normalized = str(text or "").strip()
    if not normalized:
        return False
    lowered = normalized.lower()
    blocked_markers = (
        "<!doctype html",
        "<html",
        "ac_block_page",
        "sp.eagleyun.cn",
        "form.submit()",
        "api.deepseek.com",
        "request_uri",
        "request_user_agent",
    )
    return any(marker in lowered for marker in blocked_markers)


def _extract_final_content(result_dict):
    # Defence in depth: strip any Qwen3 <think> block that slipped through (e.g.
    # via run_generator_reprompt, which skips the _ensure pass) before it can
    # render into a no-autofit text box and overflow / leak reasoning.
    from fdd_utils.ai import strip_thinking
    return strip_thinking(get_pipeline_result_text(result_dict))


def _build_statement_order(
    financial_statement_df: Optional[pd.DataFrame],
    mappings: Dict[str, Any],
) -> tuple[Dict[str, int], Dict[str, str]]:
    financial_statement_order: Dict[str, int] = {}
    statement_display_names: Dict[str, str] = {}
    if financial_statement_df is None or financial_statement_df.empty or len(financial_statement_df.columns) == 0:
        return financial_statement_order, statement_display_names

    first_col = financial_statement_df.iloc[:, 0]
    skip_keywords = STATEMENT_ORDER_SKIP_KEYWORDS
    for idx, account_name_in_statement in enumerate(first_col):
        if pd.isna(account_name_in_statement):
            continue

        account_name_str = str(account_name_in_statement).strip()
        if not account_name_str or any(skip in account_name_str.lower() for skip in skip_keywords):
            continue

        mapping_key = find_mapping_key(account_name_str, mappings)
        if mapping_key:
            financial_statement_order[mapping_key] = idx
            statement_display_names[mapping_key] = account_name_str

        financial_statement_order[account_name_str] = idx

    return financial_statement_order, statement_display_names


# Common Traditional -> Simplified character pairs that show up in
# mappings.yml's own Chinese aliases (e.g. "貨幣資金" vs "货币资金"). Not a
# general S/T converter -- just enough of this narrow FDD-account-name
# vocabulary to (a) normalize a Traditional alias to its Simplified spelling
# for detecting "these two aliases are the same concept, just different
# script" and (b) prefer Simplified for the final label, matching the
# convention CATEGORY_TRANSLATIONS_ZH already uses for section headers
# (fdd_utils/keyword_registry.py).
_TRADITIONAL_TO_SIMPLIFIED_PAIRS = {
    "貨": "货", "應": "应", "產": "产", "負": "负", "稅": "税", "資": "资",
    "積": "积", "準": "准", "幣": "币", "讓": "让", "長": "长", "遞": "递",
    "認": "认", "賬": "账", "債": "债", "現": "现", "後": "后", "裡": "里",
    "歸": "归", "屬": "属", "歷": "历", "業": "业", "當": "当", "項": "项",
    "餘": "余", "繳": "缴", "會": "会", "計": "计", "師": "师", "貴": "贵",
    "賣": "卖", "買": "买", "須": "须", "廠": "厂", "聯": "联", "繫": "系",
    "務": "务", "單": "单", "帳": "帐", "報": "报",
    # Audited (extracted directly from mappings.yml's own alias vocabulary,
    # see: "python3 -c 'grep aliases + CJK char diff'" in this commit's history)
    "內": "内", "動": "动", "發": "发", "實": "实", "損": "损", "據": "据",
    "攤": "摊", "權": "权", "減": "减", "潤": "润", "無": "无", "營": "营",
    "綜": "综", "職": "职", "譽": "誉", "財": "财", "費": "费", "賃": "赁",
    "預": "预",
}
_TRADITIONAL_TO_SIMPLIFIED = str.maketrans(_TRADITIONAL_TO_SIMPLIFIED_PAIRS)


def _find_chinese_display_name(mapping_key: str, fallback: str, mappings: Dict[str, Any]) -> str:
    """display_name (built in _build_statement_order) is whatever literal
    text sits in that account's row in the Financials-summary sheet -- for
    an English-labelled source databook (e.g. Kunshan's "Cash at bank and
    on hand"), that stays English even when the REPORT is being generated
    in Chinese, since nothing translates it. mappings.yml already carries a
    Chinese alias for essentially every account (used for matching Chinese
    source sheets) -- reuse one of those as the Chinese-report label instead
    of leaving the English source text in an otherwise fully-translated
    Chinese bullet.

    A single mapping_key's aliases list often mixes a precise term with
    broader catch-all synonyms for MATCHING purposes (e.g. Capital's aliases
    include both "实收资本"/paid-in capital AND "股东权益"/shareholders'
    equity -- correct for fuzzy-matching a client's sheet name, but "股东权益"
    would be a semantically WRONG display label for a paid-in-capital
    account). Picking "any CJK alias" isn't safe.

    Approach: group the CJK aliases by their Simplified-normalized spelling
    (so a Traditional/Simplified pair like "實收資本"/"实收资本" is treated as
    ONE concept, not two competing candidates), rank the resulting concepts
    by how close any of their members sit to where `fallback` itself
    appears in the alias list (aliases are typically authored in loosely
    paired EN/CN blocks, e.g. "...实收资本", "Paid-in capital" sit adjacent),
    and return the Simplified spelling of the winning concept."""
    config = mappings.get(mapping_key) if isinstance(mappings, dict) else None
    aliases = config.get("aliases") if isinstance(config, dict) else None
    if not isinstance(aliases, list) or not aliases:
        return fallback
    aliases = [str(a).strip() for a in aliases]
    chinese_indices = [i for i, a in enumerate(aliases) if contains_chinese_text(a)]
    if not chinese_indices:
        return fallback

    normalized_fallback = str(fallback or "").strip().lower()
    fallback_idx = next(
        (i for i, a in enumerate(aliases) if a.lower() == normalized_fallback),
        len(aliases) // 2,  # unknown position -- fall back to the list midpoint
    )

    concepts: Dict[str, Dict[str, Any]] = {}
    for i in chinese_indices:
        simplified = aliases[i].translate(_TRADITIONAL_TO_SIMPLIFIED)
        entry = concepts.setdefault(simplified, {"best_distance": None, "simplified_form": None})
        distance = abs(i - fallback_idx)
        if entry["best_distance"] is None or distance < entry["best_distance"]:
            entry["best_distance"] = distance
        if aliases[i] == simplified:
            entry["simplified_form"] = aliases[i]

    best_concept = min(concepts.items(), key=lambda kv: kv[1]["best_distance"])
    simplified_key, entry = best_concept
    return entry["simplified_form"] or simplified_key


def _translate_statement_row_label(label: str, mappings: Optional[Dict[str, Any]]) -> str:
    """Translates ONE row label from the embedded BS/IS summary table
    (embed_financial_tables) to Chinese, for a Chinese-language report.

    Unlike commentary bullets (which carry their own resolved mapping_key
    via build_pptx_structured_payloads), these rows come straight from
    extract_balance_sheet_and_income_statement's own parse of the raw
    Financials-sheet text -- there was previously NO translation path for
    them at all, so a Chinese report's embedded table stayed 100% English
    even though the title/commentary around it were fully translated.

    Two label classes need two different lookups: (1) individual account
    rows (e.g. "Cash at bank and on hand") resolve via the same
    mappings.yml alias machinery _find_chinese_display_name already uses
    for commentary, by first recovering the mapping_key with
    find_mapping_key; (2) statement-structure total/subtotal rows (e.g.
    "Total current assets") aren't mapping_key accounts at all, so they
    fall back to the small fixed STATEMENT_TOTAL_LINE_TRANSLATIONS_ZH
    table. Returns `label` unchanged if neither resolves (e.g. an
    already-Chinese source label, or a genuinely unmapped line) --
    partial coverage beats a blank or crashed cell.
    """
    label = str(label or "")
    if not label.strip() or contains_chinese_text(label):
        return label
    if mappings:
        mapping_key = find_mapping_key(label, mappings)
        if mapping_key:
            return _find_chinese_display_name(mapping_key, label, mappings)
    return translate_statement_line_to_chinese(label) or label


def _has_significant_balance(financial_data: Optional[pd.DataFrame]) -> bool:
    """Does this account carry a balance worth commenting on?

    The dataframe handed in is the PROJECTION frame -- one stage, ONE date
    column -- so scanning it answers "is the latest period non-zero", not the
    question actually being asked. An account that ran to nil by the reporting
    date but had real activity in earlier periods (a real Crescent 投资收益
    reads 0 at 2026-03-31 with prior-year movement) was silently dropped from
    the deck after its commentary had already been generated and paid for.

    workbook.py already computes exactly the right signal while parsing --
    any_period_nonzero_by_description, |v| >= 0.01 over ALL periods of the
    stage -- and filter_zero_value_rows already uses it to decide which ROWS
    survive. Use it here for the same decision at ACCOUNT level; fall back to
    the single-column scan only when the attribute is absent.
    """
    if financial_data is None or financial_data.empty:
        return True

    any_period = financial_data.attrs.get("any_period_nonzero_by_description")
    if isinstance(any_period, dict) and any_period:
        return any(bool(v) for v in any_period.values())

    numeric_cols = financial_data.select_dtypes(include=[float, int]).columns
    if len(numeric_cols) == 0:
        return True

    for col in numeric_cols:
        if (financial_data[col].abs() >= 0.01).any():
            return True
    return False


def build_pptx_structured_payloads(
    ai_results,
    mappings,
    bs_is_results=None,
    dfs=None,
):
    from .exporters import logger  # local: payloads imports the later exporters; module-level would be a cycle
    if not ai_results:
        return {"BS": [], "IS": []}

    balance_sheet_df = bs_is_results.get("balance_sheet") if bs_is_results else None
    income_statement_df = bs_is_results.get("income_statement") if bs_is_results else None
    bs_order, bs_display_names = _build_statement_order(balance_sheet_df, mappings)
    is_order, is_display_names = _build_statement_order(income_statement_df, mappings)

    payloads = {"BS": [], "IS": []}
    sortable_items = {"BS": [], "IS": []}

    for account_key, result in ai_results.items():
        mapping_key = find_mapping_key(account_key, mappings)
        if not mapping_key:
            continue

        account_type = mappings[mapping_key].get("type")
        if account_type not in {"BS", "IS"}:
            continue

        financial_data = dfs.get(account_key) if dfs and account_key in dfs else None
        if not _has_significant_balance(financial_data):
            # Dropping an account here throws away commentary the AI pipeline
            # has already produced and been paid for, and it happens silently
            # -- the only trace was a smaller "IS items: N". Say which account
            # and on what evidence, so the next run answers "why is 投资收益
            # missing" instead of another round of guessing.
            _attrs = getattr(financial_data, "attrs", {}) or {}
            _any = _attrs.get("any_period_nonzero_by_description")
            logger.warning(
                "Dropping %s from the deck: no period carries a balance. "
                "any_period_nonzero_by_description=%s, stage=%s, columns=%s",
                mapping_key,
                dict(_any) if isinstance(_any, dict) else _any,
                _attrs.get("prompt_analysis_stage"),
                list(getattr(financial_data, "columns", [])) or None,
            )
            continue

        final_content = _extract_final_content(result)
        commentary_text = (
            str(final_content).strip()
            if final_content and str(final_content).strip()
            else f"[No content generated for {account_key}]"
        )

        clause_reviews: List[Dict[str, Any]] = []
        if isinstance(result, dict):
            validator_metadata = result.get("agent_4_validation") or {}
            if isinstance(validator_metadata, dict):
                raw_reviews = validator_metadata.get("clause_reviews") or []
                if isinstance(raw_reviews, list):
                    clause_reviews = [r for r in raw_reviews if isinstance(r, dict)]

        statement_order = bs_order if account_type == "BS" else is_order
        statement_display_names = bs_display_names if account_type == "BS" else is_display_names
        order = statement_order.get(mapping_key, statement_order.get(account_key, 9999))
        display_name = statement_display_names.get(mapping_key, account_key)

        sortable_items[account_type].append(
            (
                order,
                mappings[mapping_key].get("category", ""),
                mapping_key,
                {
                    "account_name": account_key,
                    "mapping_key": mapping_key,
                    "display_name": display_name,
                    "display_name_zh": _find_chinese_display_name(mapping_key, display_name, mappings),
                    "category": mappings[mapping_key].get("category", ""),
                    "financial_data": financial_data,
                    "commentary": commentary_text,
                    "clause_reviews": clause_reviews,
                    "summary": _extract_summary(final_content) if final_content else "",
                    # Predominantly-Chinese (>30%), not "contains any CJK
                    # character" -- this flag drives CJK-vs-Latin text
                    # WRAPPING/measurement throughout the packing pipeline
                    # (_calculate_content_lines, _calculate_max_lines_for_
                    # textbox's whole-statement is_chinese_any). An English
                    # commentary that merely names a Chinese counterparty/
                    # person (e.g. "...payable to the related party 维彧")
                    # is still fundamentally Latin-script prose; measuring
                    # it (and, via is_chinese_any, EVERY slot's capacity in
                    # the whole statement) with CJK line-height/spacing/
                    # wrap rules instead of Arial's produced a systematic
                    # believed-vs-actually-rendered fill gap.
                    "is_chinese": contains_predominantly_chinese_text(commentary_text),
                },
            )
        )

    for statement_type in ["BS", "IS"]:
        payloads[statement_type] = [
            item
            for _order, _category, _mapping_key, item in sorted(
                sortable_items[statement_type],
                key=lambda row: (row[0], row[1], row[2]),
            )
        ]

    return payloads
# --- end pptx/payloads.py ---
