from __future__ import annotations

# re-added: bound by an import in another section of the pre-split module
from typing import Any, Dict, List, Optional
from ..financial_common import build_income_statement_period_label, load_required_yaml_file, package_file_path
import re

"""
Prompt loading and rendering helpers for the FDD pipeline.
"""

from .config import normalize_language_code, resolve_agent_alias
from .english import normalize_english_structure, normalize_english_text


import json
import logging
import os
from typing import Any, Dict, Optional, Tuple

import pandas as pd

from ..financial_display_format import add_language_display_columns, prepare_display_dataframe, stringify_display_dataframe
from ..financial_json_converter import df_to_json_str
from ..workbook import build_significant_movements, build_trend_summary, find_mapping_key
from ..financial_common import (
    get_pipeline_result_text,
    load_yaml_file,
    package_file_path,
    visible_descriptions,
)
from ..workbook import INTERNAL_ROW_KEY

_DEFAULT_PROMPTS_FILE = "fdd_utils/prompts.yml"
_DEFAULT_MAPPINGS_FILE = "fdd_utils/mappings.yml"
_DEFAULT_PROMPTS_PATH = package_file_path("prompts.yml")
_DEFAULT_MAPPINGS_PATH = package_file_path("mappings.yml")
_PROMPT_ENGINE_CACHE: Dict[Tuple[str, str], "PromptEngine"] = {}


_STYLE_PACK_CACHE: Dict[str, Dict[str, str]] = {}


def _style_pack_text(key: str, language: str) -> str:
    """Read one PromptStylePack string out of prompts.yml.

    The text used to sit inline in the three methods below. It is prompt data,
    so it lives with the other prompts now; the lookup is cached because
    PromptStylePack is constructed per account.
    """
    if not _STYLE_PACK_CACHE:
        doc = load_required_yaml_file(_DEFAULT_PROMPTS_PATH) or {}
        _STYLE_PACK_CACHE.update(doc.get("style_pack") or {})
    entry = _STYLE_PACK_CACHE.get(key) or {}
    return entry.get("chi" if language == "Chi" else "eng", "")


class PromptStylePack:
    """Small style pack mirroring the HR separation of prompts and style rules."""

    def __init__(self, language: str = "Eng"):
        self.language = normalize_language_code(language)

    def language_instruction(self) -> str:
        return _style_pack_text("language_instruction", self.language)

    def common_formatting_rules(self) -> str:
        return _style_pack_text("common_formatting_rules", self.language)

    def fdd_judgement_rules(self) -> str:
        return _style_pack_text("fdd_judgement_rules", self.language)

    def common_data_rules(self, data_format: str) -> str:
        if data_format == "json":
            if self.language == "Chi":
                return "数据以 JSON 提供，数值和单位已按展示口径预处理，请严格按字段和值直接使用。"
            return (
                "The data is provided as JSON. Values and units are already normalized for reporting, "
                "so use the fields and values exactly as provided."
            )

        if self.language == "Chi":
            return "数据以 Markdown 表格提供，金额已按展示口径格式化，请直接使用，不要再次换算。"
        return (
            "The data is provided as a markdown table. Amounts are already formatted for reporting, "
            "so use them directly without reconversion."
        )


def resolve_prompt_asset_path(path: Optional[str], default_file: str, default_path: str) -> str:
    if not path or path in {default_file, default_path}:
        return default_path
    if os.path.isabs(path):
        return path
    return os.path.join(os.getcwd(), path)


def get_prompt_engine(
    prompts_path: Optional[str] = None,
    mappings_path: Optional[str] = None,
) -> "PromptEngine":
    resolved_prompts = resolve_prompt_asset_path(
        prompts_path,
        _DEFAULT_PROMPTS_FILE,
        _DEFAULT_PROMPTS_PATH,
    )
    resolved_mappings = resolve_prompt_asset_path(
        mappings_path,
        _DEFAULT_MAPPINGS_FILE,
        _DEFAULT_MAPPINGS_PATH,
    )
    cache_key = (resolved_prompts, resolved_mappings)
    if cache_key not in _PROMPT_ENGINE_CACHE:
        _PROMPT_ENGINE_CACHE[cache_key] = PromptEngine(
            prompts_path=resolved_prompts,
            mappings_path=resolved_mappings,
        )
    return _PROMPT_ENGINE_CACHE[cache_key]


class PromptEngine:
    """Centralized prompt accessor aligned with the HR prompt handling pattern."""

    def __init__(
        self,
        prompts_path: Optional[str] = None,
        mappings_path: Optional[str] = None,
    ):
        self.prompts_path = prompts_path or package_file_path("prompts.yml")
        self.mappings_path = mappings_path or package_file_path("mappings.yml")
        self._prompts_data = None
        self._mappings_data = None
        self.logger = logging.getLogger(self.__class__.__name__)

    @staticmethod
    def normalize_agent_name(agent_name: str) -> str:
        return resolve_agent_alias(agent_name)

    @property
    def prompts_data(self) -> Dict[str, Any]:
        if self._prompts_data is None:
            self._prompts_data = load_yaml_file(self.prompts_path)
        return self._prompts_data

    @property
    def mappings_data(self) -> Dict[str, Any]:
        if self._mappings_data is None:
            self._mappings_data = load_yaml_file(self.mappings_path)
        return self._mappings_data

    def get_mapping_component(self, mapping_key: str, component: Optional[str] = None) -> Any:
        resolved_key = find_mapping_key(mapping_key, self.mappings_data)
        if resolved_key:
            data = self.mappings_data.get(resolved_key, {})
            return data.get(component) if component else resolved_key
        return None

    def _patterns_enabled(self) -> bool:
        """Whether mappings.yml example sentences go into the Generator prompt.

        The examples date from a much weaker local model that needed a
        concrete template to imitate. An audit found 57% are fill-in-the-blanks
        templates (see inspect_mapping_patterns.py), which can constrain a
        capable model rather than guide it. Whether they still earn their
        place is an empirical question, so it is switchable
        (processing.inject_mapping_patterns) rather than argued about.
        """
        try:
            from ..financial_common import load_yaml_file
            for candidate in ("fdd_utils/config.yml", "fdd_utils/config.example.yml"):
                cfg = load_yaml_file(candidate)
                if cfg:
                    value = (cfg.get("processing") or {}).get("inject_mapping_patterns")
                    return True if value is None else bool(value)
        except Exception:
            pass
        return True

    def resolve_mapping_key(self, mapping_key: str) -> str:
        return self.get_mapping_component(mapping_key) or mapping_key

    def _fallback_mapping_section(self, mapping_key: str) -> Optional[str]:
        if mapping_key in self.mappings_data:
            return None
        if "_general_dynamic_mapping" in self.mappings_data:
            return "_general_dynamic_mapping"
        return None

    def get_agent_defaults(self, agent_name: str, language: str) -> Tuple[str, str]:
        language = normalize_language_code(language)
        agent_key = self.normalize_agent_name(agent_name)
        if agent_key == "1_Generator":
            generic_prompts = self.mappings_data.get("_default_subagent_1", {}).get(language, {})
            return generic_prompts.get("system_prompt", ""), generic_prompts.get("user_prompt", "")

        prompt_data = self.prompts_data.get(agent_key, {}).get(language, {})
        return prompt_data.get("system_prompt", ""), prompt_data.get("user_prompt", "")

    def get_prompt_pair(self, agent_name: str, language: str, mapping_key: str) -> Tuple[str, str]:
        language = normalize_language_code(language)
        agent_key = self.normalize_agent_name(agent_name)
        resolved_mapping_key = self.resolve_mapping_key(mapping_key)

        if agent_key == "1_Generator":
            default_system_prompt, default_user_prompt = self.get_agent_defaults(agent_name, language)

            account_data = self.mappings_data.get(resolved_mapping_key, {})
            if not account_data:
                fallback_section = self._fallback_mapping_section(mapping_key)
                if fallback_section:
                    account_data = self.mappings_data.get(fallback_section, {})
            raw_sp = account_data.get("subagent_1_prompts") or {}
            account_prompts = (raw_sp if isinstance(raw_sp, dict) else {}).get(language, {})
            system_prompt = (account_prompts.get("system_prompt") or "").strip() or default_system_prompt
            user_prompt = (account_prompts.get("user_prompt") or "").strip() or default_user_prompt
            return system_prompt, user_prompt

        return self.get_agent_defaults(agent_name, language)

    def _build_markdown_prompt_payload(self, df: pd.DataFrame) -> Dict[str, str]:
        df_display = prepare_display_dataframe(
            df,
            drop_columns=(INTERNAL_ROW_KEY,),
        )
        rendered = stringify_display_dataframe(df_display).to_markdown(index=False).strip()
        return {"financial_figure": rendered, "financial_data": rendered}

    @staticmethod
    def _normalize_prompt_dataframe(df: Optional[pd.DataFrame], language: str) -> Optional[pd.DataFrame]:
        if language != "Eng" or not isinstance(df, pd.DataFrame) or df.empty:
            return df
        normalized_df = df.copy()
        normalized_df.columns = [
            column if str(column) == INTERNAL_ROW_KEY else normalize_english_text(str(column))
            for column in normalized_df.columns
        ]
        for column in normalized_df.columns:
            if str(column) == INTERNAL_ROW_KEY:
                continue
            series = normalized_df[column]
            if pd.api.types.is_numeric_dtype(series):
                continue
            normalized_df[column] = series.apply(
                lambda value: normalize_english_text(value) if isinstance(value, str) else value
            )
        normalized_df.attrs.update(df.attrs)
        return normalized_df

    @staticmethod
    def _normalize_prompt_value(value: Any, language: str) -> Any:
        if language != "Eng":
            return value
        return normalize_english_structure(value)

    def _filter_prompt_analysis_df(self, df: pd.DataFrame) -> Optional[pd.DataFrame]:
        analysis_df = df.attrs.get("prompt_analysis_df")
        if not isinstance(analysis_df, pd.DataFrame) or analysis_df.empty:
            return None
        visible_rows = visible_descriptions(df)
        if not visible_rows or len(analysis_df.columns) == 0:
            return analysis_df
        # Component ("breakdown") rows are deliberately absent from the main
        # frame -- they are not accounts of their own -- so a visibility test
        # alone would drop them straight back out. They are what lets a bullet
        # say "包括：1）...2）..." instead of quoting one total, so keep them.
        components = set(analysis_df.attrs.get("component_descriptions") or [])
        first_col = analysis_df.columns[0]
        filtered = analysis_df[
            analysis_df[first_col].astype(str).map(
                lambda value: value.strip() in visible_rows or value.strip() in components
            )
        ].copy()
        filtered.attrs["component_descriptions"] = list(components)
        return filtered if not filtered.empty else analysis_df

    def _filter_adjacent_detail_rows(self, df: pd.DataFrame) -> list[Dict[str, Any]]:
        adjacent_detail_rows = df.attrs.get("adjacent_detail_rows") or []
        if not adjacent_detail_rows:
            return []
        visible_rows = visible_descriptions(df)
        if not visible_rows:
            return adjacent_detail_rows
        filtered = [
            row for row in adjacent_detail_rows
            if str(row.get("Description", "")).strip() in visible_rows
        ]
        return filtered or adjacent_detail_rows

    @staticmethod
    def _should_skip_prompt_metadata_key(key_text: str, include_description: bool = False) -> bool:
        return (
            key_text == INTERNAL_ROW_KEY
            or (include_description and key_text == "Description")
            or key_text.endswith("| table_header")
            or key_text.endswith("| indicative_adjusted_row")
            or key_text.endswith("| date_row")
        )

    @staticmethod
    def _prompt_ready_adjacent_detail_rows(
        adjacent_detail_rows: list[Dict[str, Any]], language: str = "Eng",
    ) -> list[Dict[str, Any]]:
        """Strip prompt-metadata keys and reformat any bare numeric remark
        value (万/亿/million per format_value_by_language) before this reaches
        the AI. Unlike the main financial_data table, adjacent-detail /
        RHS-remark columns are raw cell text straight from the workbook --
        a plain check/reconciliation figure like -980981 flowed through
        untouched, so the model wrote it verbatim ("负人民币980,981元"
        instead of "-98万元") since only the main table is covered by the
        "already formatted, do not reconvert" instruction. Skips
        "Description" (the row label) and anything that isn't a clean
        numeric string (dates, percentages, free-text notes stay as-is).
        """
        from ..financial_display_format import format_value_by_language
        from ..financial_common import coerce_numeric

        cleaned_rows: list[Dict[str, Any]] = []
        for row in adjacent_detail_rows:
            if not isinstance(row, dict):
                continue
            cleaned_row: Dict[str, Any] = {}
            for key, value in row.items():
                key_text = str(key)
                if PromptEngine._should_skip_prompt_metadata_key(key_text):
                    continue
                if key_text != "Description" and isinstance(value, str) and value.strip():
                    text = value.strip()
                    if re.fullmatch(r"-?[\d,，]+(\.\d+)?", text):
                        numeric = coerce_numeric(text)
                        if numeric is not None:
                            value = format_value_by_language(numeric, language)
                cleaned_row[key_text] = value
            if cleaned_row:
                cleaned_rows.append(cleaned_row)
        return cleaned_rows

    @staticmethod
    def _table_linked_remarks(df: Optional[pd.DataFrame]) -> list[Dict[str, Any]]:
        if not isinstance(df, pd.DataFrame):
            return []
        remarks = df.attrs.get("table_linked_remarks") or []
        return [remark for remark in remarks if isinstance(remark, dict)]

    def _build_analysis_prompt_df(self, df: pd.DataFrame) -> Optional[pd.DataFrame]:
        analysis_df = self._filter_prompt_analysis_df(df)
        if not isinstance(analysis_df, pd.DataFrame) or analysis_df.empty:
            return analysis_df
        effective_analysis_df = analysis_df.copy()
        if INTERNAL_ROW_KEY in effective_analysis_df.columns:
            effective_analysis_df = effective_analysis_df.drop(columns=[INTERNAL_ROW_KEY])

        return effective_analysis_df

    @staticmethod
    def _format_analysis_prompt_df(df: Optional[pd.DataFrame], report_language: str) -> Optional[pd.DataFrame]:
        if not isinstance(df, pd.DataFrame) or df.empty or not report_language:
            return df
        formatted_analysis_df = add_language_display_columns(df.copy(), report_language)
        formatted_analysis_df.attrs.update(df.attrs)
        return formatted_analysis_df

    @staticmethod
    def _composition_guidance(df: Optional[pd.DataFrame], language: str) -> str:
        """Tell the model how to enumerate, but only when there is something to
        enumerate.

        The component rows of a schedule now reach the prompt. Without this the
        model names the categories and drops their amounts -- "主要包括租金收入、
        物业管理费收入及水电费收入" where the analyst deliverable writes "1）...
        598.1万元；2）...39.2万元". It was reading the detail (it quoted a single
        3,588元 tenant figure unprompted); it just had no instruction on the
        shape.

        Gated on components actually being present. Asking for an itemised
        breakdown of an account that has none is an invitation to invent one.
        """
        if not isinstance(df, pd.DataFrame):
            return ""
        analysis_df = df.attrs.get("prompt_analysis_df")
        components = []
        if isinstance(analysis_df, pd.DataFrame):
            components = list(analysis_df.attrs.get("component_descriptions") or [])
        if len(components) < 2:
            return ""
        if language == "Chi":
            return (
                "【组成披露】该科目的明细组成已随财务数据提供。"
                "请按组成列举，且每一项都必须带上金额——只写类别名称而不给金额是不合格的。"
                "在**有金额的最高层级**列举（例如租金收入、物业管理费收入、水电费收入各自的余额），"
                "不要逐一罗列每个交易对手；仅在某一交易对手金额重大、账龄异常或性质特殊时才具名，"
                "并说明为何值得一提。组成达三项或以上时使用'包括：1）…；2）…；3）…'的编号句式。"
                "**列举前先做加法**：把你打算列出的各项金额相加，与合计核对。若两者不等，差额本身就是一个必须交代的组成项——用数据中对应的名称说明它是什么（例如截止性调整、补计提、未解释性质的余额），并把它作为最后一项列出，写成'其余X万元为…'。只列出加总不等于合计的几项而不交代差额，是不完整的披露。不要强行凑数。"
            )
        return (
            "COMPOSITION. The account's component lines are supplied with the financial data. "
            "Enumerate the composition and give an AMOUNT for every item -- naming categories without "
            "amounts is not acceptable. Enumerate at the HIGHEST level that carries amounts (e.g. the "
            "balance of each revenue stream), not every counterparty; name an individual counterparty "
            "only where its size, ageing or nature makes it worth calling out, and say why. Use a "
            "numbered form -- \"comprising: 1) ...; 2) ...; 3) ...\" -- once there are three or more "
            "components. ADD THE ITEMS UP and check them against the total BEFORE writing them. Any "
            "difference is itself a component you have to account for: name it from the data (a "
            "cut-off adjustment, an accrual, an unexplained balance) and list it last as \"the "
            "remaining CNYx\". Listing items that do not reach the total, without saying what the "
            "rest is, is an incomplete disclosure. Do not force a reconciliation."
        )

    def _rhs_guidance_block(self, adjacent_detail_rows: list[Dict[str, Any]], language: str) -> str:
        if not adjacent_detail_rows:
            return ""
        if language == "Chi":
            return (
                "表格右侧1-5列的补充备注/工作说明已在财务数据载荷末尾提供。"
                "仅在其与数据一致时将其视为可引用的补充背景或原因。"
            )
        return (
            "Supplemental context extracted from the 1-5 side columns of the schedule is included in the financial data payload. "
            "Use it only where it is supported by the evidence, and absorb it into normal report sentences rather than repeating section labels."
        )

    @staticmethod
    def _summarize_rhs_remarks(adjacent_detail_rows: list[Dict[str, Any]], language: str) -> list[str]:
        if not adjacent_detail_rows:
            return []

        summaries: list[str] = []
        seen: set[str] = set()
        for row in adjacent_detail_rows:
            if not isinstance(row, dict):
                continue
            description = str(row.get("Description") or "").strip()
            remark_parts: list[str] = []
            for key, value in row.items():
                key_text = str(key)
                if PromptEngine._should_skip_prompt_metadata_key(key_text, include_description=True):
                    continue
                if isinstance(value, (int, float)):
                    continue
                text = str(value or "").strip()
                if not text:
                    continue
                if key_text == description:
                    continue
                remark_parts.append(text)

            unique_parts: list[str] = []
            seen_parts: set[str] = set()
            for part in remark_parts:
                if part in seen_parts:
                    continue
                seen_parts.add(part)
                unique_parts.append(part)

            if not unique_parts:
                continue
            if language == "Chi":
                summary = f"{description}: " + "；".join(unique_parts[:2]) if description else "；".join(unique_parts[:2])
            else:
                summary = f"{description}: " + "; ".join(unique_parts[:2]) if description else "; ".join(unique_parts[:2])
            if summary not in seen:
                seen.add(summary)
                summaries.append(summary)
        return summaries[:5]

    @staticmethod
    def _remarks_weight_instruction(
        *,
        has_rhs_remarks: bool,
        has_supporting_notes: bool,
        has_user_comment: bool,
        statement_type: str,
        language: str,
    ) -> str:
        if not any((has_rhs_remarks, has_supporting_notes, has_user_comment)):
            return ""
        normalized_statement_type = str(statement_type or "").strip().upper()
        if language == "Chi":
            if normalized_statement_type == "BS":
                return (
                    "若补充备注、右侧备注或用户备注提供了与数据一致的背景、原因、限制、差异解释或工作说明，"
                    "对于资产负债表科目，应优先以这些备注作为原因判断和定性解释的主要依据，数值则主要用于支持趋势、余额与重要性判断。"
                    "若备注中明确写明折旧/摊销方法、使用年限、残值率、坏账计提依据等会计政策细节，且与当前科目及数据一致，可在评论中吸收概括。"
                    "财务尽调评论应在不脱离数据的前提下，吸收这些备注并概括成简洁、顾问式的说明，而不是忽略它们。"
                )
            if normalized_statement_type == "IS":
                return (
                    "若补充备注、右侧备注或用户备注提供了与数据一致的背景、原因、限制、差异解释或工作说明，"
                    "对于利润表科目，应将数值趋势/重大变动分析与这些备注结合起来共同解释当期表现，而不是只依赖单一来源。"
                    "若备注中明确写明与当前科目相关的会计政策、成本构成、收入确认、折旧/摊销、坏账计提或其他解释性细节，且与数据一致，可在评论中吸收概括。"
                    "财务尽调评论应在不脱离数据的前提下，吸收这些备注并概括成简洁、顾问式的说明，而不是忽略它们。"
                )
            return (
                "若补充备注、右侧备注或用户备注提供了与数据一致的背景、原因、限制、差异解释或工作说明，请提高其在评论中的权重。"
                "若备注中明确写明折旧/摊销方法、使用年限、残值率、坏账计提依据等会计政策细节，且与当前科目及数据一致，可在评论中吸收概括。"
                "财务尽调评论应在不脱离数据的前提下，吸收这些备注并概括成简洁、顾问式的说明，而不是忽略它们。"
            )
        if normalized_statement_type == "BS":
            return (
                "Where supporting notes, supplemental side-column context, or user remarks provide data-consistent context such as reasons, restrictions, differences, "
                "or working explanations, treat those remarks as a required part of the reasoning on balance-sheet items. Build the reasoning from the combination of data, "
                "cross-period trend, and supported remarks, while using the remarks as the primary basis for qualitative explanation. Use the numbers mainly to confirm the trend, "
                "latest balance, composition, and materiality rather than inventing causes from the figures alone. If remarks explicitly state account-relevant accounting-policy "
                "details such as depreciation or amortization method, useful life, residual value, or provisioning basis, you may reflect them when they are clearly supported. "
                "As an FDD consultant, absorb and summarize those points into report-ready sentences rather than copying labels, dumping note headers, or leaving them unused. "
                "If the remarks contain material supported explanations, differences, restrictions, or drivers, make them visible in the commentary rather than treating them as optional background."
            )
        if normalized_statement_type == "IS":
            return (
                "Where supporting notes, supplemental side-column context, or user remarks provide data-consistent context such as reasons, restrictions, differences, or working "
                "explanations, treat them as a required part of the reasoning on income-statement items. Build the reasoning from the combination of data, cross-period trend, remarks, "
                "and broader numeric analysis such as movement, scale, mix, and reasonableness. Use both the figures and the remarks together to explain the period performance, rather "
                "than relying on only one source. If remarks explicitly state account-relevant accounting-policy details such as depreciation or amortization "
                "method, useful life, residual value, provisioning basis, or other explanation for the period movement, you may reflect them when they are clearly supported. As an FDD "
                "consultant, absorb and summarize those points into report-ready sentences rather than copying labels, dumping note headers, or leaving them unused. "
                "If the remarks contain material supported explanations, drivers, or validation-relevant observations, make them visible in the commentary rather than treating them as optional background."
            )
        return (
            "Where supporting notes, supplemental side-column context, or user remarks provide data-consistent context such as reasons, restrictions, differences, "
            "or working explanations, give them higher weight in the commentary and treat them as an active part of the reasoning rather than optional background. "
            "Build the commentary from data, trend, and supported remarks together. If remarks explicitly state account-relevant accounting-policy details such as "
            "depreciation or amortization method, useful life, residual value, or provisioning basis, you may reflect them in the commentary when they are clearly supported. "
            "As an FDD consultant, absorb and summarize those points "
            "into report-ready sentences rather than copying labels, dumping note headers, or leaving them unused. If the remarks contain material supported observations, "
            "make them visible in the commentary rather than treating them as optional background."
        )

    @staticmethod
    def _user_guidance_instruction(user_comment: str, language: str) -> str:
        guidance = str(user_comment or "").strip()
        if not guidance:
            return ""
        if language == "Chi":
            return (
                "用户备注/重提示指引已提供。若其与数据、备注或表格内容一致，请将其视为明确写作指引，"
                "并在输出的措辞、重点、结构或补充说明中体现。若其中有与数据不一致之处，不要照搬，应仅保留有依据的部分。"
            )
        return (
            "User remarks / reprompt guidance has been provided. Where it is consistent with the data, notes, and table context, "
            "treat it as explicit writing guidance and reflect it in the wording, emphasis, structure, or clarifications of the output. "
            "Do not follow any part that is not supported by the provided evidence."
        )

    @staticmethod
    def _detail_table_guidance(df: Optional[pd.DataFrame], language: str) -> str:
        """The account's report-ready breakdown, named component by component.

        Where a sheet carries one (see workbook.py's
        extract_presentation_detail_table), the main schedule the model
        otherwise works from is keyed by GL codes -- 660202, 660203 -- so it
        can only write "660202 increased", never "会计服务费". This hands over
        the human-readable component names and their figures.

        It also states the convention the real deliverable follows for these
        accounts: one or two sentences naming the composition, then hand off
        to the table rather than reciting every line in prose. Measured on a
        real deck, 18% of its paragraphs do that and 0% of ours did.
        """
        if not isinstance(df, pd.DataFrame):
            return ""
        table = (df.attrs or {}).get("presentation_detail_table")
        if not table or not table.get("rows"):
            return ""
        rows = table["rows"]
        periods = table.get("periods") or []
        latest = periods[-1] if periods else None
        # table's values are in the SAME raw-yuan internal scale every
        # account's df uses (source_multiplier -- 1000 whenever the sheet's
        # own header says 千元 -- applied so cross-account math and the
        # table's own tie-out against the account total work); divide back
        # out for this prompt-facing figure the same way pptx.py's renderer
        # does, so a model that ignores the "don't repeat figures" ban below
        # isn't shown a number ~1000x the real one.
        divisor = (df.attrs or {}).get("source_multiplier") or 1

        def _fmt(entry):
            values = entry.get("values") or {}
            if latest and latest in values and isinstance(values[latest], (int, float)):
                return f"{entry['label']} {values[latest] / divisor:,.1f}"
            return str(entry["label"])

        listed = "、".join(_fmt(r) for r in rows[:10])
        names_only = "、".join(str(r["label"]) for r in rows[:10])

        # The illustrative sentence should name the LARGEST component, not
        # whichever happens to be listed first -- on real data the first row
        # was zero in the latest period, which would model a sentence leading
        # on a component that did not occur.
        def _latest_abs(entry):
            values = entry.get("values") or {}
            v = values.get(latest) if latest else None
            return abs(v) if isinstance(v, (int, float)) else 0.0

        # A catch-all bucket can easily be the largest line, but leading a
        # sentence on "Other" models the opposite of naming a composition, so
        # prefer the largest named component and only fall back if all are
        # generic.
        _generic = ("其他", "其它", "other", "others", "misc")
        named = [r for r in rows
                 if not any(g in str(r["label"]).lower() for g in _generic)]
        lead = max(named or rows, key=_latest_abs)["label"] if rows else ""

        if language == "Chi":
            return (
                f"【本科目已有做好的明细表】构成项（按最新一期{latest or ''}）："
                f"{listed}。"
                "撰写时**必须使用这些构成项名称**，不得使用会计科目代码（如660202）或笼统的"
                "'其他明细'来代替。"
                "本科目在真实交付稿中的写法分两段，中间用换行分开：\n"
                "**第一段（必写，硬性上限约60字，真实交付稿量出来是40-80字）**：一到两句点出主要构成"
                f"（例如'主要包括{lead}等'），以'明细如下：'收尾——这一段只负责点名构成项，"
                "不得在这一段提供任服务方、金额或计费方式，那些留给第二段和表格。\n"
                "**第二段（可选——只在下方备注信息里真的能找到对应说明时才写，找不到就整段省略，"
                "禁止编造）**：紧接在'明细如下：'之后另起一行，针对**最多1-2个**真正有实质说明可写的构成项"
                "（服务提供方是谁、依什么基准计费、合同条款为何——不是每个构成项都需要，通常只有"
                "一两个有这类背景信息；表格本身已经列出所有构成项的金额，这一段只补金额以外的事实，"
                "宁可少写、只写最值得写的1-2项，不要为了凑数把每个构成项都写一条），"
                "各写一条以'➢ [构成项名称]：'开头的说明。"
                "每条不超过约80字，第二段全部加起来不超过约180字——真实交付稿这一段通常只有1-2句"
                "简短事实（如'按月计提，已计提至X月'），不是逐项展开的长段落。\n"
                "**明确禁止（两段都适用）**：不得罗列各期总额（如'2023年度、2024年度、2025年度分别为"
                "A、B、C'）；不得罗列各构成项的每期金额；不得写年化换算数。这些全部由表格承担——"
                "第二段是解释'是什么、谁提供、怎么算'，不是重复表格里已经有的数字。"
            )
        return (
            f"[A READY-MADE BREAKDOWN EXISTS FOR THIS ACCOUNT] Components (latest period "
            f"{latest or ''}): {listed}. "
            "You MUST use these component names -- never an account code (660202) or a vague "
            "'other details'. The real deliverable handles this account in TWO parts, separated "
            "by a line break:\n"
            "**Part 1 (required, hard limit ~60 words, real deliverables run 40-80)**: one or "
            f"two sentences naming the main components (e.g. 'mainly comprised {lead} and "
            "others'), handing off with 'the breakdown is set out below'. This part ONLY names "
            "components -- no provider, no figures, no charging basis; those belong to part 2 "
            "and the table.\n"
            "**Part 2 (optional -- write it ONLY if the remarks below genuinely support it for "
            "a given component; omit the whole part rather than invent anything)**: starting on "
            "a new line right after the part-1 handoff, one bullet for **at most 1-2 components** "
            "(not every component needs one -- usually only one or two have this kind of "
            "background; the table already lists every component's figures, this part only adds "
            "facts beyond the numbers -- prefer writing less and picking only the most worthwhile "
            "1-2 components over padding one bullet per component) that genuinely has supporting "
            "detail, each starting with '- [component name]: ', covering who provides it, the "
            "charging basis, or contract terms. Keep each under about 40 words, and the whole of "
            "part 2 under about 90 words -- the real deliverable's own version of this part is "
            "usually 1-2 short factual clauses (e.g. 'accrued monthly, through March'), not an "
            "expanded paragraph per component.\n"
            "**Explicitly forbidden in BOTH parts**: listing period totals ('FY23, FY24 and FY25 "
            "were A, B and C'), listing each component's per-period figures, and annualised "
            "restatements -- the table carries all of those. Part 2 explains what/who/how, not "
            "numbers the table already shows."
        )

    @staticmethod
    def _variance_analysis_guidance(
        df: Optional[pd.DataFrame],
        language: str,
        peer_context: Optional[Dict[str, Any]] = None,
        mapping_key: Optional[str] = None,
        thresholds: Optional[Dict[str, Any]] = None,
    ) -> str:
        """Guidance for an account whose own figures moved materially, an
        IS line that moved out of proportion to revenue EVEN WITHOUT a
        large movement of its own, and earlier period-over-period
        movements the latest-pair check alone would miss.

        Project-team asks this implements: explain WHY a large income-
        statement movement happened (reasoning from remarks where they
        exist, flagged as such); call out an expense line growing out of
        proportion to revenue -- including a FLAT cost against
        collapsing/surging revenue, which is itself the disproportion (this
        used to be gated behind the account's own >=30% movement, which
        made that exact case -- e.g. flat G&A against collapsing revenue --
        impossible to ever flag); mention material balance-sheet swings;
        and surface earlier movements (e.g. 2023->2024) a check that only
        ever looks at the latest two periods would silently miss, without
        letting the prompt balloon -- capped at two one-line mentions.

        Every movement is computed HERE, deterministically, and handed to
        the model as a stated fact -- rather than asking it to eyeball the
        table and derive a percentage itself, which would be both
        unreliable and ungrounded (a self-derived figure has no source to
        match against, so the hallucination check would flag it).

        Partial tail periods are excluded from every comparison: a
        one-month period against a full year reads as a ~90% collapse that
        is an artefact of period length, not a real movement.

        thresholds (optional, from config.yml's analysis: block --
        variance_threshold_pct/disproportion_gap_pp/max_extra_pairs) fall
        back to this function's own long-standing defaults (30/30/2) when
        absent, so an unconfigured deployment behaves exactly as before
        this parameter existed.
        """
        if not isinstance(df, pd.DataFrame) or df.empty:
            return ""
        thresholds = thresholds or {}
        variance_threshold_pct = float(thresholds.get("variance_threshold_pct", 30) or 30)
        disproportion_gap_pp = float(thresholds.get("disproportion_gap_pp", 30) or 30)
        max_extra_pairs = int(thresholds.get("max_extra_pairs", 2) or 2)
        attrs = df.attrs or {}
        integrity = attrs.get("integrity") or {}
        statement_type = str(integrity.get("statement_type") or "").strip().upper()
        if statement_type not in ("BS", "IS"):
            return ""

        internal_key = "__source_row_idx"
        period_cols = [
            c for c in list(df.columns)[1:]
            if str(c) != internal_key and not str(c).endswith("_formatted")
        ]
        if len(period_cols) < 2:
            return ""

        row_types = attrs.get("row_types_by_description") or {}
        desc_col = df.columns[0]
        total_idx = None
        for idx, row in df.iterrows():
            if str(row_types.get(str(row[desc_col]), "")).lower() in ("total", "subtotal"):
                total_idx = idx
        try:
            if total_idx is None:
                series = [(str(c), float(df[c].fillna(0).sum())) for c in period_cols]
            else:
                series = [(str(c), float(df.loc[total_idx, c] or 0)) for c in period_cols]
        except Exception:
            return ""

        months = attrs.get("annualization_months") or integrity.get("annualization_months")
        if isinstance(months, (int, float)) and 0 < months < 12 and len(series) > 2:
            series = series[:-1]
        if len(series) < 2:
            return ""

        scale = max((abs(v) for _p, v in series), default=0.0)
        if scale <= 0:
            return ""

        def _pair(prev_v: float, curr_v: float) -> Tuple[Optional[float], bool]:
            """(pct, is_flip). pct is None when the base is too small for a
            percentage to mean anything OR the value crosses zero (a raw %
            across a sign change is nonsense -- e.g. -2,049 to +4,844 is not
            "a 336% increase"); is_flip marks that second case specifically
            so callers can describe it qualitatively instead of with a
            percentage."""
            if abs(prev_v) < scale * 0.01:
                return None, False
            if prev_v * curr_v < 0:
                return None, True
            return (curr_v - prev_v) / abs(prev_v) * 100, False

        (p_prev, v_prev), (p_curr, v_curr) = series[-2], series[-1]
        pct, flipped = _pair(v_prev, v_curr)
        is_material = pct is not None and abs(pct) >= variance_threshold_pct

        notes = attrs.get("supporting_notes") or []
        rhs = attrs.get("adjacent_detail_rows") or []
        has_material_support = bool(notes or rhs)
        # A latest period of zero is the case the opening-sentence rule keeps
        # losing: a real run opened 预付款项 with "截至2025年12月31日，余额为17万"
        # and only mentioned the (zero) latest period afterwards, even after a
        # feedback retry. Stating the required opening explicitly, with the
        # computed dates, is more reliable than restating the rule.
        zero_latest_chi = zero_latest_eng = ""
        if abs(v_curr) < 1e-9 and abs(v_prev) > 0:
            zero_latest_chi = (
                f"【首句强制】最新一期（{p_curr}）余额为零。首句主体必须是{p_curr}且写明无余额"
                f"（如'截至{p_curr}，该科目无余额'），随后才补充{p_prev}的余额及其构成。"
                f"绝不可用{p_prev}的非零余额作为首句主体——那会让读者误以为那是当期数据。"
            )
            zero_latest_eng = (
                f"[REQUIRED OPENING] The latest period ({p_curr}) is nil. The first sentence must "
                f"be about {p_curr} and say there is no balance, with the {p_prev} balance and its "
                f"composition following afterwards. Never open on the non-nil {p_prev} figure -- a "
                "reader will take it for the current period."
            )

        # Revenue-disproportion is checked INDEPENDENTLY of whether the
        # account's OWN movement is itself large -- a fixed/flat cost that
        # simply doesn't track a collapsing or surging revenue line IS the
        # disproportion the project team asks about; gating it behind the
        # account's own >=30% move meant that exact case could never fire.
        # Skipped for the revenue account comparing against itself, and
        # when the account's own pct couldn't be computed (tiny base or a
        # sign flip -- the flip already gets its own treatment below).
        peer_line_chi = peer_line_eng = ""
        revenue_key = (peer_context or {}).get("revenue_key") if peer_context else None
        is_revenue_account = bool(revenue_key) and bool(mapping_key) and revenue_key == mapping_key
        if (
            pct is not None and not is_revenue_account and statement_type == "IS"
            and peer_context and peer_context.get("revenue_growth_pct") is not None
        ):
            rev_pct = float(peer_context["revenue_growth_pct"])
            gap = pct - rev_pct
            if abs(gap) >= disproportion_gap_pp:
                peer_line_chi = (
                    f"同期营业收入变动为{rev_pct:+.0f}%，本科目为{pct:+.0f}%，两者相差{gap:+.0f}个百分点，"
                    "属于费用与收入变动不成比例的情形，请在评论中明确指出这一不对称，并在有依据时说明原因。"
                )
                peer_line_eng = (
                    f"Revenue moved {rev_pct:+.0f}% over the same period against this account's "
                    f"{pct:+.0f}%, a gap of {gap:+.0f} percentage points. State that disproportion "
                    "explicitly and, where supported, why it arose."
                )

        # Earlier adjacent full-period pairs the latest-pair check alone
        # would miss (e.g. 2023->2024 when the latest pair is 2025->1Q26).
        # Capped at max_extra_pairs -- these are one-line mentions, not
        # full explanations; the latest pair (and the disproportion line
        # above) carry the real weight.
        earlier_lines_chi: List[str] = []
        earlier_lines_eng: List[str] = []
        for i in range(len(series) - 2):
            if len(earlier_lines_chi) >= max_extra_pairs:
                break
            e_prev_p, e_prev_v = series[i]
            e_curr_p, e_curr_v = series[i + 1]
            e_pct, e_flipped = _pair(e_prev_v, e_curr_v)
            if e_flipped:
                earlier_lines_chi.append(
                    f"（另外，{e_prev_p}至{e_curr_p}期间本科目由{e_prev_v:,.0f}转为{e_curr_v:,.0f}，"
                    "正负性质发生转变，如有支持依据可一并简述，不得臆造原因。）"
                )
                earlier_lines_eng.append(
                    f"(Separately, between {e_prev_p} and {e_curr_p} this account moved from "
                    f"{e_prev_v:,.0f} to {e_curr_v:,.0f}, crossing from one sign to the other -- "
                    "note this shift briefly if the data supports why, do not invent a cause.)"
                )
            elif e_pct is not None and abs(e_pct) >= variance_threshold_pct:
                e_dir_chi = "增长" if e_pct > 0 else "下降"
                e_dir_eng = "increased" if e_pct > 0 else "decreased"
                earlier_lines_chi.append(
                    f"（另外，{e_prev_p}至{e_curr_p}期间本科目{e_dir_chi}约{abs(e_pct):.0f}%，"
                    "如有支持依据可一并简述，无需展开为独立段落。）"
                )
                earlier_lines_eng.append(
                    f"(Separately, this account {e_dir_eng} about {abs(e_pct):.0f}% between "
                    f"{e_prev_p} and {e_curr_p} -- mention briefly if supported, no need for a "
                    "separate paragraph.)"
                )

        # Self-carrying cap exception: the Auditor/Refiner/Validator length
        # caps are enforced in prompts.yml, a completely separate file from
        # this one -- stating the exception HERE too, inside the guidance
        # text itself, means it travels with the instruction that needs it
        # and can't go stale if prompts.yml's own wording drifts later.
        cap_note_chi = "（此说明的1-2句不计入本科目篇幅上限，仅限于此说明本身。）"
        cap_note_eng = "(This explanation's 1-2 sentences are exempt from this account's length cap -- for this explanation only.)"

        if language == "Chi":
            parts = [zero_latest_chi]
            if is_material:
                head = (
                    f"【重大变动提示】{cap_note_chi}本科目合计由{p_prev}的{v_prev:,.0f}变动至{p_curr}的{v_curr:,.0f}，"
                    f"{'增长' if pct > 0 else '下降'}约{abs(pct):.0f}%。此变动幅度重大，不可只陈述金额而不作说明。"
                )
                if has_material_support:
                    body = (
                        "请结合备注/右侧说明推断并说明变动原因：可以在备注所述事实的基础上作合理推理"
                        "（例如备注说明某项目于某期完工转固，则可据此解释折旧上升），"
                        "但推理必须以备注或数据为起点，不得凭空假设市场、竞争、宏观等外部原因。"
                        "凡属推断而非备注原文直述的部分，须用'预计'、'推测'、'主要系...所致'等措辞明确标示其为判断，"
                        "使读者能区分事实与分析。"
                    )
                elif statement_type == "IS":
                    body = (
                        "备注中没有可解释此变动的信息。此时请如实说明变动幅度与方向，"
                        "并指出未取得进一步解释（如'变动原因尚待与管理层确认'），"
                        "不得臆造原因。"
                    )
                else:
                    body = (
                        "备注中没有可解释此变动的信息。请简要点出该变动的规模与方向即可，不得臆造原因。"
                    )
                parts += [head, body]
            elif flipped:
                parts.append(
                    f"【重大变动提示】{cap_note_chi}本科目由{p_prev}的{v_prev:,.0f}转为{p_curr}的{v_curr:,.0f}，"
                    "正负性质发生转变（例如由净收益转为净支出，或相反），此变动不可只陈述金额而不作说明；"
                    "如备注/右侧说明有支持依据可据此推理并标示为判断，否则如实说明性质转变但不得臆造原因。"
                )
            if peer_line_chi:
                parts.append(
                    peer_line_chi if (is_material or flipped)
                    else "【费用与收入变动不成比例提示】" + cap_note_chi + peer_line_chi
                )
            parts.extend(earlier_lines_chi)
            return " ".join(part for part in parts if part)

        parts = [zero_latest_eng]
        if is_material:
            head = (
                f"[MATERIAL MOVEMENT] {cap_note_eng} This account's total moved from {v_prev:,.0f} at {p_prev} to "
                f"{v_curr:,.0f} at {p_curr}, an {'increase' if pct > 0 else 'decrease'} of about "
                f"{abs(pct):.0f}%. A movement this size must be addressed, not merely stated as a figure."
            )
            if has_material_support:
                body = (
                    "Use the notes / side-column remarks to explain it. You may reason from what the "
                    "remarks state (e.g. if a remark says a phase was completed and transferred to fixed "
                    "assets in a period, that can explain a rise in depreciation), but the reasoning must "
                    "start from the remarks or the data -- never assume market, competitive or macro causes. "
                    "Where a point is your inference rather than something the remarks state outright, mark "
                    "it as judgement ('mainly attributable to...', 'expected to...') so the reader can tell "
                    "analysis from fact."
                )
            elif statement_type == "IS":
                body = (
                    "The notes contain nothing that explains this movement. State its size and direction "
                    "accurately and note that no further explanation has been obtained (e.g. 'the driver "
                    "remains to be confirmed with management'). Do not invent a cause."
                )
            else:
                body = (
                    "The notes contain nothing that explains this movement. Note its size and direction "
                    "briefly. Do not invent a cause."
                )
            parts += [head, body]
        elif flipped:
            parts.append(
                f"[MATERIAL MOVEMENT] {cap_note_eng} This account moved from {v_prev:,.0f} at {p_prev} to {v_curr:,.0f} "
                f"at {p_curr}, crossing from one sign to the other (e.g. a net gain turning into a net "
                "cost, or vice versa) -- this must be addressed, not merely stated as a figure. Reason "
                "from the notes/remarks if they support it, marking inference as judgement; otherwise "
                "state the shift accurately without inventing a cause."
            )
        if peer_line_eng:
            parts.append(
                peer_line_eng if (is_material or flipped)
                else "[DISPROPORTIONATE TO REVENUE] " + cap_note_eng + " " + peer_line_eng
            )
        parts.extend(earlier_lines_eng)
        return " ".join(part for part in parts if part)

    @staticmethod
    def _analytical_lens_guidance(df: Optional[pd.DataFrame], language: str) -> str:
        """A short, static set of interpretive frames -- the analytical
        vocabulary the project's own real deliverables actually use, rather
        than a generic "analyse this deeply" instruction. Unlike
        _variance_analysis_guidance (which computes a real movement from
        this account's own data), this is the same text for every account
        of a given statement type -- it exists to prime the model with the
        right CONCEPTS (which cost is rigid vs revenue-linked, when a sign
        flip is ordinary, how to read a BS reclassification) so it can spot
        and reason about a situation the deterministic checks above don't
        cover, e.g. a variable cost quietly diverging from revenue without
        crossing the >=30pp gap _variance_analysis_guidance flags.

        Still subject to every existing grounding rule -- this adds
        interpretive FRAMES, not license to invent facts; an inference
        drawn using one of these frames still needs a remark/data anchor
        and still must be marked as judgement, exactly like
        _variance_analysis_guidance's own reasoning requirement.

        Deliberately short: this is injected into EVERY account's prompt
        (Generator + Auditor only, not Refiner/Validator -- drafting and
        auditing are where analytical framing matters; final polish does
        not need to re-derive it), so every extra line here is a real,
        recurring prompt-budget cost, not a one-time one."""
        statement_type = str(
            ((df.attrs if isinstance(df, pd.DataFrame) else {}).get("integrity") or {}).get("statement_type") or ""
        ).strip().upper()
        if statement_type not in ("BS", "IS"):
            return ""

        if language == "Chi":
            if statement_type == "IS":
                return (
                    "【分析视角】折旧、房产税、专业服务费等属于刚性成本，通常不随收入同比例变动，这是正常现象，"
                    "无需强行解释；水电费净额、外包物业费等收入关联型成本若与收入趋势明显背离，则值得指出。"
                    "区分一次性事项（罚款、滞纳金）与经常性项目；汇兑损益方向转变属正常波动，除非备注另有说明。"
                    "涉及关联方的条款（计价公式、是否免息、是否于交割前结清）应说明并列明交易对手名称。"
                    "跨期（如1-3月）与全年数据比较仅可通过年化处理或与上年同期等长区间比较，不可直接比较绝对值。"
                )
            return (
                "【分析视角】应收账款/预收款项与营业收入趋势的联动关系值得关注；在建工程转固定资产会同时影响两个"
                "科目余额，可据此解释双向变动；一年内到期的非流动负债与长期借款之间的重分类是正常信号，未必是异常。"
                "涉及关联方的条款（计价公式、是否免息、是否于交割前结清）应说明并列明交易对手名称。"
            )

        if statement_type == "IS":
            return (
                "[ANALYTICAL LENS] Rigid/fixed costs (depreciation, property tax, professional-service fees) "
                "normally don't track revenue proportionally -- that's expected, not something to force an "
                "explanation for; revenue-linked costs (net utilities, outsourced property management fees) "
                "diverging from the revenue trend IS worth flagging. Distinguish one-off items (penalties, "
                "late fees) from recurring ones. An FX gain/loss flipping direction is ordinary volatility "
                "unless the remarks say otherwise. Related-party terms (pricing formula, interest-free or "
                "not, settled before close or not) should be stated with the counterparty named. Comparing a "
                "stub period (e.g. Jan-Mar) against a full year is only valid via annualisation or against an "
                "equal-length prior stub -- never a raw absolute comparison."
            )
        return (
            "[ANALYTICAL LENS] AR/advance receipts moving with the revenue trend is a normal working-capital "
            "link worth noting; a CIP-to-fixed-assets transfer moves both accounts' balances together and can "
            "explain a swing in either one; a reclassification between current-portion-of-long-term-debt and "
            "long-term loans is a normal signal, not necessarily a red flag. Related-party terms (pricing "
            "formula, interest-free or not, settled before close or not) should be stated with the "
            "counterparty named."
        )

    @staticmethod
    def _period_reference_guidance(df: Optional[pd.DataFrame], language: str) -> str:
        integrity = (df.attrs if isinstance(df, pd.DataFrame) else {}).get("integrity") or {}
        attrs = df.attrs if isinstance(df, pd.DataFrame) else {}
        statement_type = str(integrity.get("statement_type") or "").strip().upper()
        effective_date = str(integrity.get("effective_date") or "").strip()
        annualization_months = attrs.get("annualization_months")
        if annualization_months in (None, ""):
            annualization_months = integrity.get("annualization_months")
        fiscal_year_end_month = integrity.get("fiscal_year_end_month")
        fiscal_year_end_day = integrity.get("fiscal_year_end_day")

        if language == "Chi":
            if statement_type == "BS":
                return (
                    f"这是资产负债表科目。首句必须仅说明截至{effective_date}的最新期末余额（单一期间，不要罗列所有期间），"
                    f"并描述其构成（如：'截至{effective_date}余额为X万元，主要为[构成项]'或'截至{effective_date}余额合计X万元，主要包括[各组成项]'）。"
                    "首句不得罗列所有报告期间余额（避免'截至A、B、C日余额分别为X、Y、Z'式开篇），"
                    "也不得以年度对比句开篇（不得以'X较上年增加/减少'或'X同比上升/下降'作为首句）。"
                    "首句之后，再描述构成项目、对手方/集中度、合同条款及重要备注说明。"
                    "如跨期变动重大且数据支持，可简略提及前期余额，但不得作为开篇。"
                    f"请使用时点表述如【截至{effective_date}】，不要写成期间表述。"
                )
            if statement_type == "IS":
                period_label = build_income_statement_period_label(
                    effective_date,
                    months=annualization_months if isinstance(annualization_months, (int, float)) else None,
                    fiscal_year_end_month=fiscal_year_end_month if isinstance(fiscal_year_end_month, (int, float)) else None,
                    fiscal_year_end_day=fiscal_year_end_day if isinstance(fiscal_year_end_day, (int, float)) else None,
                    language="Chi",
                )
                is_partial = isinstance(annualization_months, (int, float)) and int(annualization_months) < 12
                partial_note = ""
                if is_partial:
                    months_int = int(annualization_months)
                    annualized_label = f"{period_label}（年化）"
                    partial_note = (
                        f"最新期间为截至{effective_date}的{months_int}个月（期间标签：{period_label}），属于不完整年度。"
                        f"跨期比较时，请优先使用年化后数据（x12/{months_int}，已预计算为【{annualized_label}】列）进行同口径对比，"
                        "并在评论中注明该数据已年化。"
                    )
                return (
                    "这是利润表科目。首句必须以构成开篇，描述该科目主要包含哪些项目"
                    "（例如：'X主要从租金收入及物业管理费收入产生营业收入，比例约为50:50'或"
                    "'主要包括房屋折旧费用A万元、物业管理费B万元、......'）。"
                    "不得以孤立的趋势句开篇（避免'营业收入由X增长至Y'式开篇）。"
                    "对每一重要构成项，应在句中提供所有报告期间的金额（**若本科目已有做好的明细表，则此项不适用**：各期金额由表格列示，正文不得逐项重复）（例如：'物业管理费分别为150万元、180万元、210万元，"
                    "于FY19、FY20、FY21期间发生'），而不是仅提供最新一期的金额。"
                    "构成与多期金额之后，如有重大变动，可在数据/备注支持下说明驱动因素。"
                    f"描述目标期间时，请使用【于{period_label}期间】或【在{period_label}内】等期间表述，"
                    f"不要写成【截至{effective_date}止】或时点余额表述。{partial_note}"
                    "有右侧备注的科目优先讨论。"
                )
            return "请根据科目属性正确区分时点表述与期间表述。"

        if statement_type == "BS":
            return (
                f"This is a balance-sheet item. The FIRST sentence must state ONLY the latest period-end balance as at {effective_date} "
                f"(a single period, not a list of all periods) and describe what it comprises — e.g., 'the balance as at {effective_date} "
                f"represented CNY X million of [composition]' or 'the balance as at {effective_date} totalled CNY X million, mainly entailing [components]'. "
                f"Do NOT dump all reporting periods in the opening sentence (avoid 'the balance as at A, B and C was X, Y and Z respectively'). "
                f"Do NOT open with a year-over-year movement sentence ('X increased/decreased from Y to Z'). "
                f"After the opening, describe composition, counterparty/concentration, terms, and any material remarks supported by the data. "
                f"Prior-period balances may appear briefly only when the movement is material and the data supports the explanation. "
                f"Use point-in-time wording such as 'as at {effective_date}', not period-flow wording."
            )
        if statement_type == "IS":
            period_label = build_income_statement_period_label(
                effective_date,
                months=annualization_months if isinstance(annualization_months, (int, float)) else None,
                fiscal_year_end_month=fiscal_year_end_month if isinstance(fiscal_year_end_month, (int, float)) else None,
                fiscal_year_end_day=fiscal_year_end_day if isinstance(fiscal_year_end_day, (int, float)) else None,
                language="Eng",
            )
            is_partial = isinstance(annualization_months, (int, float)) and int(annualization_months) < 12
            partial_note = ""
            if is_partial:
                months_int = int(annualization_months)
                annualized_label = f"{period_label} annualised"
                partial_note = (
                    f" The latest period covers {months_int} months ending {effective_date} (period label: {period_label}) — this is a partial year. "
                    f"For cross-year comparisons, use the annualized figures (×12/{months_int}), pre-calculated in the '{annualized_label}' column, "
                    f"and note the annualization in the commentary."
                )
            return (
                f"This is an income-statement item. The FIRST sentence must lead with COMPOSITION — describe what the line mainly comprises "
                f"(e.g., 'X mainly generated revenue from leasing income and property management service income with around a 50:50 ratio' "
                f"or 'mainly comprised depreciation and amortisation of CNY A, property management costs of CNY B, and...'). "
                f"Do NOT open with an isolated trend sentence ('Revenue increased from X to Y'). "
                f"For each material component, give the amounts across ALL reporting periods inline (e.g., 'CNY 1.5 million, CNY 1.8 million, "
                f"and CNY 2.1 million property management service costs incurred in FY19, FY20 and FY21 respectively'), not just the latest period figure. "
                f"After composition with multi-year amounts, state the driver of any material movement, supported by the data or remarks. "
                f"Refer to the target period with flow wording such as 'during {period_label}' or 'during the Period', not 'as at'.{partial_note} "
                f"Prioritize line items with supporting remarks."
            )
        return "Use point-in-time wording for balance-sheet style data and period-flow wording for income-statement style data."

    @staticmethod
    def _append_markdown_section(rendered: str, label: str, body: str) -> str:
        body = str(body or "").strip()
        if not body:
            return rendered
        return f"{rendered}\n\n{label}:\n{body}"

    def _append_markdown_table_section(self, rendered: str, label: str, rows: list[Dict[str, Any]]) -> str:
        table_df = pd.DataFrame(rows)
        if table_df.empty:
            return rendered
        return self._append_markdown_section(rendered, label, table_df.to_markdown(index=False).strip())

    def _build_financial_prompt_payload(
        self,
        df: Optional[pd.DataFrame],
        mapping_key: str,
        language: str,
        data_format: str,
        user_comment: str = "",
    ) -> Dict[str, str]:
        if df is None or df.empty:
            return {}
        analysis_df = self._build_analysis_prompt_df(df)
        analysis_label = str(df.attrs.get("prompt_analysis_label") or "All indicative adjusted periods").strip()
        supporting_notes = [str(note).strip() for note in (df.attrs.get("supporting_notes") or []) if str(note).strip()]
        adjacent_detail_rows = self.filter_adjacent_detail_rows(df)
        table_linked_remarks = self.table_linked_remarks(df)
        rhs_remark_summary = self.summarize_rhs_remarks(adjacent_detail_rows, language)
        format_language = language or str(df.attrs.get("report_language") or "").strip()
        formatted_analysis_df = self._format_analysis_prompt_df(analysis_df, format_language)
        normalized_prompt_df = self._normalize_prompt_dataframe(df, language)
        normalized_analysis_df = self._normalize_prompt_dataframe(formatted_analysis_df, language)
        normalized_supporting_notes = self._normalize_prompt_value(supporting_notes, language)
        prompt_ready_adjacent_detail_rows = self._prompt_ready_adjacent_detail_rows(adjacent_detail_rows, format_language)
        normalized_adjacent_detail_rows = self._normalize_prompt_value(prompt_ready_adjacent_detail_rows, language)
        normalized_table_linked_remarks = self._normalize_prompt_value(table_linked_remarks, language)
        normalized_rhs_remark_summary = self._normalize_prompt_value(rhs_remark_summary, language)
        normalized_user_comment = self._normalize_prompt_value(str(user_comment or "").strip(), language)
        normalized_analysis_label = self._normalize_prompt_value(analysis_label, language)
        normalized_mapping_key = self._normalize_prompt_value(mapping_key, language)
        trend_summary = build_trend_summary(analysis_df) if isinstance(analysis_df, pd.DataFrame) and not analysis_df.empty else {}
        significant_movements = (
            build_significant_movements(analysis_df)
            if isinstance(analysis_df, pd.DataFrame) and not analysis_df.empty
            else []
        )
        trend_summary = self._normalize_prompt_value(trend_summary, language)
        significant_movements = self._normalize_prompt_value(significant_movements, language)
        integrity = df.attrs.get("integrity") or {}
        latest_source_period = str(integrity.get("effective_date") or "").strip()
        target_period = latest_source_period or (str(df.columns[1]).strip() if len(df.columns) > 1 else "")
        annualization_months = df.attrs.get("annualization_months")
        if annualization_months in (None, ""):
            annualization_months = integrity.get("annualization_months")
        fiscal_year_end_month = integrity.get("fiscal_year_end_month")
        fiscal_year_end_day = integrity.get("fiscal_year_end_day")
        target_period_label = (
            build_income_statement_period_label(
                target_period,
                months=annualization_months if isinstance(annualization_months, (int, float)) else None,
                fiscal_year_end_month=fiscal_year_end_month if isinstance(fiscal_year_end_month, (int, float)) else None,
                fiscal_year_end_day=fiscal_year_end_day if isinstance(fiscal_year_end_day, (int, float)) else None,
                language=language,
            )
            if str(integrity.get("statement_type") or "").strip().upper() == "IS"
            else target_period
        )
        target_period_guidance = {
            "target_period": target_period,
            "target_period_label": target_period_label,
            "latest_source_period": latest_source_period,
            "instruction": (
                "Use the target_period as the main reporting period for the latest balance or latest period statement. "
                "For income-statement items, prefer period labels such as target_period_label and use 'during', not 'for the period ended'. "
                "Use all earlier indicative-adjusted periods for trend, comparison, reasonableness checks, and significant movement analysis."
            ),
        }
        target_period_guidance = self._normalize_prompt_value(target_period_guidance, language)

        if data_format == "json":
            payload = json.loads(
                df_to_json_str(
                    normalized_prompt_df if isinstance(normalized_prompt_df, pd.DataFrame) else df,
                    table_name=normalized_mapping_key,
                    language=language,
                    text_normalizer=normalize_english_text if language == "Eng" else None,
                )
            )
            payload["reporting_focus"] = target_period_guidance
            if isinstance(normalized_analysis_df, pd.DataFrame) and not normalized_analysis_df.empty:
                payload["analysis_periods"] = json.loads(
                    df_to_json_str(
                        normalized_analysis_df,
                        table_name=normalized_analysis_label,
                        language=language,
                        text_normalizer=normalize_english_text if language == "Eng" else None,
                    )
                )
            if trend_summary:
                payload["trend_summary"] = trend_summary
            if significant_movements:
                payload["significant_movements"] = significant_movements
            if normalized_supporting_notes:
                payload["supporting_context"] = normalized_supporting_notes
            if normalized_adjacent_detail_rows:
                payload["supplemental_side_column_context"] = normalized_adjacent_detail_rows
            if normalized_table_linked_remarks:
                payload["table_context_observations"] = normalized_table_linked_remarks
            if normalized_rhs_remark_summary:
                payload["supplemental_context_summary"] = normalized_rhs_remark_summary
            if normalized_user_comment:
                payload["user_guidance"] = [normalized_user_comment]
                payload["user_guidance_instruction"] = (
                    "Treat these user remarks as explicit writing or reprompt guidance only to the extent they are supported by the provided data, notes, and remarks."
                    if language == "Eng"
                    else "将这些用户备注视为明确的写作/重提示指引，但仅可在其与提供的数据、备注及说明一致时采用。"
                )
            rendered = json.dumps(payload, ensure_ascii=False, indent=2)
        else:
            rendered = self._build_markdown_prompt_payload(
                normalized_prompt_df if isinstance(normalized_prompt_df, pd.DataFrame) else df
            )["financial_data"]
            focus_label = "Reporting focus" if language == "Eng" else "报告重点"
            focus_lines = [
                f"- Target period: {target_period}" if language == "Eng" else f"- 目标期间: {target_period}",
                (
                    f"- Preferred narrative period label: {target_period_label}"
                    if language == "Eng"
                    else f"- 推荐叙述期间标签: {target_period_label}"
                ),
                (
                    f"- Latest source indicative-adjusted period: {latest_source_period}"
                    if language == "Eng"
                    else f"- 最新示意性调整后源期间: {latest_source_period}"
                ),
                (
                    "- Use the target period for the latest balance / latest period statement. For income-statement items, prefer 'during' wording with the preferred narrative period label rather than 'for the period ended'. Use earlier indicative-adjusted periods for trend, comparison, cross-check, and significant movement analysis."
                    if language == "Eng"
                    else '- 以目标期间作为最新余额/最新期间表述的基础。若为利润表科目，优先使用推荐叙述期间标签并采用"于...期间/在...内"表达，而不是"截至...止期间"。同时使用更早的示意性调整后期间进行趋势、比较、交叉检查及重大变动分析。'
                ),
            ]
            rendered = self._append_markdown_section(rendered, focus_label, "\n".join(focus_lines))
            if isinstance(normalized_analysis_df, pd.DataFrame) and not normalized_analysis_df.empty:
                analysis_block = normalized_analysis_df.to_markdown(index=False).strip()
                rendered = self._append_markdown_section(rendered, normalized_analysis_label, analysis_block)
            if trend_summary:
                trend_lines = [
                    f"- {key}: {value}"
                    for key, value in trend_summary.items()
                    if value not in (None, "", [], {})
                ]
                if trend_lines:
                    trend_label = "Trend summary" if language == "Eng" else "趋势摘要"
                    rendered = self._append_markdown_section(rendered, trend_label, "\n".join(trend_lines))
            if significant_movements:
                change_label = "Significant movements" if language == "Eng" else "重大变动"
                rendered = self._append_markdown_table_section(rendered, change_label, significant_movements)

        if data_format != "json" and normalized_supporting_notes:
            notes_label = "Supporting context" if language == "Eng" else "补充备注"
            notes_block = "\n".join(f"- {note}" for note in normalized_supporting_notes)
            rendered = self._append_markdown_section(rendered, notes_label, notes_block)

        if data_format != "json" and normalized_adjacent_detail_rows:
            details_label = "Supplemental side-column context" if language == "Eng" else "右侧备注/原因"
            rendered = self._append_markdown_table_section(rendered, details_label, normalized_adjacent_detail_rows)

        if data_format != "json" and normalized_table_linked_remarks:
            table_linked_label = "Table context observations" if language == "Eng" else "表格关联备注"
            rendered = self._append_markdown_table_section(rendered, table_linked_label, normalized_table_linked_remarks)

        if data_format != "json" and normalized_rhs_remark_summary:
            summary_label = "Supplemental context summary" if language == "Eng" else "右侧备注摘要"
            summary_block = "\n".join(f"- {item}" for item in normalized_rhs_remark_summary)
            rendered = self._append_markdown_section(rendered, summary_label, summary_block)

        if data_format != "json" and normalized_user_comment:
            comment_label = "User guidance" if language == "Eng" else "用户备注 / 重提示指引"
            rendered = self._append_markdown_section(rendered, comment_label, f"- {normalized_user_comment}")

        if language == "Eng":
            rendered = normalize_english_text(rendered)

        return {"financial_figure": rendered, "financial_data": rendered}

    def _safe_format(self, template: str, format_params: Dict[str, Any]) -> str:
        if not template:
            return template
        try:
            return template.format(**format_params)
        except KeyError as exc:
            self.logger.warning("Missing prompt key %s. Available keys: %s", exc, list(format_params.keys()))
            return template

    def render_prompt(
        self,
        agent_name: str,
        language: str,
        mapping_key: str,
        df: Optional[pd.DataFrame] = None,
        data_format: str = "markdown",
        **kwargs,
    ) -> Tuple[str, str]:
        system_prompt, user_prompt_template = self.get_prompt_pair(agent_name, language, mapping_key)
        style_pack = PromptStylePack(language)
        dynamic_mapping_context = {}
        if isinstance(df, pd.DataFrame):
            dynamic_mapping_context = dict(df.attrs.get("dynamic_mapping_context") or {})
        normalized_kwargs = self._normalize_prompt_value(kwargs, language)
        format_params = {
            "key": self._normalize_prompt_value(mapping_key, language),
            "language": language,
            "accounting_nature": (
                self._normalize_prompt_value(
                    kwargs.get("accounting_nature")
                    or dynamic_mapping_context.get("accounting_nature")
                    or dynamic_mapping_context.get("category")
                    or "",
                    language,
                )
            ),
            "language_instruction": style_pack.language_instruction(),
            "common_formatting": style_pack.common_formatting_rules(),
            "fdd_judgement": style_pack.fdd_judgement_rules(),
            "common_data_rules": style_pack.common_data_rules(data_format),
            "period_reference_guidance": self._period_reference_guidance(df, language),
            "variance_analysis_guidance": self._variance_analysis_guidance(
                df, language, peer_context=kwargs.get("peer_context"), mapping_key=mapping_key,
                thresholds=kwargs.get("analysis_thresholds"),
            ),
            "analytical_lens_guidance": self._analytical_lens_guidance(df, language),
            "detail_table_guidance": self._detail_table_guidance(df, language),
            "composition_guidance": self._composition_guidance(df, language),
            "rhs_guidance_block": self._rhs_guidance_block(
                self.filter_adjacent_detail_rows(df) if isinstance(df, pd.DataFrame) else [],
                language,
            ),
            "remarks_weight_instruction": self._remarks_weight_instruction(
                has_rhs_remarks=bool(self.filter_adjacent_detail_rows(df) if isinstance(df, pd.DataFrame) else []),
                has_supporting_notes=bool((df.attrs.get("supporting_notes") or []) if isinstance(df, pd.DataFrame) else []),
                has_user_comment=bool(str(kwargs.get("user_comment", "")).strip()),
                statement_type=((df.attrs.get("integrity") or {}).get("statement_type") if isinstance(df, pd.DataFrame) else ""),
                language=language,
            ),
            "user_guidance_instruction": self._user_guidance_instruction(kwargs.get("user_comment", ""), language),
            **normalized_kwargs,
        }
        format_params.update(
            self._build_financial_prompt_payload(
                df=df,
                mapping_key=mapping_key,
                language=language,
                data_format=data_format,
                user_comment=kwargs.get("user_comment", ""),
            )
        )

        if self.normalize_agent_name(agent_name) == "1_Generator":
            resolved_mapping_key = self.resolve_mapping_key(mapping_key)
            patterns = self.get_mapping_component(
                resolved_mapping_key,
                component="patterns",
            ) if self._patterns_enabled() else None
            if patterns is None and self._patterns_enabled():
                fallback_section = self._fallback_mapping_section(mapping_key)
                if fallback_section:
                    patterns = self.mappings_data.get(fallback_section, {}).get("patterns")
            # Format the patterns dict into clean numbered examples so the
            # LLM sees readable text rather than a Python dict repr.
            #
            # These examples are the strongest signal in the prompt, and an
            # audit of all 80 of them (inspect_mapping_patterns.py) found 46
            # are complete sentences with only <SLOT> gaps, 4 assert facts
            # outright, and Cash's two contradict each other -- one says the
            # bank statements have NOT been obtained while the other says no
            # differences were noted. A real databook states they WERE
            # checked, so an unframed example can teach the opposite of the
            # data. Framing them explicitly as structure-and-register only,
            # with facts reserved to the data and remarks, keeps their value
            # as a style anchor without letting them decide what is true.
            if isinstance(patterns, dict):
                examples = []
                for idx, (pname, v) in enumerate(patterns.items(), 1):
                    text = str(v or "").strip()
                    if text and text.upper() != "N/A":
                        # The KEY carries the precondition where an account's
                        # variants genuinely conflict -- Cash reads differently
                        # depending on whether the bank statements were
                        # obtained. Rendering only the value dropped that, and
                        # left the model choosing between contradictory
                        # sentences with nothing to choose on.
                        cond = re.search(r"[（(](.+)[)）]\s*$", str(pname).strip())
                        suffix = f" [{cond.group(1).strip()}]" if cond else ""
                        examples.append(f"Example {idx}{suffix}: {text}")
                patterns = "\n".join(examples) if examples else ""
            patterns_text = str(patterns or "").strip()
            if patterns_text:
                if language == "Chi":
                    framing = (
                        "以下示例仅用于说明**句式结构与用语风格**（开篇方式、动词选择、构成的表述顺序），"
                        "**不是可以照抄的事实**。示例中出现的任何具体情况——例如是否已取得银行对账单、"
                        "是否已执行核对、是否未发现差异、是否已验资——都必须以本科目的实际数据与备注为准；"
                        "若备注未提及，就不要写入该说法。示例之间如有矛盾（同一科目一个说已核对、"
                        "另一个说尚未取得），说明它们只是不同项目的历史写法，不代表本项目的事实。"
                        "请按本科目的真实情况撰写，只借用其行文方式：\n"
                    )
                else:
                    framing = (
                        "The examples below illustrate SENTENCE STRUCTURE AND REGISTER only -- how to "
                        "open, which verbs to use, the order in which composition is stated. They are "
                        "NOT facts to copy. Anything specific they assert -- whether bank statements "
                        "were obtained, whether a check was performed, whether no differences were "
                        "found, whether capital was verified -- must come from THIS account's own data "
                        "and remarks; if the remarks do not say it, do not write it. Where two examples "
                        "contradict each other, that only reflects different past engagements, not this "
                        "one. Write what is true here, borrowing only the phrasing:\n"
                    )
                patterns_text = framing + patterns_text
            format_params["patterns"] = patterns_text

        rendered_system_prompt = self._safe_format(system_prompt, format_params)
        rendered_user_prompt = self._safe_format(user_prompt_template, format_params)

        if self.normalize_agent_name(agent_name) == "1_Generator":
            previous_content = str(kwargs.get("previous_content") or "").strip()
            if previous_content:
                if language == "Chi":
                    rendered_user_prompt = (
                        f"{rendered_user_prompt}\n\n"
                        "已验证旧评论（请在保留数据支持内容的基础上按新的用户备注进行改写，而不是完全重写方向）：\n"
                        f"{previous_content}\n\n"
                        "请将上述旧评论视为待修订底稿。优先保留其中仍被当前数据、备注及右侧说明支持的内容，并结合用户最新指引进行定向修改。"
                    )
                else:
                    rendered_user_prompt = (
                        f"{rendered_user_prompt}\n\n"
                        "Existing validated commentary to revise (treat this as the draft to update rather than starting from scratch):\n"
                        f"{previous_content}\n\n"
                        "Use the existing validated commentary as the baseline draft. Keep the parts that are still supported by the current data, remarks, and notes, and revise it directionally based on the latest user guidance."
                    )

        return rendered_system_prompt, rendered_user_prompt

    def filter_adjacent_detail_rows(self, df: pd.DataFrame) -> list[Dict[str, Any]]:
        return self._filter_adjacent_detail_rows(df)

    def table_linked_remarks(self, df: Optional[pd.DataFrame]) -> list[Dict[str, Any]]:
        return self._table_linked_remarks(df)

    def summarize_rhs_remarks(self, adjacent_detail_rows: list[Dict[str, Any]], language: str) -> list[str]:
        return self._summarize_rhs_remarks(adjacent_detail_rows, language)

    def build_prompt_context_snapshot(
        self,
        df: Optional[pd.DataFrame],
        language: str = "Eng",
        user_comment: str = "",
        previous_output: str = "",
    ) -> Dict[str, Any]:
        if not isinstance(df, pd.DataFrame):
            return {}
        attrs = df.attrs or {}
        integrity = attrs.get("integrity") or {}
        supporting_notes = attrs.get("supporting_notes") or []
        adjacent_detail_rows = self.filter_adjacent_detail_rows(df)
        table_linked_remarks = self.table_linked_remarks(df)
        rhs_remark_summary = self.summarize_rhs_remarks(adjacent_detail_rows, language)
        return {
            "sheet_name": integrity.get("sheet_name"),
            "statement_type": integrity.get("statement_type"),
            "effective_date": integrity.get("effective_date"),
            "selected_variant": attrs.get("selected_variant"),
            "prompt_analysis_label": attrs.get("prompt_analysis_label"),
            "supporting_notes": supporting_notes,
            "supporting_notes_count": len(supporting_notes),
            "adjacent_detail_rows": adjacent_detail_rows,
            "rhs_remark_count": len(adjacent_detail_rows),
            "rhs_remark_summary": rhs_remark_summary,
            "table_linked_remarks": table_linked_remarks,
            "table_linked_remarks_count": len(table_linked_remarks),
            "user_comment": str(user_comment or "").strip(),
            "has_previous_output": bool(str(previous_output or "").strip()),
            "previous_output_excerpt": str(previous_output or "").strip()[:500],
        }
# --- end ai/prompts.py ---
