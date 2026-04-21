from __future__ import annotations

# --- begin pptx/text.py ---
from concurrent.futures import ThreadPoolExecutor, as_completed
import re
from typing import Optional

from pptx.util import Pt


def get_tab_name(project_name: str) -> Optional[str]:
    if not project_name:
        return None
    words = project_name.split()
    if words:
        return words[0]
    return None


def clean_content_quotes(content: str) -> str:
    if not content:
        return ""
    content = re.sub(r'^"*|"*$', "", content.strip())
    content = re.sub(r'""+', '"', content)
    return content


def detect_chinese_text(text: str, force_chinese_mode: bool = False) -> bool:
    if force_chinese_mode:
        return True
    return contains_predominantly_chinese_text(text)


def get_font_size_for_text(text: str, base_size: int = 9, force_chinese_mode: bool = False) -> Pt:
    is_chinese = detect_chinese_text(text, force_chinese_mode)
    return Pt(base_size + 1) if is_chinese else Pt(base_size)


def get_font_name_for_text(text: str, default_font: str = "Arial") -> str:
    return "Microsoft YaHei" if detect_chinese_text(text) else default_font


def get_line_spacing_for_text(text: str, force_chinese_mode: bool = False) -> float:
    return 0.9 if detect_chinese_text(text, force_chinese_mode) else 1.0


def get_space_after_for_text(text: str, force_chinese_mode: bool = False) -> Pt:
    return Pt(6) if detect_chinese_text(text, force_chinese_mode) else Pt(4)


def get_space_before_for_text(text: str, force_chinese_mode: bool = False) -> Pt:
    return Pt(3) if detect_chinese_text(text, force_chinese_mode) else Pt(2)


def replace_entity_placeholders(content: str, project_name: str) -> str:
    if not content or not project_name:
        return content
    replacements = {
        "[PROJECT]": project_name,
        "[Entity]": project_name,
        "[Company]": project_name,
    }
    for placeholder, replacement in replacements.items():
        content = content.replace(placeholder, replacement)
    return content
# --- end pptx/text.py ---

# --- begin pptx/payloads.py ---
from typing import Any, Dict, Iterable, List, Optional

import pandas as pd

from .financial_common import (
    contains_chinese_text,
    contains_predominantly_chinese_text,
    get_pipeline_result_text,
    load_yaml_file,
    package_file_path,
)
from .keyword_registry import STATEMENT_ORDER_SKIP_KEYWORDS, translate_category_to_chinese
from .workbook import find_mapping_key


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
        "max_workers": 4,
        "enable_validation": True,
        "generation_temperature": 0.2,
        "validation_temperature": 0.1,
    },
    "commentary_packing": {
        "use_pillow_text_fitting": True,
        "shape_height_utilization": 1.08,
        "minimum_slot_lines": 22,
        "split_min_remaining_lines": 3,
        "split_min_content_lines": 5,
        "move_whole_min_fill_ratio": 0.74,
        "target_fill_min_ratio": 0.9,
        "target_fill_max_ratio": 0.96,
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
                "shape_height_utilization": 1.13,
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


def _build_compact_summary_text(
    text: str,
    *,
    is_chinese: bool,
    max_sentences: int,
    max_chars: int,
    max_numeric_sentences: int,
) -> str:
    sentences = _split_text_sentences(text, is_chinese)
    if not sentences:
        return str(text or "").strip()

    selected: List[str] = []
    numeric_sentences = 0
    for sentence in sentences:
        heavy = _sentence_is_numeric_heavy(sentence)
        if heavy and numeric_sentences >= max_numeric_sentences:
            continue
        candidate = _join_text_sentences(selected + [sentence], is_chinese)
        if selected and len(candidate) > max_chars:
            break
        selected.append(sentence)
        if heavy:
            numeric_sentences += 1
        if len(selected) >= max_sentences:
            break

    if not selected:
        selected = sentences[:1]

    summary = _join_text_sentences(selected, is_chinese).strip()
    if len(summary) > max_chars:
        summary = summary[:max_chars].rstrip(" ,;:/-") + "..."
    return summary.strip()


def _normalize_slide_commentary_text(text: str) -> str:
    normalized = clean_content_quotes(str(text or ""))
    if not normalized:
        return ""
    normalized = normalized.replace("\r\n", "\n")
    normalized = re.sub(r"[ \t]+", " ", normalized)
    normalized = re.sub(r"\n{3,}", "\n\n", normalized)
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
    return get_pipeline_result_text(result_dict)


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


def _has_significant_balance(financial_data: Optional[pd.DataFrame]) -> bool:
    if financial_data is None or financial_data.empty:
        return True

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
            continue

        final_content = _extract_final_content(result)
        commentary_text = (
            str(final_content).strip()
            if final_content and str(final_content).strip()
            else f"[No content generated for {account_key}]"
        )

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
                    "category": mappings[mapping_key].get("category", ""),
                    "financial_data": financial_data,
                    "commentary": commentary_text,
                    "summary": _extract_summary(final_content) if final_content else "",
                    "is_chinese": contains_chinese_text(commentary_text),
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

# --- begin pptx/exporters.py ---
import logging
import os
import time
import traceback
from typing import Dict, List, Optional

from pptx import Presentation

logger = logging.getLogger(__name__)


class ReportGenerator:
    """Report generator that orchestrates PPTX creation from markdown."""

    def __init__(
        self,
        template_path: str,
        markdown_file: str,
        output_path: str,
        project_name: Optional[str] = None,
        language: str = "english",
        row_limit: int = 20,
    ):
        self.template_path = template_path
        self.markdown_file = markdown_file
        self.output_path = output_path
        self.project_name = project_name
        self.language = language
        self.row_limit = row_limit

    def generate(self):
        logger.info("Starting PPTX generation...")
        logger.info("Template: %s", self.template_path)
        logger.info("Markdown: %s", self.markdown_file)
        logger.info("Output: %s", self.output_path)
        logger.info("Language: %s", self.language)
        logger.info("Project: %s", self.project_name)

        with open(self.markdown_file, "r", encoding="utf-8") as handle:
            md_content = handle.read()

        logger.info("Content length: %s characters", len(md_content))
        generator = PowerPointGenerator(self.template_path, self.language, self.row_limit)

        try:
            generator.generate_full_report(md_content, None, self.output_path)
            if self.project_name:
                generator.update_project_titles(self.project_name, "BS")
        except Exception as exc:
            logger.error("Report generation failed: %s", exc)
            raise

        logger.info("PPTX generation completed: %s", self.output_path)


def export_pptx(
    template_path: str,
    markdown_path: str,
    output_path: str,
    project_name: Optional[str] = None,
    _excel_file_path: Optional[str] = None,
    language: str = "english",
    statement_type: str = "BS",
    row_limit: int = 20,
    model_type: Optional[str] = None,
):
    generator = ReportGenerator(template_path, markdown_path, output_path, project_name, language, row_limit)
    generator.generate()

    if not os.path.exists(output_path):
        raise FileNotFoundError(f"PPTX file was not created at {output_path}")

    if project_name:
        temp_presentation = Presentation(output_path)
        pptx_gen = PowerPointGenerator(template_path, language, row_limit, model_type=model_type)
        pptx_gen.presentation = temp_presentation
        pptx_gen.update_project_titles(project_name, statement_type)
        temp_presentation.save(output_path)

    logger.info("PowerPoint presentation successfully exported to: %s", output_path)
    return output_path


def export_pptx_from_structured_data_combined(
    template_path: str,
    bs_data: List[Dict],
    is_data: List[Dict],
    output_path: str,
    project_name: Optional[str] = None,
    language: str = "english",
    temp_path: Optional[str] = None,
    selected_sheet: Optional[str] = None,
    is_chinese_databook: bool = False,
    bs_is_results: Optional[Dict[str, Any]] = None,
    model_type: Optional[str] = None,
):
    try:
        export_started_at = time.perf_counter()
        logger.info("Starting COMBINED PPTX generation...")
        logger.info("Template: %s", template_path)
        logger.info("Output: %s", output_path)
        logger.info("Language: %s", language)
        logger.info("BS accounts: %s, IS accounts: %s", len(bs_data), len(is_data))

        generator = PowerPointGenerator(template_path, language, row_limit=20, model_type=model_type)
        stage_started_at = time.perf_counter()
        generator.load_template()
        logger.info("PPTX stage load_template took %.2fs", time.perf_counter() - stage_started_at)

        if bs_data:
            stage_started_at = time.perf_counter()
            generator.apply_structured_data_to_slides(bs_data, 1, project_name, "BS", is_chinese_databook=is_chinese_databook)
            logger.info("PPTX stage apply_bs_slides took %.2fs", time.perf_counter() - stage_started_at)
        if is_data:
            stage_started_at = time.perf_counter()
            generator.apply_structured_data_to_slides(is_data, 5, project_name, "IS", is_chinese_databook=is_chinese_databook)
            logger.info("PPTX stage apply_is_slides took %.2fs", time.perf_counter() - stage_started_at)
        if temp_path and selected_sheet:
            stage_started_at = time.perf_counter()
            generator.embed_financial_tables(
                temp_path,
                selected_sheet,
                project_name,
                language,
                bs_is_results=bs_is_results,
            )
            logger.info("PPTX stage embed_financial_tables took %.2fs", time.perf_counter() - stage_started_at)
        if hasattr(generator, "_unused_slides_to_remove") and generator._unused_slides_to_remove:
            stage_started_at = time.perf_counter()
            unused_slides_sorted = sorted(set(generator._unused_slides_to_remove), reverse=True)
            logger.info("Removing %s unused slides at the end: %s", len(unused_slides_sorted), [idx + 1 for idx in unused_slides_sorted])
            generator._remove_slides(unused_slides_sorted)
            logger.info("Removed %s unused slides", len(unused_slides_sorted))
            logger.info("PPTX stage remove_unused_slides took %.2fs", time.perf_counter() - stage_started_at)
        if project_name:
            stage_started_at = time.perf_counter()
            generator.refresh_project_placeholders(project_name)
            logger.info("PPTX stage refresh_project_placeholders took %.2fs", time.perf_counter() - stage_started_at)

        stage_started_at = time.perf_counter()
        generator.save(output_path)
        logger.info("PPTX stage save_presentation took %.2fs", time.perf_counter() - stage_started_at)
        logger.info("PPTX combined export total took %.2fs", time.perf_counter() - export_started_at)
        logger.info("Combined PPTX generation completed: %s", output_path)
        return output_path
    except Exception as exc:
        logger.error("PPTX generation failed: %s", exc)
        logger.error(traceback.format_exc())
        raise


def export_pptx_from_structured_data(
    template_path: str,
    structured_data: List[Dict],
    output_path: str,
    project_name: Optional[str] = None,
    language: str = "english",
    statement_type: str = "BS",
    start_slide: int = 1,
    model_type: Optional[str] = None,
):
    try:
        logger.info("Starting PPTX generation from structured data...")
        logger.info("Template: %s", template_path)
        logger.info("Output: %s", output_path)
        logger.info("Language: %s", language)
        logger.info("Statement type: %s, Start slide: %s", statement_type, start_slide)
        logger.info("Accounts to process: %s", len(structured_data))

        generator = PowerPointGenerator(template_path, language, row_limit=20, model_type=model_type)
        generator.load_template()
        generator.apply_structured_data_to_slides(structured_data, start_slide, project_name, statement_type)
        generator.save(output_path)

        logger.info("PPTX generation completed: %s", output_path)
        return output_path
    except Exception as exc:
        logger.error("PPTX generation failed: %s", exc)
        raise


def merge_presentations(bs_presentation_path: str, is_presentation_path: str, output_path: str):
    try:
        logger.info("🔄 Starting presentation merge...")
        logger.info("   BS: %s", bs_presentation_path)
        logger.info("   IS: %s", is_presentation_path)

        merged_prs = Presentation(bs_presentation_path)
        is_prs = Presentation(is_presentation_path)

        from copy import deepcopy

        for slide_idx, slide in enumerate(is_prs.slides):
            try:
                slide_layout = slide.slide_layout
                new_slide = merged_prs.slides.add_slide(slide_layout)

                source_slide_xml = slide._element
                target_slide_xml = new_slide._element

                shapes_to_remove = list(new_slide.shapes)
                for shape in shapes_to_remove:
                    try:
                        sp_tree = target_slide_xml.get_or_add_spTree()
                        sp_tree.remove(shape._element)
                    except Exception:
                        pass

                source_sp_tree = source_slide_xml.get_or_add_spTree()
                target_sp_tree = target_slide_xml.get_or_add_spTree()
                for shape_element in source_sp_tree:
                    target_sp_tree.append(deepcopy(shape_element))

            except Exception as exc:
                logger.error("Error copying slide %s, using fallback method: %s", slide_idx, exc)
                slide_layout = slide.slide_layout
                new_slide = merged_prs.slides.add_slide(slide_layout)
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for new_shape in new_slide.shapes:
                            if (
                                hasattr(new_shape, "name")
                                and hasattr(shape, "name")
                                and new_shape.name == shape.name
                                and new_shape.has_text_frame
                            ):
                                new_shape.text_frame.text = shape.text_frame.text
                                break

        merged_prs.save(output_path)
        del merged_prs
        del is_prs

        import gc

        gc.collect()
        logger.info("✅ Presentation merge completed successfully")
    except Exception as exc:
        logger.error("Presentation merge failed: %s", exc)
        raise
# --- end pptx/exporters.py ---

# --- begin pptx/generation.py ---
"""
PowerPoint Generation Module for Financial Reports
Based on the backup methods but implemented fresh for the new system
"""

import os
import re
import logging
import threading
import time
from typing import Any, Dict, List, Optional, Tuple
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE_TYPE

logger = logging.getLogger(__name__)
logger.setLevel(logging.WARNING)


class PowerPointGenerator:
    """Main PowerPoint generator class"""

    def __init__(
        self,
        template_path: str,
        language: str = 'english',
        row_limit: int = 20,
        model_type: Optional[str] = None,
    ):
        self.template_path = template_path
        self.language = language.lower()
        self.row_limit = row_limit
        self.model_type = str(model_type or "").strip() or None
        self.presentation = None
        self.pptx_settings = _load_pptx_settings()

    def load_template(self):
        """Load the PowerPoint template"""
        if not os.path.exists(self.template_path):
            raise FileNotFoundError(f"Template not found: {self.template_path}")

        self.presentation = Presentation(self.template_path)
        logger.info("Loaded template: %s", self.template_path)

    def find_shape_by_name(self, shapes, name: str):
        """Find shape by name in slide (case-insensitive), recursive"""
        name_lower = name.lower()
        for shape in shapes:
            if hasattr(shape, 'name') and (shape.name == name or shape.name.lower() == name_lower):
                return shape
            
            # Check for group
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                found = self.find_shape_by_name(shape.shapes, name)
                if found:
                    return found
        return None

    @staticmethod
    def _is_commentary_text_shape(shape) -> bool:
        if not getattr(shape, "has_text_frame", False):
            return False
        shape_name = str(getattr(shape, "name", "") or "").lower()
        excluded_tokens = (
            "title",
            "projtitle",
            "summary",
            "cosummaryshape",
            "table",
            "subtitle",
        )
        return not any(token in shape_name for token in excluded_tokens)

    def _resolve_commentary_slot_shape(self, slide, slot_name: str, used_shape_ids=None):
        """Resolve the best text box for a commentary slot on a slide."""
        used_shape_ids = used_shape_ids or set()
        preferred_names = {
            "single": [
                "textMainBullets",
                "textMainBullets_L",
                "textMainBullets_R",
                "Text-commentary",
                "Content",
                "MainContent",
                "Body",
            ],
            "L": [
                "textMainBullets_L",
                "textMainBullets",
                "Text-commentary",
                "Content",
                "MainContent",
                "Body",
            ],
            "R": [
                "textMainBullets_R",
                "textMainBullets",
                "Text-commentary",
                "Content",
                "MainContent",
                "Body",
            ],
        }

        for name in preferred_names.get(slot_name, []):
            shape = self.find_shape_by_name(slide.shapes, name)
            if shape and getattr(shape, "has_text_frame", False) and id(shape) not in used_shape_ids:
                return shape

        generic_candidates = [
            shape for shape in slide.shapes
            if self._is_commentary_text_shape(shape) and id(shape) not in used_shape_ids
        ]
        if not generic_candidates:
            return None

        if slot_name == "L":
            return min(generic_candidates, key=lambda shape: (getattr(shape, "left", 0), -getattr(shape, "width", 0)))
        if slot_name == "R":
            return max(generic_candidates, key=lambda shape: (getattr(shape, "left", 0), getattr(shape, "width", 0)))
        return max(generic_candidates, key=lambda shape: (getattr(shape, "width", 0), -getattr(shape, "left", 0)))

    def _add_commentary_slot_shape(self, slide, slot_name: str):
        top = Inches(2.22)
        width = Inches(4.78)
        height = Inches(4.13)
        if slot_name == "L":
            left = Inches(0.13)
        elif slot_name == "R":
            left = Inches(5.09)
        else:
            # Page 1 template uses a single commentary box on the right beside the table.
            left = Inches(5.09)
        return slide.shapes.add_textbox(left, top, width, height)

    def _summary_settings(self) -> Dict[str, Any]:
        return dict(self.pptx_settings.get("executive_summary") or {})

    def _packing_settings(self, statement_type: Optional[str] = None) -> Dict[str, Any]:
        packing = dict(self.pptx_settings.get("commentary_packing") or {})
        if not statement_type:
            return packing
        overrides = ((packing.get("statement_overrides") or {}).get(statement_type) or {})
        if not overrides:
            return packing
        return _merge_nested_dict(packing, overrides)

    def _resolve_summary_model_type(self, is_chinese: bool) -> str:
        cached = getattr(self, "_summary_model_type_cache", None)
        if cached:
            return cached

        if self.model_type:
            resolved_model_type = str(self.model_type).strip()
            self._summary_model_type_cache = resolved_model_type
            return resolved_model_type

        try:
            from fdd_utils.ai import FDDConfig

            config = load_yaml_file(package_file_path("config.yml")) or {}
            requested_model_type = (
                str(self.model_type).strip()
                if self.model_type
                else str(((config.get("default") or {}).get("ai_provider")) or "deepseek")
            )
            config_manager = FDDConfig(
                language="Chi" if is_chinese else "Eng",
                model_type=requested_model_type,
            )
            resolved_model_type = str(config_manager.model_type or requested_model_type)
        except Exception as exc:
            logger.warning("Could not resolve PPTX summary model type, defaulting to deepseek: %s", exc)
            resolved_model_type = "deepseek"

        self._summary_model_type_cache = resolved_model_type
        return resolved_model_type

    def _summary_max_workers(self, summary_jobs: List[Dict[str, Any]]) -> int:
        if not summary_jobs:
            return 1

        summary_settings = self._summary_settings()
        configured_workers = int(summary_settings.get("max_workers", 4) or 4)
        model_type = self._resolve_summary_model_type(bool(summary_jobs[0].get("is_chinese")))
        if model_type == "local":
            configured_workers = int(summary_settings.get("local_max_workers", 1) or 1)
        return max(1, min(configured_workers, len(summary_jobs)))

    def _call_with_timeout(self, func, timeout_seconds: float, timeout_label: str):
        if timeout_seconds <= 0:
            return func()

        result_container = {"value": None, "error": None, "completed": False}

        def _run():
            try:
                result_container["value"] = func()
            except Exception as exc:
                result_container["error"] = exc
            finally:
                result_container["completed"] = True

        worker = threading.Thread(target=_run, daemon=True)
        worker.start()
        worker.join(timeout=timeout_seconds)

        if not result_container["completed"]:
            raise TimeoutError(f"{timeout_label} timed out after {timeout_seconds:.1f} seconds")
        if result_container["error"] is not None:
            raise result_container["error"]
        return result_container["value"]

    def _call_with_timeout_retry(
        self,
        func,
        timeout_seconds: float,
        max_retries: int,
        timeout_label: str,
    ):
        """Call ``func`` with a per-attempt timeout and retry on TimeoutError
        or other transient failures. Raises the last exception if all retries
        fail. Use ``max_retries >= 1`` (1 means "no retry, just run once")."""
        attempts = max(1, int(max_retries or 1))
        last_error: Optional[BaseException] = None
        for attempt in range(1, attempts + 1):
            label = (
                timeout_label
                if attempts == 1
                else f"{timeout_label} (attempt {attempt}/{attempts})"
            )
            try:
                return self._call_with_timeout(func, timeout_seconds, label)
            except TimeoutError as te:
                last_error = te
                logger.warning(
                    "%s timed out after %.1fs; %s",
                    label,
                    timeout_seconds,
                    "retrying" if attempt < attempts else "giving up",
                )
            except Exception as exc:
                last_error = exc
                logger.warning(
                    "%s errored (%s); %s",
                    label,
                    exc,
                    "retrying" if attempt < attempts else "giving up",
                )
        assert last_error is not None
        raise last_error

    def _generate_slide_summaries(self, summary_jobs: List[Dict[str, Any]]) -> Dict[int, str]:
        if not summary_jobs:
            return {}

        max_workers = self._summary_max_workers(summary_jobs)
        model_type = self._resolve_summary_model_type(bool(summary_jobs[0].get("is_chinese")))
        jobs_by_slide = {job["slide_idx"]: job for job in summary_jobs}
        results: Dict[int, str] = {}

        logger.info(
            "Generating %s PPTX slide summaries with model_type=%s, max_workers=%s",
            len(summary_jobs),
            model_type,
            max_workers,
        )

        def _generate_summary(job: Dict[str, Any]) -> str:
            slide_number = int(job["slide_idx"]) + 1
            summary_started_at = time.perf_counter()
            ai_summary = self._generate_ai_summary(
                job["page_commentary"] or job["page_summary_source"],
                job["page_summary_source"],
                job["is_chinese"],
            )
            if ai_summary:
                logger.info(
                    "PPTX summary slide %s completed via AI in %.2fs",
                    slide_number,
                    time.perf_counter() - summary_started_at,
                )
                return ai_summary
            fallback_summary = self._generate_page_summary(job["page_summary_source"], job["is_chinese"])
            logger.info(
                "PPTX summary slide %s completed via fallback in %.2fs",
                slide_number,
                time.perf_counter() - summary_started_at,
            )
            return fallback_summary

        if max_workers == 1:
            for slide_idx, job in jobs_by_slide.items():
                results[slide_idx] = _generate_summary(job)
            return results

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_slide = {
                executor.submit(_generate_summary, job): slide_idx
                for slide_idx, job in jobs_by_slide.items()
            }
            for future in as_completed(future_to_slide):
                slide_idx = future_to_slide[future]
                job = jobs_by_slide[slide_idx]
                try:
                    results[slide_idx] = future.result()
                except Exception as exc:
                    logger.warning(
                        "Slide %s summary generation failed, using fallback summary: %s",
                        slide_idx + 1,
                        exc,
                    )
                    results[slide_idx] = self._generate_page_summary(
                        job["page_summary_source"],
                        job["is_chinese"],
                    )

        return results

    def _prepare_commentary_blocks(self, commentary: str) -> List[str]:
        normalized = str(commentary or "").replace("\r\n", "\n").strip()
        if not normalized:
            return []

        blocks: List[str] = []
        for raw_block in re.split(r"\n\s*\n", normalized):
            lines = [line.strip() for line in raw_block.split("\n") if line.strip()]
            if not lines:
                continue
            if len(lines) == 1:
                blocks.append(lines[0])
                continue

            rebuilt: List[str] = []
            current = ""
            for line in lines:
                is_bullet_like = bool(re.match(r"^([-*•]|\d+[.)])\s+", line))
                if is_bullet_like:
                    if current:
                        rebuilt.append(current.strip())
                        current = ""
                    rebuilt.append(line)
                    continue
                current = f"{current} {line}".strip() if current else line
            if current:
                rebuilt.append(current.strip())
            blocks.extend(rebuilt)
        return blocks

    def _compact_commentary_for_ppt(self, commentary: str, is_chinese: bool) -> str:
        normalized = _normalize_slide_commentary_text(commentary)
        if not normalized:
            return ""

        packing = self._packing_settings()
        min_chars = int(
            packing.get("ppt_min_chars_chi" if is_chinese else "ppt_min_chars_eng", 110 if is_chinese else 190)
        )
        if len(normalized) <= min_chars:
            return normalized

        target_ratio = float(packing.get("ppt_length_ratio", 0.72) or 0.72)
        target_chars = max(min_chars, int(len(normalized) * target_ratio))
        max_sentences = int(
            packing.get("ppt_max_sentences_chi" if is_chinese else "ppt_max_sentences_eng", 3)
        )
        max_numeric_sentences = int(packing.get("ppt_max_numeric_sentences", 2) or 2)

        compact = _build_compact_summary_text(
            normalized,
            is_chinese=is_chinese,
            max_sentences=max_sentences,
            max_chars=target_chars,
            max_numeric_sentences=max_numeric_sentences,
        )
        compact = _normalize_slide_commentary_text(compact)
        if not compact:
            return normalized

        minimum_retained_chars = max(90 if is_chinese else 140, int(len(normalized) * 0.35))
        if len(compact) < minimum_retained_chars:
            return normalized
        return compact if len(compact) < len(normalized) else normalized

    def _prepare_structured_data_for_slides(self, structured_data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        prepared: List[Dict[str, Any]] = []
        for account_data in structured_data or []:
            item = dict(account_data or {})
            commentary = _normalize_slide_commentary_text(item.get("commentary", ""))
            if commentary:
                item["original_commentary"] = commentary
            item["commentary"] = commentary  # Keep full length; fill optimizer handles fit
            prepared.append(item)
        return prepared

    # Average rendered character width (pt) for the fonts we use.
    # English: Arial 9pt mixed text ≈ 5.0 pt/char (incl. spaces & punctuation).
    # Chinese: YaHei 10pt CJK characters are square — 1 em ≈ 10 pt/char.
    # A small word-wrap slack (≈8 %) is subtracted because lines always break
    # at a word/character boundary, not at the exact pixel edge.
    _AVG_CHAR_WIDTH_ENG = 5.0
    _AVG_CHAR_WIDTH_CHI = 10.0
    _WORD_WRAP_SLACK    = 0.92   # use 92 % of the theoretical line width

    def _estimate_chars_per_line(
        self,
        slot_name: str,
        is_chinese: bool,
        shape=None,
        *,
        statement_type: Optional[str] = None,
    ) -> int:
        """Return the number of characters that fit on one line.

        When the actual shape is available we measure directly from its width
        and the text-frame insets, then divide by the known average character
        width for the font in use.  This removes all dependency on the
        ``chars_per_line`` config block for shapes that exist in the template.

        Falls back to the config-based estimate only when no shape is supplied.
        """
        if shape is not None and hasattr(shape, "width"):
            width_pt = shape.width * 72 / 914400
            # Read actual text-frame left/right insets; default is 0.1 in = 7.2 pt.
            left_pt = right_pt = 7.2
            try:
                tf = shape.text_frame
                if tf.margin_left is not None:
                    left_pt  = tf.margin_left  * 72 / 914400
                if tf.margin_right is not None:
                    right_pt = tf.margin_right * 72 / 914400
            except Exception:
                pass
            effective_pt = max(10.0, width_pt - left_pt - right_pt)
            avg_char = self._AVG_CHAR_WIDTH_CHI if is_chinese else self._AVG_CHAR_WIDTH_ENG
            return max(16, int(effective_pt * self._WORD_WRAP_SLACK / avg_char))

        # No shape — fall back to config values.
        packing = self._packing_settings(statement_type)
        chars_per_line = packing.get("chars_per_line") or {}
        slot_key = slot_name if slot_name in {"single", "L", "R"} else "default"
        language_key = "chi" if is_chinese else "eng"
        base_value = (
            ((chars_per_line.get(slot_key) or {}).get(language_key))
            or ((chars_per_line.get("default") or {}).get(language_key))
            or (32 if is_chinese else 60)
        )
        return int(base_value)

    @staticmethod
    def _build_page_summary_source(slide_accounts: List[Dict]) -> Tuple[str, str]:
        """Build the exact slide commentary set used for summary generation."""
        commentary_parts = []
        summary_parts = []

        for account_data in slide_accounts or []:
            commentary = str(account_data.get("commentary", "") or "").strip()
            summary = str(account_data.get("summary", "") or "").strip()
            if commentary:
                commentary_parts.append(commentary)
            if summary:
                summary_parts.append(summary)

        page_commentary = "\n\n".join(commentary_parts).strip()
        page_summary_source = page_commentary or " ".join(summary_parts).strip()
        return page_commentary, page_summary_source

    @staticmethod
    def _shape_name(shape) -> str:
        return str(getattr(shape, "name", "") or "")

    @staticmethod
    def _shape_has_table(shape) -> bool:
        try:
            if getattr(shape, "has_table", False):
                return True
        except Exception:
            pass

        try:
            table = getattr(shape, "table", None)
            return table is not None
        except Exception:
            return False

    def _resolve_table_target_shape(self, slide, statement_type: str):
        """Resolve the best existing target for a BS/IS table on a slide."""
        statement_type = (statement_type or "").upper()
        preferred_names = [
            "Table Placeholder",
            "Table Placeholder 2",
            "Content Placeholder 2",
        ]
        if statement_type == "IS":
            preferred_names.extend(["Table 3", "Table 2"])
        preferred_names.extend(["Table", "table", "TABLE"])

        for name in preferred_names:
            shape = self.find_shape_by_name(slide.shapes, name)
            if shape:
                return shape

        named_table_candidates = []
        table_candidates = []
        text_placeholder_candidates = []
        for shape in slide.shapes:
            shape_name = self._shape_name(shape)
            shape_name_lower = shape_name.lower()
            if "table" in shape_name_lower and "placeholder" in shape_name_lower:
                text_placeholder_candidates.append(shape)
                continue
            if self._shape_has_table(shape):
                table_candidates.append(shape)
                continue
            if "table" in shape_name_lower:
                named_table_candidates.append(shape)

        if text_placeholder_candidates:
            return text_placeholder_candidates[0]
        if table_candidates:
            return table_candidates[0]
        if named_table_candidates:
            return named_table_candidates[0]
        return None

    def _calculate_table_bounds(self, slide, target_shape=None, statement_type: str = "BS") -> Dict[str, int]:
        """Use target geometry when available, otherwise derive bounds from slide layout.

        The table top is aligned with the textMainBullets commentary body so the
        financial table and the Commentary blue box sit on the same horizontal
        baseline. The blue "Commentary"/"Table" label boxes (TextBox 10/11 in
        the template) act as headers above this baseline and are not covered
        by the table.
        """
        if target_shape is not None:
            return {
                "left": target_shape.left,
                "top": target_shape.top,
                "width": target_shape.width,
                "height": target_shape.height,
            }

        slide_width = getattr(self.presentation, "slide_width", Inches(10))
        slide_height = getattr(self.presentation, "slide_height", Inches(7.5))

        title_like_shapes = []
        body_like_shapes = []
        subtitle_shapes = []
        label_shapes = []  # Blue "Commentary"/"Table" label boxes — used as baseline anchor.
        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            name = self._shape_name(shape).lower()
            try:
                label_text = (shape.text_frame.text or "").strip().lower()
            except Exception:
                label_text = ""
            if "subtitle" in name:
                subtitle_shapes.append(shape)
                body_like_shapes.append(shape)
            elif any(token in name for token in ["title", "projtitle"]):
                title_like_shapes.append(shape)
            elif any(token in name for token in ["text-commentary", "textmainbullets", "content"]):
                body_like_shapes.append(shape)
            elif label_text in ("commentary", "table"):
                label_shapes.append(shape)

        left = Inches(0.5)
        width = max(Inches(5.5), slide_width - Inches(1.0))
        top = Inches(1.45 if statement_type.upper() == "BS" else 1.6)
        height = slide_height - top - Inches(0.45)

        if title_like_shapes:
            bottom = max(shape.top + shape.height for shape in title_like_shapes)
            top = max(top, bottom + Inches(0.15))

        generic_is_layout = statement_type.upper() == "IS" and subtitle_shapes and target_shape is None
        if generic_is_layout:
            earliest_subtitle_top = min(shape.top for shape in subtitle_shapes)
            left = Inches(0.5)
            top = Inches(1.5)
            width = min(slide_width - left - Inches(0.35), int((slide_width - Inches(1.0)) * 0.5))
            height = max(Inches(2.0), earliest_subtitle_top - top - Inches(0.12))

        # Horizontal extent: anchor width/left to body and label shapes so the
        # table spans the full commentary gutter.
        horizontal_anchors = list(body_like_shapes) + list(label_shapes)
        if horizontal_anchors:
            left = min(left, min(shape.left for shape in horizontal_anchors))
            right_edge = max(shape.left + shape.width for shape in horizontal_anchors)
            if not generic_is_layout:
                width = max(width, right_edge - left)

        # Vertical alignment: anchor the table TOP to the "Commentary" /
        # "Table" blue label box (TextBox 10/11 in the template). This puts
        # the navy title row of the table at the exact same visible level as
        # the Commentary label on the right, replacing the need for a separate
        # "Table" label above the table.
        if label_shapes:
            label_top = min(shape.top for shape in label_shapes)
            top = max(top, label_top)
            if not generic_is_layout:
                height = max(Inches(2.5), slide_height - top - Inches(0.35))
        elif body_like_shapes:
            earliest_body_top = min(shape.top for shape in body_like_shapes)
            if not generic_is_layout:
                # No label shapes (dynamically added slide): fall back to the
                # commentary body top as the anchor.
                top = max(top, earliest_body_top)
                height = max(Inches(2.5), slide_height - top - Inches(0.35))
            else:
                height = min(height, max(Inches(2.0), earliest_body_top - top - Inches(0.12)))

        width = min(width, slide_width - left - Inches(0.25))
        height = max(Inches(2.5), min(height, slide_height - top - Inches(0.2)))
        return {
            "left": int(left),
            "top": int(top),
            "width": int(width),
            "height": int(height),
        }

    def _add_table_to_slide(self, slide, df, bounds: Dict[str, int], table_name: str = None):
        total_rows = len(df) + 2 if table_name else len(df) + 1
        return slide.shapes.add_table(
            total_rows,
            len(df.columns),
            bounds["left"],
            bounds["top"],
            bounds["width"],
            bounds["height"],
        )

    def _fit_table_columns(self, table, df):
        """Allocate width by role and content length for better readability."""
        if len(table.columns) == 0:
            return

        try:
            total_width = sum(col.width for col in table.columns)
        except Exception:
            total_width = 0
        if total_width <= 0:
            return

        weights = []
        for col_idx, col_name in enumerate(df.columns[: len(table.columns)]):
            col_series = df.iloc[:, col_idx].astype(str) if col_idx < len(df.columns) else pd.Series(dtype=str)
            max_len = max([len(str(col_name))] + [len(val) for val in col_series.head(25).tolist()]) if len(col_series) else len(str(col_name))
            col_name_str = str(col_name).lower()
            if col_idx == 0:
                weight = max(2.0, min(3.2, max_len / 10))
            elif any(token in col_name_str for token in ["20", "19", "date", "年", "月"]):
                weight = max(1.4, min(2.0, max_len / 10))
            else:
                weight = max(1.2, min(1.9, max_len / 9))
            weights.append(weight)

        total_weight = sum(weights) or 1
        assigned = 0
        for col_idx, weight in enumerate(weights):
            if col_idx == len(weights) - 1:
                width = total_width - assigned
            else:
                width = int(total_width * weight / total_weight)
                assigned += width
            table.columns[col_idx].width = max(int(Inches(0.7)), width)

    @staticmethod
    def _format_table_value(value, is_numeric_column: bool) -> str:
        def _fmt_number(n: float) -> str:
            if n == 0:
                return "-"
            # Accounting convention: negatives in parentheses, not with a minus sign.
            return f"({abs(n):,.0f})" if n < 0 else f"{n:,.0f}"

        if pd.isna(value):
            return ""
        if isinstance(value, (int, float)) and is_numeric_column:
            return _fmt_number(float(value))

        text_val = str(value).strip()
        if is_numeric_column:
            numeric_candidate = text_val.replace(",", "")
            if re.fullmatch(r"-?\d+(\.\d+)?", numeric_candidate):
                try:
                    return _fmt_number(float(numeric_candidate))
                except Exception:
                    return text_val
        return text_val

    def _embed_statement_table(self, slide, df, statement_type: str, table_name: str = None, currency_unit: str = None):
        target_shape = self._resolve_table_target_shape(slide, statement_type)
        bounds = self._calculate_table_bounds(slide, target_shape=target_shape, statement_type=statement_type)
        target_name = self._shape_name(target_shape) if target_shape is not None else "(new table)"
        logger.info(
            f"Resolved {statement_type} table target on slide using {target_name} "
            f"at left={bounds['left']} top={bounds['top']} width={bounds['width']} height={bounds['height']}"
        )

        # Remove the redundant "Table" label box (TextBox 11 in the template).
        # The table's navy title row now lives at the same vertical position
        # as that label, so keeping the label would double-print the header.
        # Leave the right-side "Commentary" label intact — there's still a
        # commentary box on this slide that needs its header.
        for shape in list(slide.shapes):
            if not getattr(shape, "has_text_frame", False):
                continue
            try:
                label_text = (shape.text_frame.text or "").strip().lower()
            except Exception:
                continue
            if label_text != "table":
                continue
            try:
                sp = shape._element
                sp.getparent().remove(sp)
            except Exception as e:
                logger.debug("Could not remove 'Table' label shape: %s", e)

        if target_shape is None:
            table_shape = self._add_table_to_slide(slide, df, bounds, table_name=table_name)
            self._fill_table_placeholder(
                table_shape,
                df,
                table_name=table_name,
                currency_unit=currency_unit,
                bounds=bounds,
            )
            return

        self._fill_table_placeholder(
            target_shape,
            df,
            table_name=table_name,
            currency_unit=currency_unit,
            bounds=bounds,
        )
    
    def find_content_shape(self, shapes):
        """Find content shape by trying multiple possible names"""
        # Try different possible names for content shapes
        possible_names = [
            'Content',
            'Text-commentary',
            'textMainBullets',
            'Text',
            'Commentary',
            'MainContent',
            'Body'
        ]
        
        for name in possible_names:
            shape = self.find_shape_by_name(shapes, name)
            if shape and shape.has_text_frame:
                return shape
        
        # If no named shape found, try to find any text frame shape that's not a title
        for shape in shapes:
            if hasattr(shape, 'has_text_frame') and shape.has_text_frame:
                shape_name = getattr(shape, 'name', '')
                # Skip title shapes and other non-content shapes
                if shape_name and 'title' not in shape_name.lower() and 'proj' not in shape_name.lower():
                    return shape
        
        return None

    def replace_text_preserve_formatting(self, shape, replacements: Dict[str, str]) -> bool:
        """Replace text while preserving formatting"""
        if not shape.has_text_frame:
            return False

        replaced = False

        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                original_text = run.text
                for old_text, new_text in replacements.items():
                    if old_text in run.text:
                        run.text = run.text.replace(old_text, new_text)
                if run.text != original_text:
                    replaced = True

        if not replaced:
            current_text = shape.text_frame.text
            updated_text = current_text
            for old_text, new_text in replacements.items():
                updated_text = updated_text.replace(old_text, new_text)
            if updated_text != current_text:
                shape.text_frame.text = updated_text
                replaced = True

        return replaced

    def refresh_project_placeholders(self, project_name: str):
        """Refresh placeholder tokens such as [PROJECT], [Current], and [Total]."""
        if not self.presentation or not project_name:
            return

        display_entity = str(project_name).strip()
        total_slides = len(self.presentation.slides)
        if not display_entity or total_slides <= 0:
            return

        base_replacements = {
            "[PROJECT]": display_entity,
            "[Entity]": display_entity,
            "[Company]": display_entity,
        }

        for slide_index, slide in enumerate(self.presentation.slides):
            proj_title_shape = self.find_shape_by_name(slide.shapes, "projTitle")
            if not proj_title_shape or not proj_title_shape.has_text_frame:
                continue

            replacements = dict(base_replacements)
            replacements["[Current]"] = str(slide_index + 1)
            replacements["[Total]"] = str(total_slides)
            current_text = proj_title_shape.text

            if any(token in current_text for token in replacements):
                self.replace_text_preserve_formatting(proj_title_shape, replacements)

    def update_project_titles(self, project_name: str, statement_type: str = 'BS'):
        """Update project titles in presentation"""
        if not self.presentation:
            return

        display_entity = str(project_name or "").strip()
        self.refresh_project_placeholders(display_entity)

        # Define title templates based on language and statement type
        if statement_type.upper() == 'BS':
            if self.language == 'chinese':
                title_template = f"资产负债表概览 - {display_entity}"
            else:
                title_template = f"Entity Overview - {display_entity}"
        elif statement_type.upper() == 'IS':
            if self.language == 'chinese':
                title_template = f"利润表概览 - {display_entity}"
            else:
                title_template = f"Income Statement - {display_entity}"
        else:
            if self.language == 'chinese':
                title_template = f"财务报表概览 - {display_entity}"
            else:
                title_template = f"Financial Report - {display_entity}"

        # Update titles in all slides
        for slide_index, slide in enumerate(self.presentation.slides):
            current_slide_number = slide_index + 1
            proj_title_shape = self.find_shape_by_name(slide.shapes, "projTitle")

            if proj_title_shape:
                current_text = proj_title_shape.text
                if "[PROJECT]" in current_text:
                    replacements = {
                        "[PROJECT]": display_entity,
                        "[Current]": str(current_slide_number),
                        "[Total]": str(len(self.presentation.slides))
                    }
                    self.replace_text_preserve_formatting(proj_title_shape, replacements)
                else:
                    # Replace the entire title
                    if proj_title_shape.has_text_frame:
                        proj_title_shape.text_frame.text = title_template

    def generate_full_report(self, markdown_content: str, summary_md: Optional[str] = None,
                           output_path: str = None):
        """Generate full PowerPoint report from markdown content"""
        if not self.presentation:
            self.load_template()

        # Process markdown content
        processed_content = self._process_markdown_content(markdown_content)

        # Apply content to presentation
        self._apply_content_to_presentation(processed_content)

        # Save if output path provided
        if output_path:
            self.save(output_path)

    def _process_markdown_content(self, content: str) -> Dict:
        """Process markdown content into structured data"""
        if not content:
            logger.warning("Empty content provided to _process_markdown_content")
            return {}

        logger.info("Processing markdown content, length: %s", len(content))
        logger.debug("Content preview (first 500 chars): %s", content[:500])

        # Split by headers (## Account Name)
        sections = re.split(r'^##\s+(.+)$', content, flags=re.MULTILINE)

        logger.info("Found %s sections after splitting", len(sections))

        processed_sections = {}

        # Process each section
        for i in range(1, len(sections), 2):
            if i + 1 < len(sections):
                account_name = sections[i].strip()
                account_content = sections[i + 1].strip()

                logger.info("Processing section: %s, content length: %s", account_name, len(account_content))

                processed_sections[account_name] = {
                    'content': account_content,
                    'is_chinese': detect_chinese_text(account_content)
                }

        logger.info("Processed %s sections", len(processed_sections))
        return processed_sections

    def _apply_content_to_presentation(self, sections: Dict):
        """Apply processed content to presentation slides"""
        if not self.presentation:
            logger.warning("No presentation loaded")
            return

        logger.info("Applying %s sections to presentation with %s slides", len(sections), len(self.presentation.slides))

        # Find content placeholders and fill them
        slide_idx = 0
        for slide in self.presentation.slides:
            if slide_idx >= len(sections):
                logger.warning("More slides (%s) than sections (%s)", len(self.presentation.slides), len(sections))
                break

            account_name = list(sections.keys())[slide_idx]
            section_data = sections[account_name]

            logger.info("Processing slide %s for account: %s", slide_idx + 1, account_name)

            # Find content shape using flexible name matching
            content_shape = self.find_content_shape(slide.shapes)
            if content_shape:
                logger.info("Found content shape '%s' on slide %s", content_shape.name, slide_idx + 1)
                if content_shape.has_text_frame:
                    # Apply content to shape
                    self._fill_content_shape(content_shape, section_data)
                    logger.info("Applied content to slide %s", slide_idx + 1)
                else:
                    logger.warning("Content shape found but has no text_frame on slide %s", slide_idx + 1)
            else:
                logger.warning("No content shape found on slide %s, available shapes: %s", slide_idx + 1, [s.name if hasattr(s, 'name') else 'unnamed' for s in slide.shapes])
                # Try to use the first available text frame as fallback
                for shape in slide.shapes:
                    if hasattr(shape, 'has_text_frame') and shape.has_text_frame:
                        shape_name = getattr(shape, 'name', 'unnamed')
                        if 'title' not in shape_name.lower() and 'proj' not in shape_name.lower():
                            logger.info("Using fallback shape '%s' on slide %s", shape_name, slide_idx + 1)
                            self._fill_content_shape(shape, section_data)
                            break

            slide_idx += 1

    def _fill_content_shape(self, shape, section_data: Dict):
        """Fill content shape with processed data"""
        if not shape.has_text_frame:
            logger.warning("Shape does not have text_frame")
            return

        content = section_data.get('content', '')
        is_chinese = section_data.get('is_chinese', False)

        logger.info("Filling shape with content length: %s", len(content))

        # Clear existing content
        shape.text_frame.clear()
        
        if not content or not content.strip():
            logger.warning("No content to fill")
            return
        
        # Split content into paragraphs if it contains newlines
        content_lines = content.split('\n')
        
        # Add content with proper formatting
        for idx, line in enumerate(content_lines):
            line = line.strip()
            if not line and idx > 0:
                # Skip empty lines except add a paragraph break
                continue
            
            if idx == 0:
                # Use first paragraph or create one
                if shape.text_frame.paragraphs:
                    p = shape.text_frame.paragraphs[0]
                else:
                    p = shape.text_frame.add_paragraph()
            else:
                p = shape.text_frame.add_paragraph()
            
            p.text = line
            
            # Apply formatting to runs
            for run in p.runs:
                run.font.size = get_font_size_for_text(line, force_chinese_mode=is_chinese)
                run.font.name = get_font_name_for_text(line)

            # Set paragraph formatting
            p.space_after = get_space_after_for_text(line, force_chinese_mode=is_chinese)
            p.space_before = get_space_before_for_text(line, force_chinese_mode=is_chinese)
            p.line_spacing = get_line_spacing_for_text(line, force_chinese_mode=is_chinese)
        
        logger.info("Successfully filled shape with %s paragraphs", len([l for l in content_lines if l.strip()]))

    def _pillow_fitting_enabled(self, packing: Dict[str, Any]) -> bool:
        if os.environ.get("FDD_USE_PILLOW_FITTING") == "1":
            return True
        if os.environ.get("FDD_USE_PILLOW_FITTING") == "0":
            return False
        return bool(packing.get("use_pillow_text_fitting", False))

    def _pillow_measure(
        self,
        shape,
        text: str,
        *,
        is_chinese: bool,
    ) -> Optional[Tuple[int, int]]:
        """Returns (used_lines, capacity_lines) using real font metrics, or
        None on any failure (caller falls back to legacy CPL heuristic)."""
        if not shape or not hasattr(shape, "height") or not hasattr(shape, "width"):
            return None
        try:
            from fdd_utils.text_metrics import (
                get_font,
                line_height_pt as _line_h,
                lines_that_fit,
                text_box_from_shape,
                wrap_text,
            )
        except Exception:
            return None
        try:
            font_size_pt = 10 if is_chinese else 9
            family = "Microsoft YaHei" if is_chinese else "Arial"
            line_spacing = 0.95 if is_chinese else 1.0
            font = get_font(family, font_size_pt, is_cjk=is_chinese)
            box = text_box_from_shape(shape)
            line_h = _line_h(font, line_spacing=line_spacing)
            capacity = lines_that_fit(box.height_pt, line_h)
            if not text:
                return (0, capacity)
            lines = wrap_text(text, font, box.width_pt)
            return (len(lines), capacity)
        except Exception:
            return None

    # Space-after (pt) applied to every paragraph in _fill_text_main_bullets.
    # Used consistently in both capacity and content-line calculations.
    # Vertical gap (in pt) between bullet paragraphs. Used in both the
    # line-cost estimator and the render step so they stay consistent.
    # Tighter spacing = more accounts fit per slot before overflow.
    _PARA_SPACE_AFTER = 3.0

    def _calculate_max_lines_for_textbox(
        self,
        shape,
        *,
        is_chinese: bool = False,
        slot_name: str = "single",
        statement_type: Optional[str] = None,
    ):
        """Return the number of 'line units' that fit in this text box.

        Measures the effective height directly from the shape and its
        text-frame insets (top/bottom margins), then divides by the standard
        line height used by ``_calculate_content_lines``:

            std_line_h = font_size × line_spacing + PARA_SPACE_AFTER

        Both capacity and content are expressed in the same unit so the fill
        ratios are accurate without any fudge factors.
        """
        packing = self._packing_settings(statement_type)
        if not shape or not hasattr(shape, "height"):
            return int(packing.get("minimum_slot_lines", 20) or 20)

        font_size_pt = 10 if is_chinese else 9
        line_spacing = 0.95 if is_chinese else 1.0
        family       = "Microsoft YaHei" if is_chinese else "Arial"

        # ── Real font metrics via text_metrics ───────────────────────────────────
        # text_box_from_shape reads bodyPr tIns/bIns directly from shape XML.
        # line_height_pt uses ascent + descent from the actual font file.
        try:
            from fdd_utils.text_metrics import (
                get_font, line_height_pt as _line_h_fn, text_box_from_shape,
            )
            font     = get_font(family, font_size_pt, is_cjk=is_chinese)
            box      = text_box_from_shape(shape)
            line_h   = _line_h_fn(font, line_spacing=line_spacing)
            std_lh   = line_h + self._PARA_SPACE_AFTER
            max_rows = int(box.height_pt / std_lh)
            # Trust the real measurement — do NOT apply minimum_slot_lines as a
            # floor here, because that would allow the DP to pack more content
            # (max_rows × std_lh pt) than the box can physically hold.
            return max(1, max_rows)
        except Exception:
            pass   # font file missing — fall through to heuristic

        # ── Heuristic fallback ───────────────────────────────────────────────────
        height_pt    = shape.height * 72 / 914400
        top_pt = bottom_pt = 3.6          # OOXML default tIns/bIns = 0.05" = 3.6 pt
        try:
            tf = shape.text_frame
            if tf.margin_top    is not None: top_pt    = tf.margin_top    * 72 / 914400
            if tf.margin_bottom is not None: bottom_pt = tf.margin_bottom * 72 / 914400
        except Exception:
            pass
        effective_pt = max(1.0, height_pt - top_pt - bottom_pt)
        std_lh       = font_size_pt * line_spacing + self._PARA_SPACE_AFTER
        max_rows     = int(effective_pt / std_lh)
        return max(int(packing.get("minimum_slot_lines", 20) or 20), max_rows)

    def _calculate_content_lines(
        self,
        category: str,
        mapping_key: str,
        commentary: str,
        *,
        slot_name: str = "single",
        shape=None,
        is_chinese: Optional[bool] = None,
        statement_type: Optional[str] = None,
    ) -> float:
        """Return the physical height of this content expressed in std_lh units.

        Returns a *float* (no ceil) so that the DP and greedy fill can track
        actual physical space consumed.  Using ceil was inflating every
        multi-line account to the next integer boundary, causing the DP to
        report 100 % fill when the box was only ~75 % physically used.

        One "unit" = std_lh = line_h + PARA_SPACE_AFTER (17 pt for English).
        Capacity from _calculate_max_lines_for_textbox is int(box_height/std_lh),
        so comparing float content against int capacity gives physically accurate
        fill ratios.
        """
        is_chinese = any('\u4e00' <= c <= '\u9fff' for c in commentary) if is_chinese is None else is_chinese

        font_size_pt = 10 if is_chinese else 9
        line_spacing = 0.95 if is_chinese else 1.0
        family = "Microsoft YaHei" if is_chinese else "Arial"

        # ── Real glyph metrics via text_metrics ─────────────────────────────────
        if shape is not None:
            try:
                from fdd_utils.text_metrics import (
                    get_font, line_height_pt as _line_h_fn,
                    text_box_from_shape, wrap_paragraph,
                )
                font     = get_font(family, font_size_pt, is_cjk=is_chinese)
                box      = text_box_from_shape(shape)
                line_h   = _line_h_fn(font, line_spacing=line_spacing)
                std_lh   = line_h + self._PARA_SPACE_AFTER

                total_pt = 0.0
                if category:
                    total_pt += line_h          # category header: no space_after

                paras = [p for p in commentary.split('\n') if p.strip()] if commentary else []
                key_prefix = f"\u25a0 {mapping_key} - "
                if paras:
                    first_wrapped = wrap_paragraph(key_prefix + paras[0], font, box.width_pt)
                    total_pt += len(first_wrapped) * line_h + self._PARA_SPACE_AFTER
                    for para in paras[1:]:
                        wrapped = wrap_paragraph(para, font, box.width_pt)
                        total_pt += len(wrapped) * line_h + self._PARA_SPACE_AFTER
                else:
                    total_pt += line_h + self._PARA_SPACE_AFTER

                # Return float — no ceil so actual physical proportion is preserved.
                return total_pt / std_lh
            except Exception:
                pass    # font file missing — fall through to heuristic

        # ── Heuristic fallback (no shape or font unavailable) ───────────────────
        space_after  = self._PARA_SPACE_AFTER
        std_lh       = font_size_pt * line_spacing + space_after
        cpl          = self._estimate_chars_per_line(slot_name, is_chinese, shape=shape,
                                                     statement_type=statement_type)
        total_pt     = 0.0
        if category:
            total_pt += font_size_pt * line_spacing
        paras = [p for p in commentary.split('\n') if p.strip()] if commentary else []
        key_pfx_len  = len(str(mapping_key)) + 5
        if paras:
            first_len    = key_pfx_len + len(paras[0])
            first_w      = max(1, (first_len + cpl - 1) // cpl)
            total_pt    += first_w * font_size_pt * line_spacing + space_after
            for para in paras[1:]:
                w         = max(1, (len(para) + cpl - 1) // cpl)
                total_pt += w * font_size_pt * line_spacing + space_after
        else:
            total_pt += font_size_pt * line_spacing + space_after
        return total_pt / std_lh

    def _distribute_content_across_slots(
        self,
        structured_data: List[Dict],
        max_slides: int = 4,
        start_slide: int = 1,
        statement_type: Optional[str] = None,
    ):
        """
        Distribute content across textbox slots based on capacity.
        Slot structure is derived from the actual template slides when they exist.
        Auto-added slides follow the template convention: page 1 keeps a single right-side
        commentary box beside the table, while later slides use left/right commentary slots.
        
        Returns: List of (slide_idx, slot_idx, [account_data], is_partial, continuation_of)
        """
        if not structured_data:
            return []

        # Find a textbox shape to calculate capacity
        sample_shape = None
        for slide in self.presentation.slides:
            for alt_name in ["textMainBullets", "textMainBullets_L", "textMainBullets_R"]:
                shape = self.find_shape_by_name(slide.shapes, alt_name)
                if shape:
                    sample_shape = shape
                    break
            if sample_shape:
                break
        max_lines_per_textbox = (
            self._calculate_max_lines_for_textbox(sample_shape, statement_type=statement_type)
            if sample_shape
            else 40
        )
        
        logger.info("\n%s", '='*80)
        logger.info("CONTENT DISTRIBUTION STARTING")
        logger.info("%s", '='*80)
        logger.info("Total accounts: %s", len(structured_data))
        logger.info("Max lines per textbox: %s", max_lines_per_textbox)
        if sample_shape:
            logger.info("Sample shape height: %.2f inches", sample_shape.height / 914400)
            logger.info("Estimated capacity: %s lines", max_lines_per_textbox)
        logger.info("%s\n", '='*80)
        
        def slot_names_for_actual_slide(actual_slide_idx: int) -> List[str]:
            if 0 <= actual_slide_idx < len(self.presentation.slides):
                slide = self.presentation.slides[actual_slide_idx]
                has_left = self.find_shape_by_name(slide.shapes, "textMainBullets_L") is not None
                has_right = self.find_shape_by_name(slide.shapes, "textMainBullets_R") is not None
                has_single = self.find_shape_by_name(slide.shapes, "textMainBullets") is not None
                if has_left and has_right:
                    return ["L", "R"]
                if has_single:
                    return ["single"]
            return ["single"] if actual_slide_idx == 0 else ["L", "R"]

        # Define slot structure: (slide_idx, slot_name)
        slots: List[Tuple[int, str]] = []
        for slide_idx in range(max_slides):
            actual_slide_idx = start_slide - 1 + slide_idx
            for slot_name in slot_names_for_actual_slide(actual_slide_idx):
                slots.append((slide_idx, slot_name))

        slot_shapes: Dict[int, Any] = {}
        for slot_idx, (slide_idx, slot_name) in enumerate(slots):
            actual_slide_idx = start_slide - 1 + slide_idx
            slot_shape = None
            if 0 <= actual_slide_idx < len(self.presentation.slides):
                slide = self.presentation.slides[actual_slide_idx]
                slot_shape = self._resolve_commentary_slot_shape(slide, slot_name)
            slot_shapes[slot_idx] = slot_shape or sample_shape

        logger.info("Total slots available: %s", len(slots))
        
        # Distribution result: [(slide_idx, slot_name, [account_data])]
        distribution = []
        
        current_slot_idx = 0
        current_slot_content = []
        current_slot_lines = 0
        previous_category = None

        def slot_capacity_for(slot_idx: int, *, is_chinese: bool, slot_name_override: Optional[str] = None) -> int:
            _slide_idx, derived_slot_name = slots[slot_idx]
            slot_name_local = slot_name_override or derived_slot_name
            slot_shape_local = slot_shapes.get(slot_idx)
            capacity = self._calculate_max_lines_for_textbox(
                slot_shape_local,
                is_chinese=is_chinese,
                slot_name=slot_name_local,
                statement_type=statement_type,
            )
            if slot_name_local == 'L':
                capacity = int(capacity * 0.98)
            return capacity
        
        for account_idx, account_data in enumerate(structured_data):
            mapping_key_debug = account_data.get('mapping_key', account_data.get('account_name', ''))
            logger.info("\nAccount %s/%s: %s", account_idx + 1, len(structured_data), mapping_key_debug)
            if current_slot_idx >= len(slots):
                dropped_accounts = len(structured_data) - account_idx
                logger.warning(
                    "Ran out of commentary slots; dropping %s remaining account(s) starting from '%s'",
                    dropped_accounts,
                    mapping_key_debug,
                )
                break
            
            category = account_data.get('category', '')
            mapping_key = account_data.get('mapping_key', account_data.get('account_name', ''))
            commentary = account_data.get('commentary', '')

            slide_idx_check, slot_name_check = slots[current_slot_idx]
            is_chinese_content = any('\u4e00' <= c <= '\u9fff' for c in commentary)
            chars_setting = 35 if is_chinese_content else 70
            category_lines = 1 if (category and category != previous_category) else 0
            content_lines = self._calculate_content_lines(
                '',
                mapping_key,
                commentary,
                slot_name=slot_name_check,
                shape=slot_shapes.get(current_slot_idx),
                is_chinese=is_chinese_content,
                statement_type=statement_type,
            )
            total_lines = category_lines + content_lines
            logger.info("  Category: '%s', Lines: cat=%s, content=%s, total=%s", category, category_lines, content_lines, total_lines)
            logger.info("  Commentary length: %s chars, Language: %s, Chars/line: %s", len(commentary), 'Chinese' if is_chinese_content else 'English', chars_setting)

            adjusted_capacity = slot_capacity_for(current_slot_idx, is_chinese=is_chinese_content, slot_name_override=slot_name_check)
            logger.info("  Current slot %s (%s): %s/%s lines used", current_slot_idx, slot_name_check, current_slot_lines, adjusted_capacity)

            if current_slot_lines + total_lines <= adjusted_capacity:
                current_slot_content.append(account_data)
                current_slot_lines += total_lines
                previous_category = category
                logger.info("  Slot %s (%s): Added '%s' (%s lines), now %s/%s lines used", current_slot_idx, slot_name_check, mapping_key, total_lines, current_slot_lines, adjusted_capacity)
            else:
                remaining_lines = adjusted_capacity - current_slot_lines
                logger.info("  Doesn't fit. Remaining: %s lines, Content: %s lines", remaining_lines, content_lines)

                next_slot_idx = current_slot_idx + 1

                split_remaining_min = float(self._packing_settings(statement_type).get("split_min_remaining_lines", 3))
                split_content_min = int(self._packing_settings(statement_type).get("split_min_content_lines", 5))
                if remaining_lines > split_remaining_min and content_lines > split_content_min:
                    logger.info("  Attempting to split content...")
                    paragraphs = commentary.split('\n\n')
                    if len(paragraphs) == 1:
                        paragraphs = commentary.split('\n')

                    chars_per_line = self._estimate_chars_per_line(
                        slot_name_check,
                        is_chinese_content,
                        shape=slot_shapes.get(current_slot_idx),
                        statement_type=statement_type,
                    )
                    available_for_commentary = remaining_lines - category_lines - 1

                    # Convert float line-units to visual display lines for the
                    # paragraph-fitting loop below.  available_for_commentary is
                    # in "std_lh units" (one unit = line_h + space_after pt),
                    # but para_lines is computed via chars_per_line and counts
                    # visual display lines.  Multiply by (std_lh / line_h) so
                    # both are in the same unit.
                    #   English: std_lh=9+6=15pt, line_h=9pt  → factor ≈ 1.667
                    #   Chinese: std_lh=9.5+6=15.5pt, line_h=9.5pt → factor ≈ 1.632
                    _lh_est = (10 * 0.95) if is_chinese_content else (9 * 1.0)
                    _std_lh_est = _lh_est + self._PARA_SPACE_AFTER
                    available_visual = available_for_commentary * (_std_lh_est / _lh_est)

                    if available_for_commentary > 0:
                        part1_paragraphs = []
                        part1_lines_used = 0
                        split_index = 0

                        for i, para in enumerate(paragraphs):
                            para_lines = max(1, (len(para) + chars_per_line - 1) // chars_per_line)
                            if part1_lines_used + para_lines <= available_visual:
                                part1_paragraphs.append(para)
                                part1_lines_used += para_lines
                                split_index = i + 1
                            else:
                                break

                        if part1_paragraphs and split_index < len(paragraphs):
                            # Clean paragraph-boundary split — always safe.
                            part1_commentary = '\n\n'.join(part1_paragraphs).strip()
                            part2_commentary = '\n\n'.join(paragraphs[split_index:]).strip()
                        elif not part1_paragraphs and len(paragraphs) > 0:
                            para = paragraphs[0]
                            chars_available = int(max(1, available_visual * chars_per_line))

                            if len(para) > chars_available:
                                # Only split at SENTENCE boundaries. The older
                                # code fell back to commas, word boundaries,
                                # and hard character cuts — those produced
                                # ugly "(cont'd)" fragments that ended mid-row
                                # at ~30% of the line. If no sentence ending
                                # lands in our window, we refuse to split and
                                # let the account move whole to the next slot.
                                split_point = chars_available
                                best_split = None
                                window_back = max(split_point - 300, 0)
                                window_fwd = min(len(para), split_point + 150)
                                for end_char in [
                                    '. ', '。', '! ', '！', '? ', '？',
                                    '。 ', '！ ', '？ ',
                                ]:
                                    pos = para.rfind(end_char, window_back, window_fwd)
                                    if pos > 0:
                                        best_split = pos + len(end_char)
                                        break

                                if best_split is None:
                                    # No sentence boundary in window — skip
                                    # this split attempt; the code below
                                    # will move the whole account to the
                                    # next slot.
                                    part1_commentary = None
                                    part2_commentary = None
                                else:
                                    part1_commentary = para[:best_split].strip()
                                    remaining_para = para[best_split:].strip()
                                    if len(paragraphs) > 1:
                                        part2_commentary = remaining_para + '\n\n' + '\n\n'.join(paragraphs[1:])
                                    else:
                                        part2_commentary = remaining_para
                            else:
                                part1_commentary = para
                                part2_commentary = '\n\n'.join(paragraphs[1:]) if len(paragraphs) > 1 else ""
                        else:
                            part1_commentary = commentary
                            part2_commentary = None

                        if part1_commentary and part2_commentary:
                            account_part1 = account_data.copy()
                            account_part1['commentary'] = part1_commentary
                            account_part1['is_partial'] = True
                            account_part1['part_num'] = 1
                            current_slot_content.append(account_part1)
                            
                            # Save current slot
                            slide_idx, slot_name = slots[current_slot_idx]
                            distribution.append((slide_idx, slot_name, current_slot_content))
                            logger.info("Split '%s': Part 1 (%s chars) to slot %s, Part 2 (%s chars) to next slot", mapping_key, len(part1_commentary), current_slot_idx, len(part2_commentary))

                            if current_slot_idx + 1 >= len(slots):
                                logger.warning(
                                    "Ran out of commentary slots after splitting '%s'; dropping the remaining continuation",
                                    mapping_key,
                                )
                                break
                            current_slot_idx += 1

                            account_part2 = account_data.copy()
                            account_part2['commentary'] = part2_commentary
                            account_part2['is_continuation'] = True
                            account_part2['part_num'] = 2
                            account_part2['original_key'] = mapping_key

                            next_slot_name = slots[current_slot_idx][1]
                            part2_lines = self._calculate_content_lines(
                                '',
                                mapping_key,
                                part2_commentary,
                                slot_name=next_slot_name,
                                shape=slot_shapes.get(current_slot_idx),
                                is_chinese=is_chinese_content,
                            )
                            current_slot_content = [account_part2]
                            current_slot_lines = part2_lines
                            previous_category = None
                            continue
                else:
                    logger.info("  Not splitting: remaining_lines=%s, content_lines=%s", remaining_lines, content_lines)

                if current_slot_content:
                    slide_idx, slot_name = slots[current_slot_idx]
                    distribution.append((slide_idx, slot_name, current_slot_content))
                    logger.info("  Slot %s (%s): FULL with %s accounts, %s lines used", current_slot_idx, slot_name, len(current_slot_content), current_slot_lines)

                current_slot_idx += 1
                if current_slot_idx >= len(slots):
                    logger.warning(
                        "Ran out of commentary slots while placing '%s'; dropping that account from the remaining slides",
                        mapping_key,
                    )
                    break

                slide_idx_new, slot_name_new = slots[current_slot_idx]
                moved_account = account_data.copy()
                moved_account["commentary"] = commentary
                moved_category_lines = 1 if category else 0
                moved_lines = self._calculate_content_lines(
                    '',
                    mapping_key,
                    moved_account["commentary"],
                    slot_name=slot_name_new,
                    shape=slot_shapes.get(current_slot_idx),
                    is_chinese=is_chinese_content,
                    statement_type=statement_type,
                )
                current_slot_content = [moved_account]
                current_slot_lines = moved_category_lines + moved_lines
                previous_category = category
                logger.info("  Moving '%s' to next slot %s (%s), %s lines", mapping_key, current_slot_idx, slot_name_new, current_slot_lines)
        
        # Save last slot if it has content
        if current_slot_content and current_slot_idx < len(slots):
            slide_idx, slot_name = slots[current_slot_idx]
            distribution.append((slide_idx, slot_name, current_slot_content))
            logger.info("  Slot %s (%s): FINAL with %s accounts, %s lines", current_slot_idx, slot_name, len(current_slot_content), current_slot_lines)
        
        slot_position_map = {slot: idx for idx, slot in enumerate(slots)}

        logger.info("\nDistribution complete: %s slots filled", len(distribution))
        for distribution_idx, (slide_idx, slot_name, accounts) in enumerate(distribution):
            slot_idx = slot_position_map.get((slide_idx, slot_name), distribution_idx)
            account_is_chinese = any(bool(account.get("is_chinese", False)) for account in accounts)
            slot_capacity = slot_capacity_for(slot_idx, is_chinese=account_is_chinese, slot_name_override=slot_name)
            used_lines = 0
            previous_category_in_slot = None
            for account in accounts:
                account_category = account.get("category", "")
                category_cost = 1 if (account_category and account_category != previous_category_in_slot) else 0
                used_lines += category_cost
                used_lines += self._calculate_content_lines(
                    "",
                    account.get("mapping_key", account.get("account_name", "")),
                    account.get("commentary", ""),
                    slot_name=slot_name,
                    shape=slot_shapes.get(slot_idx),
                    is_chinese=account.get("is_chinese", False),
                    statement_type=statement_type,
                )
                previous_category_in_slot = account_category
            fill_ratio = (used_lines / slot_capacity) if slot_capacity else 0.0
            logger.info(
                "  Slot %s (Slide %s, %s): %s accounts, est. fill %.2f",
                slot_idx,
                slide_idx,
                slot_name,
                len(accounts),
                fill_ratio,
            )
        
        # --- Fill optimization pass: pull accounts forward into under-filled slots ---
        distribution = self._optimize_slot_fill(
            distribution,
            slot_shapes=slot_shapes,
            slot_meta=slots,
            statement_type=statement_type,
        )
        return distribution

    def _compute_slot_used_lines(
        self,
        accounts: List[Dict],
        slot_name: str,
        slot_shape=None,
        statement_type: Optional[str] = None,
    ) -> float:
        """Return used line-units for *accounts* in this slot (float).

        Uses the same accounting as ``slot_cost`` in the DP: each category
        header costs 1 line unit, and each account's commentary costs the
        float value returned by ``_calculate_content_lines`` (actual pt /
        std_lh, no ceil).  Comparing against int capacity from
        ``_calculate_max_lines_for_textbox`` gives accurate fill ratios.
        """
        used: float = 0.0
        prev_cat = None
        for account in accounts:
            cat = str(account.get("category", "") or "")
            if cat and cat != prev_cat:
                used += 1.0   # category header (same as slot_cost)
            prev_cat = cat
            used += self._calculate_content_lines(
                "",
                account.get("mapping_key", account.get("account_name", "")),
                account.get("commentary", ""),
                slot_name=slot_name,
                shape=slot_shape,
                is_chinese=bool(account.get("is_chinese", False)),
                statement_type=statement_type,
            )
        return max(0.0, used)

    def _optimize_slot_fill(
        self,
        distribution: List[tuple],
        slot_shapes: Optional[Dict[int, Any]] = None,
        slot_meta: Optional[List[Tuple[int, str]]] = None,
        statement_type: Optional[str] = None,
    ) -> List[tuple]:
        """Dynamic-programming balanced partition.

        Flattens all accounts into reading order, then partitions them into
        contiguous groups (one per slot) so that the maximum slot fill ratio
        is minimised. Line counts come from _compute_slot_used_lines measured
        against each slot's actual shape, so when Pillow fitting is enabled
        this uses real font metrics. Preserves reading order; drops trailing
        empty slots.

        DP: dp[s][i] = min achievable "max fill ratio" when placing
        accounts[0..i] into slots[0..s]. Transition: slot s takes a suffix
        accounts[j+1..i]; combine with dp[s-1][j]. O(S * N^2) states, but
        N ≤ ~20 accounts and S ≤ ~8 slots in practice, so this is trivial.
        """
        if not distribution:
            return distribution

        slot_lookup: Dict[Tuple[int, str], Any] = {}
        if slot_meta and slot_shapes:
            for slot_idx, (s_idx, s_name) in enumerate(slot_meta):
                slot_lookup[(s_idx, s_name)] = slot_shapes.get(slot_idx)

        def _resolve_shape(slide_idx: int, slot_name: str):
            shape = slot_lookup.get((slide_idx, slot_name))
            if shape is not None:
                return shape
            try:
                slide = self.presentation.slides[slide_idx]
            except Exception:
                return None
            return self._resolve_commentary_slot_shape(slide, slot_name)

        flat_accounts: List[Dict[str, Any]] = []
        for _slide_idx, _slot_name, accounts in distribution:
            flat_accounts.extend(accounts)

        if not flat_accounts:
            return distribution

        slots: List[Dict[str, Any]] = []
        is_chinese_any = any(bool(a.get("is_chinese")) for a in flat_accounts)
        for slide_idx, slot_name, _accounts in distribution:
            shape = _resolve_shape(slide_idx, slot_name)
            capacity = self._calculate_max_lines_for_textbox(
                shape,
                is_chinese=is_chinese_any,
                slot_name=slot_name,
                statement_type=statement_type,
            )
            slots.append({
                "slide_idx": slide_idx,
                "slot_name": slot_name,
                "shape": shape,
                "capacity": max(1, int(capacity or 1)),
            })

        N = len(flat_accounts)
        S = len(slots)

        # ── Pre-compute per-account content lines for each unique slot type ──
        # Key: (slot_name, shape_width_emu).  Two slots that share the same
        # name and width get the same measurements, so we only call
        # _calculate_content_lines (and Pillow when enabled) once per
        # (account, slot_type) pair — O(N × slot_types) total instead of the
        # O(S × N²) calls that the old range-based approach produced.
        _acct_cost: Dict[Tuple[int, str, int], float] = {}
        seen_slot_types: set = set()
        for slot in slots:
            shape = slot["shape"]
            w_key = int(shape.width) if shape and hasattr(shape, "width") else 0
            type_key = (slot["slot_name"], w_key)
            if type_key in seen_slot_types:
                continue
            seen_slot_types.add(type_key)
            for a_i, account in enumerate(flat_accounts):
                _acct_cost[(a_i, slot["slot_name"], w_key)] = self._calculate_content_lines(
                    "",
                    account.get("mapping_key", account.get("account_name", "")),
                    account.get("commentary", ""),
                    slot_name=slot["slot_name"],
                    shape=shape,
                    is_chinese=bool(account.get("is_chinese", False)),
                    statement_type=statement_type,
                )

        # ── Tight packing: use minimum slots, expand only if infeasible ─────────
        # The DP objective (min max-fill-ratio) spreads content across ALL
        # available slots at ~80% fill.  We want ~90-95%.  Fix: compute the
        # minimum number of slots that can hold all content, try that first;
        # if infeasible (split accounts can push content above S_min capacity),
        # expand by one and retry until feasible or S_orig is reached.
        _slots_all = list(slots)
        _S_orig = S

        import math as _math
        _est_sname = slots[0]["slot_name"] if slots else "single"
        _est_wkey = (
            int(slots[0]["shape"].width)
            if slots and slots[0].get("shape") and hasattr(slots[0]["shape"], "width")
            else 0
        )
        _total_est: float = 0.0
        _prev_cat_e: Optional[str] = None
        for _a_i, _acct_e in enumerate(flat_accounts):
            _cat_e = str(_acct_e.get("category", "") or "")
            if _cat_e and _cat_e != _prev_cat_e:
                _total_est += 1.0
            _prev_cat_e = _cat_e
            _total_est += _acct_cost.get((_a_i, _est_sname, _est_wkey), 2.0)
        _min_cap = min(slot["capacity"] for slot in slots) if slots else 1
        S_min = max(1, _math.ceil(_total_est / _min_cap))

        # cost_cache and slot_cost are defined before the retry loop.
        # slot_cost captures `slots` by reference — updating slots = _slots_all[:S_try]
        # inside the loop automatically changes what slot_cost sees.
        cost_cache: Dict[Tuple[int, int, int], float] = {}

        # Progressive relax factors for DP feasibility. Start at 1.0 (strict
        # capacity) and widen until the DP finds a partition. The final
        # factor (very large) guarantees feasibility, so the DP ALWAYS
        # returns a balanced result — we never fall through to greedy, which
        # would force-place oversized accounts and break the slide layout.
        # shape_height_utilization is the "natural" first relaxation because
        # PPT auto-fit can absorb that much overflow at render time.
        _packing_relax = self._packing_settings(statement_type)
        _shape_util = float(_packing_relax.get("shape_height_utilization", 1.15) or 1.15)
        _relax_factors: List[float] = [1.0, max(1.05, _shape_util), 1.35, 1.6, 10.0]

        def slot_cost(s: int, j: int, i: int) -> float:
            """Return float line-units for placing flat_accounts[j..i] in slot s.
            Category headers cost 1.0; account content costs the float from
            _calculate_content_lines (actual_pt / std_lh, no ceil)."""
            if j > i:
                return 0.0
            key = (s, j, i)
            if key in cost_cache:
                return cost_cache[key]
            slot = slots[s]
            shape = slot["shape"]
            w_key = int(shape.width) if shape and hasattr(shape, "width") else 0
            sname = slot["slot_name"]
            used: float = 0.0
            prev_cat = None
            for a_i in range(j, i + 1):
                account = flat_accounts[a_i]
                cat = str(account.get("category", "") or "")
                if cat and cat != prev_cat:
                    used += 1.0  # category header line
                prev_cat = cat
                used += _acct_cost.get((a_i, sname, w_key), 0.0)
            cost_cache[key] = used
            return used

        INF = float("inf")
        # Lexicographic DP state: (num_nonempty_slots, max_fill_ratio).
        # Compared as Python tuples, so Python prefers FEWER non-empty slots
        # first, then LOWER max ratio among ties. That's the "pack tight,
        # don't spread" behaviour the user wants — a single full slot beats
        # two half-full slots even though the latter has a lower max ratio.
        INF_ST = (INF, INF)
        dp: List[List[Tuple[float, float]]] = []
        # split[s][i] = j such that slot s holds flat_accounts[j+1..i]; j == i
        # means slot s is empty (carries i through from previous slot).
        split: List[List[int]] = []

        # Progressive relax loop. For each factor, run the DP at full S_orig
        # slots — because the DP's own tight-packing objective will already
        # leave trailing slots empty when the content fits in fewer. The
        # final 10× factor guarantees feasibility, so we never need to fall
        # through to greedy.
        _dp_solved = False
        _solved_factor = 1.0
        S = _S_orig
        slots = _slots_all[:S]
        for _cap_mult in _relax_factors:
            cost_cache.clear()

            dp = [[INF_ST] * N for _ in range(S)]
            split = [[-1] * N for _ in range(S)]

            _cap0 = slots[0]["capacity"] * _cap_mult
            for i in range(N):
                lines = slot_cost(0, 0, i)
                if lines <= _cap0:
                    # slot 0 is non-empty (holds accounts 0..i). Ratio uses
                    # true (unrelaxed) capacity so the tie-break reflects
                    # actual visual fill.
                    dp[0][i] = (1.0, lines / slots[0]["capacity"])
                split[0][i] = -1

            for s in range(1, S):
                cap_true = slots[s]["capacity"]
                cap_check = cap_true * _cap_mult
                for i in range(N):
                    # Case A: slot s non-empty, holds accounts[j+1..i]
                    for j in range(-1, i):
                        if j < 0:
                            prev_state: Tuple[float, float] = (0.0, 0.0)
                        else:
                            prev_state = dp[s - 1][j]
                            if prev_state[0] >= INF:
                                continue
                        lines = slot_cost(s, j + 1, i)
                        if lines > cap_check:
                            continue
                        curr_state = (
                            prev_state[0] + 1.0,
                            max(prev_state[1], lines / cap_true),
                        )
                        if curr_state <= dp[s][i]:
                            dp[s][i] = curr_state
                            split[s][i] = j
                    # Case B: slot s empty — carry dp[s-1][i] forward unchanged
                    if dp[s - 1][i] < dp[s][i]:
                        dp[s][i] = dp[s - 1][i]
                        split[s][i] = i  # marker: slot s is empty

            if dp[S - 1][N - 1][0] < INF:
                _dp_solved = True
                _solved_factor = _cap_mult
                break

            logger.info(
                "  DP infeasible at relax × %.2f; widening tolerance",
                _cap_mult,
            )

        _used_slots = int(dp[S - 1][N - 1][0]) if _dp_solved else S
        _final_ratio = dp[S - 1][N - 1][1] if _dp_solved else INF
        if _solved_factor > 1.0:
            logger.info(
                "  DP feasible with relax × %.2f; using %s of %s slots, max ratio %.2f",
                _solved_factor, _used_slots, _S_orig, _final_ratio,
            )
        else:
            logger.info(
                "  DP tight-pack: using %s of %s slots (min=%s), max ratio %.2f",
                _used_slots, _S_orig, S_min, _final_ratio,
            )

        # Reconstruct the assignment.
        assignment: List[List[Dict[str, Any]]] = [[] for _ in range(S)]
        i = N - 1
        for s in range(S - 1, -1, -1):
            j = split[s][i]
            if j == i:
                assignment[s] = []
                continue
            assignment[s] = list(flat_accounts[j + 1 : i + 1])
            i = j
            if i < 0:
                break

        for s_i, slot in enumerate(slots):
            lines = slot_cost(s_i, 0, -1) if not assignment[s_i] else self._compute_slot_used_lines(
                assignment[s_i],
                slot["slot_name"],
                slot_shape=slot["shape"],
                statement_type=statement_type,
            )
            logger.info(
                "  Balanced DP slot %s (%s): %s/%s lines, accts=%s",
                s_i, slot["slot_name"], lines, slot["capacity"],
                [a.get("mapping_key", "?") for a in assignment[s_i]],
            )

        rebuilt = [
            (slot["slide_idx"], slot["slot_name"], self._merge_contd_pairs(assignment[s_i]))
            for s_i, slot in enumerate(slots)
            if assignment[s_i]
        ]
        return rebuilt

    @staticmethod
    def _merge_contd_pairs(accounts: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Merge any consecutive (part1, cont'd-part2) pair that landed in the
        same slot.  This happens when the DP re-balances: the split was created
        because the *previous* slot was almost full, but the two halves together
        fit in the empty slot the DP chose.  Merging removes the spurious
        (cont'd) label and restores the original single account."""
        result: List[Dict[str, Any]] = []
        skip = False
        for i, acct in enumerate(accounts):
            if skip:
                skip = False
                continue
            nxt = accounts[i + 1] if i + 1 < len(accounts) else None
            if (
                acct.get("is_partial")
                and nxt is not None
                and nxt.get("is_continuation")
                and nxt.get("original_key", nxt.get("mapping_key")) == acct.get("mapping_key")
            ):
                combined = acct.copy()
                p1 = str(acct.get("commentary", "") or "")
                p2 = str(nxt.get("commentary", "") or "")
                combined["commentary"] = (p1.rstrip() + " " + p2.lstrip()).strip()
                for flag in ("is_partial", "is_continuation", "part_num", "original_key"):
                    combined.pop(flag, None)
                result.append(combined)
                skip = True
            else:
                result.append(acct)
        return result

    def _greedy_forward_fill(
        self,
        flat_accounts: List[Dict[str, Any]],
        slots: List[Dict[str, Any]],
        statement_type: Optional[str],
    ) -> List[tuple]:
        """Fallback: fill each slot to capacity greedily. Used only if DP
        can't find a feasible partition (e.g. a single account overflows a
        slot). Always places every account — if an account alone exceeds a
        slot's capacity it is force-placed rather than dropped."""
        def measure(accts, slot):
            return self._compute_slot_used_lines(
                accts, slot["slot_name"], slot_shape=slot["shape"],
                statement_type=statement_type,
            )

        idx = 0
        assignment: List[List[Dict[str, Any]]] = [[] for _ in slots]
        for s_i, slot in enumerate(slots):
            while idx < len(flat_accounts):
                trial = assignment[s_i] + [flat_accounts[idx]]
                if measure(trial, slot) > slot["capacity"] and assignment[s_i]:
                    # Slot already has content and adding this account overflows — move on
                    break
                # Place the account: either the slot is empty (force-place to avoid
                # dropping) or it still fits within capacity.
                assignment[s_i] = trial
                idx += 1

        # If any accounts are still unplaced (more accounts than slots can absorb),
        # append them to the last slot rather than silently dropping them.
        if idx < len(flat_accounts) and slots:
            for remaining in flat_accounts[idx:]:
                assignment[-1].append(remaining)

        return [
            (slot["slide_idx"], slot["slot_name"], self._merge_contd_pairs(assignment[s_i]))
            for s_i, slot in enumerate(slots)
            if assignment[s_i]
        ]

    def apply_structured_data_to_slides(self, structured_data: List[Dict], start_slide: int,
                                       project_name: str, statement_type: str, is_chinese_databook: bool = False):
        """Apply structured data directly to slides (slides 1-4 for BS, 5-8 for IS)"""
        if not self.presentation:
            self.load_template()

        stage_started_at = time.perf_counter()
        logger.info("Applying %s accounts to slides starting at %s", len(structured_data), start_slide)

        # Normalize commentary and store originals for fill optimization
        structured_data = self._prepare_structured_data_for_slides(structured_data)

        # Distribute content across textbox slots based on capacity
        max_slides = int(self.pptx_settings.get("max_commentary_slides_per_statement", 4) or 4)
        slot_distribution = self._distribute_content_across_slots(
            structured_data,
            max_slides=max_slides,
            start_slide=start_slide,
            statement_type=statement_type,
        )
        
        # Group slot distribution by slide for easier processing
        slides_content = {}  # {slide_idx: {'single': [...], 'L': [...], 'R': [...]}}
        for slot_slide_idx, slot_name, account_list in slot_distribution:
            if slot_slide_idx not in slides_content:
                slides_content[slot_slide_idx] = {}
            slides_content[slot_slide_idx][slot_name] = account_list
        
        # Ensure we have enough slides
        if slides_content:
            max_slide_idx = max(slides_content.keys())
            needed_slides = start_slide + max_slide_idx
            current_slide_count = len(self.presentation.slides)
            
            if needed_slides > current_slide_count:
                # Add slides if needed
                if current_slide_count > 0:
                    slide_layout = self.presentation.slides[0].slide_layout
                    for _ in range(needed_slides - current_slide_count):
                        self.presentation.slides.add_slide(slide_layout)
        
        # Track which slides are used
        used_slide_indices = set()
        summary_jobs: List[Dict[str, Any]] = []
        
        # Apply content to slides
        for slide_idx in sorted(slides_content.keys()):
            actual_slide_idx = start_slide - 1 + slide_idx  # Convert to 0-based
            if actual_slide_idx >= len(self.presentation.slides):
                logger.warning("Slide index %s exceeds available slides", actual_slide_idx + 1)
                continue
            
            used_slide_indices.add(actual_slide_idx)
            slide = self.presentation.slides[actual_slide_idx]
            slot_contents = slides_content[slide_idx]  # {'single': [...], 'L': [...], 'R': [...]}
            
            # Note: Financial tables are filled by embed_financial_tables()
            
            # Collect all accounts on this slide for summary generation
            all_slide_accounts = []
            for slot_name, accounts in slot_contents.items():
                all_slide_accounts.extend(accounts)

            used_slot_shape_ids = set()
            
            # Fill each slot (single, L, or R) on this slide
            for slot_name, account_data_list in slot_contents.items():
                if not account_data_list:
                    continue
                
                # Find the appropriate shape based on slot_name
                bullets_shape = self._resolve_commentary_slot_shape(
                    slide,
                    slot_name,
                    used_shape_ids=used_slot_shape_ids,
                )
                if not bullets_shape and slot_name == "single":
                    bullets_shape = self._resolve_commentary_slot_shape(
                        slide,
                        "L",
                        used_shape_ids=used_slot_shape_ids,
                    ) or self._resolve_commentary_slot_shape(
                        slide,
                        "R",
                        used_shape_ids=used_slot_shape_ids,
                    )
                if not bullets_shape:
                    bullets_shape = self._add_commentary_slot_shape(slide, slot_name)
                
                if not bullets_shape.has_text_frame:
                    logger.warning("Slide %s: Shape for slot '%s' has no text frame", actual_slide_idx + 1, slot_name)
                    continue
                used_slot_shape_ids.add(id(bullets_shape))
                
                # Fill this slot
                tf = bullets_shape.text_frame
                tf.clear()
                tf.word_wrap = True
                self._force_no_autofit(tf)  # keep text at 9pt/10pt, never shrink
                from pptx.enum.text import MSO_VERTICAL_ANCHOR
                tf.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
                
                # Determine optimal font size to prevent overflow
                slot_font_size = self._determine_slot_font_size(
                    account_data_list, bullets_shape, slot_name, statement_type=statement_type,
                )
                if slot_font_size < 9:
                    logger.info("Slide %s, slot '%s': Reducing font to %spt to fit %s accounts",
                                actual_slide_idx + 1, slot_name, slot_font_size, len(account_data_list))
                else:
                    logger.info("Slide %s, slot '%s': Filling with %s accounts at %spt",
                                actual_slide_idx + 1, slot_name, len(account_data_list), slot_font_size)

                # Fill with accounts, grouped by category
                # Show category header only once per category group
                current_category = None
                for account_idx, account_data in enumerate(account_data_list):
                    category = account_data.get('category', '')
                    mapping_key = account_data.get('mapping_key', account_data.get('account_name', ''))
                    display_name = account_data.get('display_name', mapping_key)
                    commentary = account_data.get('commentary', '')
                    is_chinese = account_data.get('is_chinese', False)
                    is_continuation = account_data.get('is_continuation', False)
                    
                    # Skip category header if this is a continuation of a split account
                    # Show category header only when category changes
                    if category and category != current_category and not is_continuation:
                        # Add category header - use Chinese if databook is Chinese
                        p_category = tf.add_paragraph()
                        p_category.level = 0
                        try:
                            p_category.left_indent = Inches(0.21)
                            p_category.first_line_indent = Inches(-0.19)
                            p_category.space_before = Pt(3) if current_category else Pt(0)
                            p_category.space_after = Pt(0)
                            p_category.line_spacing = 1.0
                        except:
                            pass
                        
                        run_category = p_category.add_run()
                        # Use Chinese category name if databook is Chinese
                        category_text = translate_category_to_chinese(category) if is_chinese_databook else category
                        
                        run_category.text = category_text
                        run_category.font.size = Pt(slot_font_size)
                        run_category.font.name = 'Arial'
                        run_category.font.bold = False
                        try:
                            from pptx.dml.color import RGBColor
                            run_category.font.color.rgb = RGBColor(0, 51, 102)  # Dark blue
                        except:
                            pass
                        
                        current_category = category
                    
                    # Fill commentary with key formatting
                    # For continuation accounts, show "(cont'd)" or "(续)" after key name
                    if is_continuation:
                        # More prominent continuation marker
                        if is_chinese_databook:
                            display_name_with_cont = f"{display_name} (续)"
                        else:
                            display_name_with_cont = f"{display_name} (cont'd)"
                        
                        # Log continuation for debugging
                        logger.info("Displaying continuation: %s", display_name_with_cont)
                    else:
                        display_name_with_cont = display_name
                    
                    self._fill_text_main_bullets_with_category_and_key(
                        tf, None, display_name_with_cont, commentary, is_chinese,
                        is_chinese_databook=is_chinese_databook, needs_continuation=False,
                        font_size_pt=slot_font_size,
                    )
            
            page_commentary, page_summary_source = self._build_page_summary_source(all_slide_accounts)

            # Collect coSummaryShape jobs and fill after summaries are generated.
            summary_shape = self.find_shape_by_name(slide.shapes, "coSummaryShape")
            if summary_shape and summary_shape.has_text_frame:
                summary_shape.text_frame.clear()
                self._force_no_autofit(summary_shape.text_frame)
                if page_summary_source:
                    summary_jobs.append({
                        "slide_idx": actual_slide_idx,
                        "summary_shape": summary_shape,
                        "page_commentary": page_commentary,
                        "page_summary_source": page_summary_source,
                        "is_chinese": is_chinese_databook,
                        "font_is_chinese": all_slide_accounts[0].get('is_chinese', False) if all_slide_accounts else False,
                    })
            else:
                logger.info("Slide %s: coSummaryShape not present; skipping page summary", actual_slide_idx + 1)
            
            logger.info("Filled slide %s with %s accounts across %s slots", actual_slide_idx + 1, len(all_slide_accounts), len(slot_contents))

        summary_results = self._generate_slide_summaries(summary_jobs)
        for job in summary_jobs:
            final_summary = str(summary_results.get(job["slide_idx"]) or "").strip()
            if not final_summary:
                continue
            summary_shape = job["summary_shape"]
            p = summary_shape.text_frame.paragraphs[0] if summary_shape.text_frame.paragraphs else summary_shape.text_frame.add_paragraph()
            p.text = final_summary
            for run in p.runs:
                run.font.size = get_font_size_for_text(final_summary, force_chinese_mode=job["font_is_chinese"])
                run.font.name = get_font_name_for_text(final_summary)
        
        # Note: Unused slides will be removed at the end, after all content and tables are embedded
        # Store unused slides for later removal
        statement_slide_range = list(range(start_slide - 1, min(start_slide + 3, len(self.presentation.slides))))
        unused_slides = [idx for idx in statement_slide_range if idx not in used_slide_indices]
        if unused_slides:
            # Store for later removal - don't remove now
            if not hasattr(self, '_unused_slides_to_remove'):
                self._unused_slides_to_remove = []
            self._unused_slides_to_remove.extend(unused_slides)
            logger.info("Marked %s unused slides for %s for later removal: %s", len(unused_slides), statement_type, [idx + 1 for idx in unused_slides])
        logger.info(
            "PPTX stage apply_structured_data_to_slides[%s] took %.2fs across %s populated slides",
            statement_type,
            time.perf_counter() - stage_started_at,
            len(slides_content),
        )
    
    def _remove_slides(self, slide_indices):
        """Remove slides by indices (from backup method)"""
        # Sort in reverse order to maintain indices while removing
        for slide_idx in sorted(slide_indices, reverse=True):
            if slide_idx < len(self.presentation.slides):
                try:
                    # Use XML-based removal (from backup method)
                    xml_slides = self.presentation.slides._sldIdLst
                    slides = list(xml_slides)
                    
                    if slide_idx < len(slides):
                        # Get the slide element to remove
                        slide_element = slides[slide_idx]
                        # Remove relationship
                        rId = slide_element.rId
                        self.presentation.part.drop_rel(rId)
                        # Remove from XML
                        xml_slides.remove(slide_element)
                        logger.info("Removed slide %s", slide_idx + 1)
                    else:
                        logger.warning("Slide index %s out of range (only %s slides)", slide_idx, len(slides))
                except Exception as e:
                    logger.warning("Could not remove slide %s: %s", slide_idx + 1, e)
                    logger.debug(traceback.format_exc())
    
    def _set_cell_border(self, cell, border_position='top', color_rgb=None, width=Pt(1)):
        """Set cell border"""
        from pptx.oxml.xmlchemy import OxmlElement
        
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        
        # Map position to tag name
        tag_map = {'top': 'lnT', 'bottom': 'lnB', 'left': 'lnL', 'right': 'lnR'}
        tag_name = tag_map.get(border_position)
        if not tag_name:
            return
            
        # Check if line element exists
        ln = tcPr.find(f"{{http://schemas.openxmlformats.org/drawingml/2006/main}}{tag_name}")
        if ln is None:
            ln = OxmlElement(f"a:{tag_name}")
            tcPr.append(ln)
            
        # Set properties
        ln.set('w', str(int(width)))
        ln.set('cap', 'flat')
        ln.set('cmpd', 'sng')
        ln.set('algn', 'ctr')
        
        # Set color
        if color_rgb:
            solidFill = OxmlElement('a:solidFill')
            srgbClr = OxmlElement('a:srgbClr')
            # Convert RGBColor or tuple to hex string
            hex_color = "000000"
            if isinstance(color_rgb, str):
                hex_color = color_rgb.replace('#', '')
            elif isinstance(color_rgb, tuple) and len(color_rgb) == 3:
                hex_color = f"{color_rgb[0]:02x}{color_rgb[1]:02x}{color_rgb[2]:02x}"
            # If it's an RGBColor object, user should pass str or tuple for this low-level func
                
            srgbClr.set('val', hex_color)
            solidFill.append(srgbClr)
            ln.append(solidFill)
            
            prstDash = OxmlElement('a:prstDash')
            prstDash.set('val', 'solid')
            ln.append(prstDash)
            
            round_ = OxmlElement('a:round')
            ln.append(round_)
            
            headEnd = OxmlElement('a:headEnd')
            headEnd.set('type', 'none')
            headEnd.set('w', 'med')
            headEnd.set('len', 'med')
            ln.append(headEnd)
            
            tailEnd = OxmlElement('a:tailEnd')
            tailEnd.set('type', 'none')
            tailEnd.set('w', 'med')
            tailEnd.set('len', 'med')
            ln.append(tailEnd)

    def _fill_table_placeholder(self, shape, df, table_name: str = None, currency_unit: str = None, bounds: Dict[str, int] = None):
        """Fill table placeholder with DataFrame data, preserving original formatting
        Args:
            shape: Table shape or placeholder
            df: DataFrame with data
            table_name: Name of the table (e.g., "示意性调整后资产负债表 - xxxx")
            currency_unit: Currency unit (e.g., "人民币千元" or "CNY'000") to replace "Description"
        """
        try:
            # Debug: Log DataFrame content
            logger.info("Filling table with DF shape: %s", df.shape)
            if not df.empty:
                logger.info("First row data: %s", df.iloc[0].to_dict())
                # Check if any data is non-zero
                numeric_cols = df.select_dtypes(include=['number']).columns
                if len(numeric_cols) > 0:
                    non_zero_count = (df[numeric_cols] != 0).sum().sum()
                    logger.info("Non-zero values in DF: %s", non_zero_count)
            
            # Find parent slide
            slide = None
            for s in self.presentation.slides:
                for shp in s.shapes:
                    if shp == shape:
                        slide = s
                        break
                if slide:
                    break
            
            if bounds is None:
                bounds = {
                    "left": shape.left,
                    "top": shape.top,
                    "width": shape.width,
                    "height": shape.height,
                }

            # Adjust position and size from resolved layout bounds
            try:
                shape.left = bounds["left"]
                shape.top = bounds["top"]
                shape.width = bounds["width"]
                shape.height = bounds["height"]
            except Exception as e:
                logger.warning("Could not adjust table position/width: %s", e)

            # Check if shape is a TablePlaceholder (textbox placeholder)
            from pptx.shapes.placeholder import TablePlaceholder
            
            table = None
            # Check if it's a TablePlaceholder (textbox placeholder named "Table Placeholder")
            is_table_placeholder = False
            try:
                is_table_placeholder = isinstance(shape, TablePlaceholder)
            except:
                # Check by name
                if hasattr(shape, 'name') and 'Table' in shape.name and 'Placeholder' in shape.name:
                    is_table_placeholder = True
            
            if is_table_placeholder:
                # It's a table placeholder - insert a table into it
                logger.info("Found TablePlaceholder (%s), inserting table with %s rows and %s columns", shape.name if hasattr(shape, 'name') else 'unnamed', len(df), len(df.columns))
                try:
                    left = bounds["left"]
                    top = bounds["top"]
                    width = bounds["width"]
                    height = bounds["height"]
                    
                    # Find the slide containing this shape (already found above)
                    if slide:
                        # Remove the placeholder shape
                        sp = shape._element
                        slide.shapes._spTree.remove(sp)
                        
                        # Add new table at the same position
                        # Need: 1 row for title (if table_name), 1 for header, N for data
                        total_rows = len(df) + 2 if table_name else len(df) + 1
                        table_shape = slide.shapes.add_table(
                            rows=total_rows,
                            cols=len(df.columns),
                            left=left,
                            top=top,
                            width=width,
                            height=height
                        )
                        table = table_shape.table
                        logger.info("Inserted new table: %s rows, %s columns", len(table.rows), len(table.columns))
                except Exception as e:
                    logger.error("Could not insert table into placeholder: %s", e)
                    logger.debug(traceback.format_exc())
            elif hasattr(shape, 'table'):
                # Try to access existing table
                try:
                    table = shape.table
                    logger.info("Found existing table with %s rows and %s columns", len(table.rows), len(table.columns))
                except ValueError:
                    # Shape doesn't contain a table
                    logger.warning("Shape has table attribute but doesn't contain a table")
                    table = None
            
            if table:
                # Colors
                DARK_BLUE = RGBColor(0, 51, 102)
                TIFFANY_BLUE = RGBColor(10, 186, 181)
                GREY = RGBColor(217, 217, 217)
                WHITE = RGBColor(255, 255, 255)
                BLACK = RGBColor(0, 0, 0)
                
                self._fit_table_columns(table, df)

                total_visible_rows = len(df) + 1 + (1 if table_name else 0)
                if total_visible_rows >= 26:
                    data_font_size = Pt(7)
                    data_row_height = Inches(0.16)
                elif total_visible_rows >= 20:
                    data_font_size = Pt(7.5)
                    data_row_height = Inches(0.18)
                else:
                    data_font_size = Pt(8)
                    data_row_height = Inches(0.20)
                
                # Add table name as first row if provided
                if table_name:
                    # Insert a new row at the top for table name
                    try:
                        # Ensure table has at least one row
                        if len(table.rows) == 0:
                            table.rows.add_row()
                            
                        name_row = table.rows[0]  # Use first row for name
                        name_row.height = Inches(0.25)
                        
                        # Merge all cells in first row for table name
                        if len(table.columns) > 1:
                            name_row.cells[0].merge(name_row.cells[len(table.columns) - 1])
                        name_cell = name_row.cells[0]
                        name_cell.text = table_name
                        # Format table name: Arial 9, bold, centered, Dark Blue bg, White font
                        if name_cell.text_frame.paragraphs:
                            p = name_cell.text_frame.paragraphs[0]
                            p.alignment = PP_ALIGN.CENTER  # Center alignment
                            if p.runs:
                                run = p.runs[0]
                            else:
                                run = p.add_run()
                            run.font.name = 'Arial'
                            run.font.size = Pt(9)
                            run.font.bold = True
                            run.font.color.rgb = WHITE
                            
                            name_cell.fill.solid()
                            name_cell.fill.fore_color.rgb = DARK_BLUE
                            
                        # Shift data down - we'll use rows starting from index 1
                        data_start_row = 1
                    except:
                        data_start_row = 0
                else:
                    data_start_row = 0
                
                # Fill header row with formatting
                max_cols = min(len(df.columns), len(table.columns))
                header_row_idx = data_start_row
                
                # Ensure header row exists
                if len(table.rows) <= header_row_idx:
                    table.rows.add_row()
                
                # Set header row height slightly taller for readability
                try:
                    table.rows[header_row_idx].height = Inches(0.25)
                except:
                    pass
                    
                for col_idx, col_name in enumerate(df.columns[:max_cols]):
                    if col_idx < len(table.columns):
                        cell = table.cell(header_row_idx, col_idx)
                        # Replace "Description" with currency unit if found
                        if currency_unit and (col_name.lower() == 'description' or '描述' in str(col_name) or '项目' in str(col_name)):
                            cell.text = currency_unit
                        else:
                            cell.text = str(col_name)
                        # Apply header formatting: Arial 9, bold, Tiffany Blue bg, White font
                        if cell.text_frame.paragraphs:
                            p = cell.text_frame.paragraphs[0]
                            p.alignment = PP_ALIGN.CENTER
                            
                            if p.runs:
                                run = p.runs[0]
                            else:
                                run = p.add_run()
                            
                            run.font.name = 'Arial'
                            run.font.size = Pt(9)
                            run.font.bold = True
                            run.font.color.rgb = WHITE # White font for header

                            cell.fill.solid()
                            cell.fill.fore_color.rgb = TIFFANY_BLUE

                        try:
                            cell.margin_left = Inches(0.04)
                            cell.margin_right = Inches(0.04)
                            cell.margin_top = Inches(0.02)
                            cell.margin_bottom = Inches(0.02)
                        except Exception:
                            pass

                        logger.debug("Filled header cell %s: %s", col_idx, cell.text)
                
                # Fill data rows with formatting - show ALL rows (no limit)
                # Check if table has enough rows, if not, limit to available rows
                max_rows = len(df)  # Show all rows
                rows_needed = max_rows + data_start_row + 1  # +1 for header row
                available_rows = len(table.rows)
                
                if available_rows < rows_needed:
                    logger.warning("Table has %s rows but needs %s. Will only fill %s data rows.", available_rows, rows_needed, available_rows - data_start_row - 1)
                    max_rows = min(max_rows, available_rows - data_start_row - 1)
                    if max_rows < 0:
                        max_rows = 0
                
                logger.info("Table has %s rows available, will fill %s data rows", available_rows, max_rows)
                
                # Now fill all rows with Arial 9 font
                # Check for title, date, total, and subtotal rows to highlight
                logger.info("Filling %s data rows, starting at row index %s, table has %s rows", max_rows, header_row_idx + 1, len(table.rows))
                for row_idx in range(max_rows):
                    if row_idx >= len(df):
                        break
                    df_row = df.iloc[row_idx]
                    first_col_value = str(df_row.iloc[0]) if len(df_row) > 0 else ""
                    
                    # Check if this is a title, date, total, or subtotal row
                    is_special_row = False
                    is_total_row = False
                    is_date_row = False
                    first_col_lower = first_col_value.lower()
                    total_keywords = ['total', '合计', '总计', '小计', 'subtotal', 'sub-total', 'sub total']
                    date_keywords = ['date', '日期', '年', '月']
                    special_keywords = total_keywords + ['title', '标题'] + date_keywords
                    
                    if any(keyword in first_col_lower for keyword in special_keywords):
                        is_special_row = True
                    
                    if any(keyword in first_col_lower for keyword in total_keywords):
                        is_total_row = True
                    
                    if any(keyword in first_col_lower for keyword in date_keywords):
                        is_date_row = True
                    
                    # Data row index = header_row_idx + 1 + row_idx
                    data_row_idx = header_row_idx + 1 + row_idx
                    if data_row_idx >= len(table.rows):
                        logger.warning("Data row index %s exceeds table rows %s, skipping", data_row_idx, len(table.rows))
                        break
                    
                    # Set data row height based on table density
                    try:
                        table.rows[data_row_idx].height = data_row_height
                    except:
                        pass
                    
                    # Log first row processing
                    if row_idx == 0:
                        logger.info("Processing first data row: %s", df_row.values[:3])

                    # Unit scaling policy for the financial table:
                    #   The extractor is called with multiply_values=False
                    #   (embed_financial_tables), so numeric values flow through
                    #   in the ORIGINAL source units declared by the workbook
                    #   header. If the header says CNY'000 / 人民币千元, values
                    #   already represent thousands and must NOT be multiplied.
                    #   Same for CNY'M / 人民币百万 (millions). The column
                    #   header shows the unit so the reader interprets them
                    #   correctly. Any accidental scaling here would double-count.
                    for col_idx, col_name in enumerate(df.columns[:max_cols]):
                        if col_idx >= len(table.columns):
                            break
                        cell = table.cell(data_row_idx, col_idx)

                        # Get value from DataFrame safely
                        value = df_row[col_name] if col_name in df_row.index else ""
                        text_val = self._format_table_value(value, is_numeric_column=col_idx > 0)
                        
                        # Set text
                        cell.text = text_val
                        
                        # Log first cell value of first row
                        if row_idx == 0 and col_idx < 2:
                            logger.info("Setting cell (%s, %s) to: '%s'", data_row_idx, col_idx, text_val)
                        
                        # Apply formatting: Arial 7pt (reduced from 9pt) for all cells
                        # Note: Always access paragraphs[0] AFTER setting text
                        if not cell.text_frame.paragraphs:
                            cell.text_frame.add_paragraph()
                            
                        p = cell.text_frame.paragraphs[0]
                        if not p.runs:
                            p.add_run()
                            
                        # Apply formatting to ALL runs (setting cell.text might create one run, but best to be safe)
                        for run in p.runs:
                            run.text = text_val # Ensure text is set in the run
                            run.font.name = 'Arial'
                            run.font.size = Pt(7) if is_date_row else data_font_size

                            # Force Black color for data rows
                            try:
                                run.font.color.rgb = BLACK
                            except:
                                pass

                            # Bold for special rows
                            run.font.bold = is_special_row

                        # Give cells a small internal margin so text doesn't hug the border
                        try:
                            cell.margin_left = Inches(0.04)
                            cell.margin_right = Inches(0.04)
                            cell.margin_top = Inches(0.01)
                            cell.margin_bottom = Inches(0.01)
                        except Exception:
                            pass

                        # First column left-aligned, numeric columns right-aligned
                        try:
                            p.alignment = PP_ALIGN.LEFT if col_idx == 0 else PP_ALIGN.RIGHT
                        except Exception:
                            pass
                        
                        # Highlight special rows
                        if is_special_row:
                            try:
                                cell.fill.solid()
                                cell.fill.fore_color.rgb = GREY
                            except:
                                pass
                        else:
                            # White background for normal rows
                            try:
                                cell.fill.solid()
                                cell.fill.fore_color.rgb = WHITE
                            except:
                                pass
                                
                        # Add bold top horizontal line for total/subtotal rows
                        if is_total_row:
                            try:
                                # Top border in Dark Blue, 2pt width. Pass hex string "003366"
                                self._set_cell_border(cell, 'top', color_rgb="003366", width=Pt(2))
                            except:
                                pass
                    
                    logger.debug("Filled table row %s (data_row_idx: %s, special: %s)", row_idx + 1, data_row_idx, is_special_row)
                
                logger.info("Updated table with Excel data (formatting preserved)")
            else:
                # If no table, this is an error - table placeholder should be a table shape
                logger.error("Table Placeholder is not a table shape! Cannot embed financial table.")
                logger.error("Shape type: %s, has_table: %s", type(shape), hasattr(shape, 'table'))
                logger.error("Shape name: %s", shape.name if hasattr(shape, 'name') else 'unnamed')
                # Check if shape has table attribute but it's None
                if hasattr(shape, 'table'):
                    logger.error("shape.table is: %s", shape.table)
                # Try to create a table representation in text frame as last resort
                if shape.has_text_frame:
                    shape.text_frame.clear()
                    # Convert DataFrame to formatted text table - show ALL rows
                    try:
                        # Show all rows, no limit
                        text_table = df.to_string(index=False)
                    except:
                        text_table = str(df)
                    
                    p = shape.text_frame.paragraphs[0] if shape.text_frame.paragraphs else shape.text_frame.add_paragraph()
                    p.text = text_table
                    logger.warning("Added text table representation with all %s rows (%s chars) - NOT IDEAL, should be table format", len(df), len(text_table))
        except Exception as e:
            logger.error("Could not fill table placeholder: %s", e)
            logger.error(traceback.format_exc())
            # Fallback: add text representation - show ALL rows
            if shape.has_text_frame:
                shape.text_frame.clear()
                # Show all rows, not just first 10
                text_repr = df.to_string(index=False)
                p = shape.text_frame.paragraphs[0] if shape.text_frame.paragraphs else shape.text_frame.add_paragraph()
                p.text = text_repr
    
    def _detect_bullet_levels(self, text: str) -> List[Tuple[int, str]]:
        """
        Detect bullet levels (1-3) from commentary text
        Returns list of (level, text) tuples where level 0 = no bullet, 1-3 = bullet levels
        """
        lines = text.split('\n')
        bullet_lines = []
        
        for line in lines:
            stripped = line.strip()
            original_line = line
            
            # Detect bullet lines with '- ' prefix
            if original_line.lstrip().startswith('- '):
                # Calculate indentation level (based on spaces/tabs before the bullet)
                indent_spaces = len(original_line) - len(original_line.lstrip())
                
                # Determine bullet level based on indentation (2 spaces per level)
                level = min(3, (indent_spaces // 2) + 1)  # Cap at level 3
                
                # Clean and store bullet line
                clean_line = stripped[2:]  # Remove '- '
                
                # Special handling for level 3 bullets that contain a dash indicating sub-level
                if level == 3 and " - " in clean_line:
                    # Split at the first occurrence of " - "
                    parts = clean_line.split(" - ", 1)
                    if len(parts) > 1:
                        # Add level 3 content
                        bullet_lines.append((level, parts[0].strip()))
                        # Add continuation as level 3 (indented)
                        bullet_lines.append((level, parts[1].strip()))
                    else:
                        bullet_lines.append((level, clean_line))
                else:
                    bullet_lines.append((level, clean_line))
            elif stripped:
                # Regular content (no bullet) - level 0
                bullet_lines.append((0, stripped))
        
        return bullet_lines
    
    def _determine_slot_font_size(
        self,
        slot_accounts: List[Dict],
        shape,
        slot_name: str,
        statement_type: Optional[str] = None,
    ) -> int:
        """Return the fixed base font size for commentary slots.

        Font size is intentionally kept constant across all slides — 10pt for
        Chinese content, 9pt for English.  The distribution algorithm is
        responsible for fitting content within capacity; shrinking the font to
        compensate produces inconsistent slide-to-slide typography."""
        is_chinese_slot = any(
            any("\u4e00" <= c <= "\u9fff" for c in str(a.get("commentary", "")))
            for a in slot_accounts
        )
        return 10 if is_chinese_slot else 9

    @staticmethod
    def _force_no_autofit(text_frame) -> None:
        """Set the text frame's bodyPr autofit to ``<a:noAutofit/>`` so
        PowerPoint never shrinks the text to fit the shape. The template
        ships with ``<a:spAutoFit/>`` (resize shape to fit text), which in
        some viewers falls back to shrinking the text when the shape can't
        grow. Forcing ``noAutofit`` keeps the text at the exact point size
        we set (9pt / 10pt); overflow is simply clipped at the shape edge."""
        try:
            from lxml import etree  # noqa: F401
            from pptx.oxml.ns import qn
            bodyPr = text_frame._txBody.bodyPr
            # Remove any existing autofit child (spAutoFit / normAutofit / noAutofit).
            for tag in ("a:spAutoFit", "a:normAutofit", "a:noAutofit"):
                for child in bodyPr.findall(qn(tag)):
                    bodyPr.remove(child)
            from pptx.oxml import parse_xml
            bodyPr.append(parse_xml(
                '<a:noAutofit xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'
            ))
        except Exception as exc:
            logger.debug("Could not force noAutofit on text frame: %s", exc)

    def _determine_slot_font_size_UNUSED(
        self,
        slot_accounts: List[Dict],
        shape,
        slot_name: str,
        statement_type: Optional[str] = None,
    ) -> int:
        """KEPT FOR REFERENCE — old shrink-to-fit logic (9→8→7pt)."""
        packing = self._packing_settings(statement_type)
        if not shape or not hasattr(shape, "height"):
            return 9

        is_chinese_slot = any(
            any("\u4e00" <= c <= "\u9fff" for c in str(a.get("commentary", "")))
            for a in slot_accounts
        )

        pillow_ok = self._pillow_fitting_enabled(packing)
        if pillow_ok:
            try:
                from fdd_utils.text_metrics import (
                    get_font,
                    line_height_pt as _line_h,
                    lines_that_fit,
                    text_box_from_shape,
                    wrap_text,
                )
                box = text_box_from_shape(shape)
                family = "Microsoft YaHei" if is_chinese_slot else "Arial"
                line_spacing = 0.95 if is_chinese_slot else 1.0
                for candidate_pt in (9, 8, 7):
                    font = get_font(family, candidate_pt, is_cjk=is_chinese_slot)
                    line_h = _line_h(font, line_spacing=line_spacing)
                    capacity = lines_that_fit(box.height_pt, line_h)
                    total_lines = 0
                    prev_cat = None
                    for acct in slot_accounts:
                        cat = acct.get("category", "")
                        if cat and cat != prev_cat:
                            total_lines += 1
                            prev_cat = cat
                        parts: List[str] = []
                        mapping_key = acct.get("mapping_key", acct.get("account_name", ""))
                        if mapping_key:
                            parts.append(str(mapping_key))
                        commentary = str(acct.get("commentary", ""))
                        if commentary:
                            parts.append(commentary)
                        joined = "\n".join(parts)
                        if joined:
                            total_lines += len(wrap_text(joined, font, box.width_pt))
                    if total_lines <= capacity:
                        return candidate_pt
                return 7
            except Exception:
                pass  # fall through to legacy

        for candidate_pt in (9, 8, 7):
            shape_height_pt = shape.height * 72 / 914400
            effective_height = shape_height_pt * float(packing.get("shape_height_utilization", 1.02))
            line_spacing = 0.95 if is_chinese_slot else 1.0
            line_height = (candidate_pt * line_spacing) + float(packing.get("line_height_padding_pt", 1.6))
            max_lines = int(effective_height / line_height)

            total_lines = 0
            prev_cat = None
            for acct in slot_accounts:
                cat = acct.get("category", "")
                if cat and cat != prev_cat:
                    total_lines += 1
                    prev_cat = cat
                commentary = str(acct.get("commentary", ""))
                is_chi = any("\u4e00" <= c <= "\u9fff" for c in commentary)
                base_cpl = self._estimate_chars_per_line(slot_name, is_chi, shape=shape, statement_type=statement_type)
                scale = 9.0 / candidate_pt
                cpl = max(16, int(base_cpl * scale))
                total_lines += 1  # key line
                for line in commentary.split("\n"):
                    if line.strip():
                        total_lines += max(1, (len(line) + cpl - 1) // cpl)

            if total_lines <= max_lines:
                return candidate_pt
        return 7

    @staticmethod
    def _truncate_commentary_to_fit(commentary: str, max_chars: int) -> str:
        """Hard truncation at sentence boundary, with ellipsis."""
        if len(commentary) <= max_chars:
            return commentary
        truncated = commentary[:max_chars]
        # Try to cut at sentence boundary
        for end_char in (". ", "。", "! ", "？"):
            pos = truncated.rfind(end_char)
            if pos > max_chars * 0.5:
                return truncated[: pos + len(end_char)].rstrip()
        # Fall back to word boundary
        pos = truncated.rfind(" ", int(max_chars * 0.7))
        if pos > 0:
            return truncated[:pos].rstrip() + "..."
        return truncated.rstrip() + "..."

    def _fill_text_main_bullets_with_category_and_key(self, text_frame, category: str, display_name: str,
                                                      commentary: str, is_chinese: bool, is_chinese_databook: bool = False,
                                                      needs_continuation: bool = False, font_size_pt: int = 9):
        """
        Fill textMainBullets shape with commentary formatted as:
        - Category as first level (dark blue Arial 9) - only if category is provided
        - Key name with filled round bullet + space + key name (black bold Arial 9) + "-" (not bold) + plain text
        - Indentation 0.15" with special hanging 0.15", spacing after 6pt
        """
        from pptx.util import Inches
        from pptx.dml.color import RGBColor
        from pptx.enum.text import PP_ALIGN
        
        # Add category as first level (if category exists and is not None)
        # Note: category is now handled at slide level, so this is only for individual calls
        if category:
            p_category = text_frame.add_paragraph()
            p_category.level = 0
            try:
                p_category.left_indent = Inches(0.21)
                p_category.first_line_indent = Inches(-0.19)
                p_category.space_before = Pt(0)
                p_category.space_after = Pt(0)
                p_category.line_spacing = 1.0
            except:
                pass
            
            run_category = p_category.add_run()
            run_category.text = category
            run_category.font.size = Pt(font_size_pt)
            run_category.font.name = 'Arial'
            run_category.font.bold = False
            try:
                run_category.font.color.rgb = RGBColor(0, 51, 102)  # Dark blue
            except:
                pass
        
        # Add key name with grey char + space + key name (black bold) + "-" (not bold) + plain text
        p_key = text_frame.add_paragraph()
        p_key.level = 0  # No bullet level, we'll use grey character
        try:
            # Set formatting
            p_key.left_indent = Inches(0.15)  # 0.15" indent
            p_key.first_line_indent = Inches(-0.15)  # 0.15" special hanging
            p_key.space_before = Pt(0)
            p_key.space_after = Pt(3)  # Matches _PARA_SPACE_AFTER (cost estimator)
            p_key.line_spacing = 1.0
        except Exception as e:
            logger.warning("Could not set paragraph formatting: %s", e)
            pass
        
        # Grey char (U+25A0) + space
        run_bullet = p_key.add_run()
        run_bullet.text = '\u25A0 '  # U+25A0 (black square) + space
        run_bullet.font.size = Pt(font_size_pt)
        run_bullet.font.name = 'Arial'
        run_bullet.font.bold = False
        try:
            run_bullet.font.color.rgb = RGBColor(128, 128, 128)  # Grey
        except:
            pass

        # Key name (black bold)
        run_key = p_key.add_run()
        run_key.text = display_name
        run_key.font.size = Pt(font_size_pt)
        run_key.font.name = 'Arial'
        run_key.font.bold = True
        try:
            run_key.font.color.rgb = RGBColor(0, 0, 0)  # Black
        except:
            pass

        # "-" (not bold)
        run_dash = p_key.add_run()
        run_dash.text = " - "
        run_dash.font.size = Pt(font_size_pt)
        run_dash.font.name = 'Arial'
        run_dash.font.bold = False
        try:
            run_dash.font.color.rgb = RGBColor(0, 0, 0)  # Black
        except:
            pass
        
        # Plain text (commentary content)
        commentary_lines = commentary.split('\n')
        first_line_added = False
        for line_idx, line in enumerate(commentary_lines):
            line = line.strip()
            if not line:
                continue
            
            if not first_line_added:
                # First line continues on same paragraph after the dash
                run_text = p_key.add_run()
                run_text.text = line
                first_line_added = True
            else:
                # Subsequent lines as new paragraphs (indented continuation)
                p_text = text_frame.add_paragraph()
                p_text.level = 0  # No bullet for continuation
                try:
                    p_text.left_indent = Inches(0.15)  # 0.15" indent (same as key text)
                    p_text.first_line_indent = Inches(0)  # No hanging for continuation lines
                    p_text.space_before = Pt(0)
                    p_text.space_after = Pt(3)
                    p_text.line_spacing = 1.0
                except:
                    pass
                run_text = p_text.add_run()
                run_text.text = line
            
            # Apply formatting to the run
            run_text.font.size = Pt(font_size_pt)
            run_text.font.name = 'Arial'
            run_text.font.bold = False
            try:
                run_text.font.color.rgb = RGBColor(0, 0, 0)  # Black
            except:
                pass

        # Note: "(continued)" is now added to category header, not here
    
    def _fill_text_main_bullets_with_levels(self, text_frame, commentary: str, is_chinese: bool):
        """
        Fill textMainBullets shape with commentary using detailed line break logic
        and level 1-3 text handling with page breaks (legacy method, kept for compatibility)
        """
        from pptx.util import Inches
        from pptx.dml.color import RGBColor
        from pptx.enum.text import PP_ALIGN
        
        # Detect bullet levels
        bullet_lines = self._detect_bullet_levels(commentary)
        
        # Calculate max lines that can fit in the shape
        # Estimate based on shape height (conservative estimate)
        max_lines = 35  # Default conservative estimate
        
        lines_added = 0
        
        for level, text in bullet_lines:
            if not text.strip():
                continue
            
            # Check if we need a page break (if shape is getting full)
            # Note: Actual page breaks would require creating new slides, which is handled
            # at a higher level. Here we just ensure content fits.
            if lines_added >= max_lines:
                # Add continuation indicator
                p = text_frame.add_paragraph()
                p.level = 0
                run = p.add_run()
                run.text = "... (continued on next page)" if not is_chinese else "... (续下页)"
                run.font.size = get_font_size_for_text(run.text, force_chinese_mode=is_chinese)
                run.font.name = get_font_name_for_text(run.text)
                run.font.italic = True
                break
            
            # Create paragraph with appropriate level
            p = text_frame.add_paragraph()
            p.level = level  # Set bullet level (0-3)
            
            # Apply paragraph formatting based on level
            try:
                # Level 0 (no bullet) or Level 1 (main bullet)
                if level == 0 or level == 1:
                    p.left_indent = Inches(0.21)  # 0.21" indent before text
                    p.first_line_indent = Inches(-0.19)  # 0.19" special hanging
                    p.space_before = Pt(0)  # 0pt spacing before
                    p.space_after = Pt(0)  # 0pt spacing after
                    p.line_spacing = 1.0  # Single line spacing
                elif level == 2:
                    # Level 2 - more indented
                    p.left_indent = Inches(0.4)
                    p.first_line_indent = Inches(-0.19)
                    p.space_before = Pt(0)
                    p.space_after = Pt(0)
                    p.line_spacing = 1.0
                elif level == 3:
                    # Level 3 - most indented
                    p.left_indent = Inches(0.6)
                    p.first_line_indent = Inches(-0.19)
                    p.space_before = Pt(0)
                    p.space_after = Pt(0)
                    p.line_spacing = 1.0
            except:
                pass  # Silently handle formatting errors
            
            # Add text with proper formatting
            run = p.add_run()
            run.text = text
            run.font.size = get_font_size_for_text(text, force_chinese_mode=is_chinese)
            run.font.name = get_font_name_for_text(text)
            
            # Apply level-specific formatting
            if level == 1:
                run.font.bold = True
                try:
                    run.font.color.rgb = RGBColor(0, 51, 102)  # Dark blue for level 1
                except:
                    pass
            elif level == 0:
                # Regular text - no special formatting
                pass
            
            lines_added += 1
    
    def _validate_ai_summary(
        self,
        source_text: str,
        draft_summary: str,
        is_chinese: bool,
        ai_helper: Optional[Any] = None,
    ) -> Optional[str]:
        summary_settings = self._summary_settings()
        if not bool(summary_settings.get("enable_validation", True)):
            return draft_summary

        try:
            from fdd_utils.ai import AIClient

            model_type = ai_helper.model_type if ai_helper is not None else self._resolve_summary_model_type(is_chinese)
            if model_type == "local" and not bool(summary_settings.get("local_enable_validation", False)):
                logger.info("Skipping PPTX summary validation for local model; using draft summary directly")
                return draft_summary

            max_input_chars = int(summary_settings.get("max_input_chars", 1400))
            validation_max_tokens = int(summary_settings.get("validation_max_tokens", 90))
            max_numeric_sentences = int(summary_settings.get("max_numeric_sentences", 1))
            validation_timeout_seconds = float(summary_settings.get("validation_timeout_seconds", 25) or 25)
            target_chars_chi = int(summary_settings.get("target_chars_chi", 120))
            target_words_eng = int(summary_settings.get("target_words_eng", 95))

            if is_chinese:
                prompt = f"""请校验并压缩以下PPT执行摘要草稿，使其适合作为财务PPT摘要框内容。

要求：
1. 只保留与原始评论一致的高层结论、趋势和核心驱动。
2. 控制在3句话以内，长度约{target_chars_chi}字。
3. 最多保留{max_numeric_sentences}个数字或百分比，除非删除后会影响结论准确性。
4. 删除重复、堆叠金额和逐项罗列，将草稿进一步压缩到当前信息量的大约70%，但不得丢失核心趋势、驱动和结论。
5. 优先合并重复句、删去铺垫和次要背景，只保留最重要的业务含义。
6. 只输出最终摘要，不要解释。

原始评论：
{source_text[:max_input_chars]}

摘要草稿：
{draft_summary}"""
            else:
                prompt = f"""Validate and tighten the draft executive summary for a financial PPT summary box.

Requirements:
1. Keep only source-supported themes, trend, and core driver.
2. Limit the result to no more than 3 sentences and about {target_words_eng} words.
3. Keep at most {max_numeric_sentences} number or percentage unless removing it would make the summary inaccurate.
4. Compress the draft to roughly 70% of its current information density without losing the key trend, driver, or conclusion.
5. Remove repeated phrasing, stacked figures, scene-setting language, and account-by-account detail.
6. Output only the final validated summary paragraph.

Source commentary:
{source_text[:max_input_chars]}

Draft summary:
{draft_summary}"""

            ai_helper = ai_helper or AIClient(
                model_type=model_type,
                language='Chi' if is_chinese else 'Eng',
            )
            validation_max_retries = int(summary_settings.get("validation_max_retries", 2) or 2)
            response = self._call_with_timeout_retry(
                lambda: ai_helper.get_response(
                    user_prompt=prompt,
                    system_prompt=(
                        "You validate executive summaries for financial presentation slides. "
                        "Keep only source-supported, concise, presentation-ready conclusions."
                    ),
                    temperature=float(summary_settings.get("validation_temperature", 0.1) or 0.1),
                    max_tokens=validation_max_tokens,
                ),
                timeout_seconds=validation_timeout_seconds,
                max_retries=validation_max_retries,
                timeout_label="PPTX summary validation",
            )
            validated_summary = str((response or {}).get("content") or "").strip()
            if _looks_like_blocked_ai_content(validated_summary):
                logger.warning("PPTX summary validation returned blocked/network HTML content; using draft summary fallback")
                return draft_summary
            return validated_summary or draft_summary
        except Exception as exc:
            logger.warning("Could not validate AI summary: %s", exc)
            return draft_summary

    def _generate_ai_summary(self, commentary: str, summary_source: str, is_chinese: bool) -> Optional[str]:
        """Generate and validate AI summary from page commentary."""
        try:
            from fdd_utils.ai import AIClient
            summary_settings = self._summary_settings()
            if not bool(summary_settings.get("enable_ai", True)):
                logger.info("PPTX summary AI disabled by config; using fallback summary")
                return None
            model_type = self._resolve_summary_model_type(is_chinese)
            max_input_chars = int(summary_settings.get("max_input_chars", 1600))
            max_tokens = int(summary_settings.get("max_tokens", 180))
            max_numeric_sentences = int(summary_settings.get("max_numeric_sentences", 1))
            generation_timeout_seconds = float(summary_settings.get("generation_timeout_seconds", 45) or 45)
            target_chars_chi = int(summary_settings.get("target_chars_chi", 120))
            target_words_eng = int(summary_settings.get("target_words_eng", 95))
            source_text = str(commentary or summary_source or "").strip()
            if not source_text:
                return None

            if is_chinese:
                prompt = f"""请将以下财务评论改写成适合PPT摘要框的高层执行摘要。

目标长度：约{target_chars_chi}字，控制在3句话以内。

要求：
1. 仅保留高层结论、趋势和核心驱动。
2. 除非极其必要，最多保留{max_numeric_sentences}个数字或百分比。
3. 不要逐项复述账户，不要堆叠金额细节。
4. 语气要像管理层摘要，压缩成一个紧凑自然的短段落，整体信息量控制在当前评论的大约70%。
5. 优先删去次要说明、重复铺垫和账户层级细节，只保留最重要的业务结论、驱动和影响。

原始内容：
{source_text[:max_input_chars]}"""
            else:
                prompt = f"""Write a short executive summary for a PPT summary box based on the following financial commentary.

Target length: about {target_words_eng} words, with no more than 3 sentences.

Requirements:
1. Focus on overall trend, key driver, and business implication.
2. Keep it high level and presentation-friendly.
3. Include at most {max_numeric_sentences} number or percentage unless absolutely necessary.
4. Do not list account-by-account detail or repeat many figures.
5. Write one compact management-style paragraph and trim the information load to roughly 70% of the source while preserving the key trend, driver, and implication.
6. Remove secondary detail, scene-setting language, and repeated wording.

Original content:
{source_text[:max_input_chars]}"""

            ai_helper = AIClient(
                model_type=model_type,
                language='Chi' if is_chinese else 'Eng',
            )
            generation_max_retries = int(summary_settings.get("generation_max_retries", 3) or 3)
            response = self._call_with_timeout_retry(
                lambda: ai_helper.get_response(
                    user_prompt=prompt,
                    system_prompt=(
                        "You write concise executive summaries for financial presentation slides. "
                        "Prefer themes, drivers, and implications over detailed figures."
                    ),
                    temperature=float(summary_settings.get("generation_temperature", 0.2) or 0.2),
                    max_tokens=max_tokens,
                ),
                timeout_seconds=generation_timeout_seconds,
                max_retries=generation_max_retries,
                timeout_label="PPTX summary generation",
            )
            summary = str((response or {}).get("content") or "").strip()
            if _looks_like_blocked_ai_content(summary):
                logger.warning(
                    "PPTX summary generation returned blocked/network HTML content; falling back to compact summary"
                )
                return None

            if summary:
                return self._validate_ai_summary(source_text, summary, is_chinese, ai_helper=ai_helper)
        except Exception as e:
            logger.warning("Could not generate AI summary: %s", e)
            logger.debug(traceback.format_exc())
        
        return None
    
    def _generate_page_summary(self, commentary: str, is_chinese: bool) -> str:
        """
        Generate a per-page summary from commentary text
        This is a fallback when AI summary is not available
        Extracts first few complete sentences without hard truncation
        """
        if not commentary or not commentary.strip():
            return ""
        is_chinese_text = is_chinese or detect_chinese_text(commentary)
        summary_settings = self._summary_settings()
        return _build_compact_summary_text(
            commentary,
            is_chinese=is_chinese_text,
            max_sentences=int(summary_settings.get("max_sentences_chi" if is_chinese_text else "max_sentences_eng", 3)),
            max_chars=(
                int(summary_settings.get("target_chars_chi", 120))
                if is_chinese_text
                else int(summary_settings.get("target_words_eng", 95)) * 6
            ),
            max_numeric_sentences=int(summary_settings.get("max_numeric_sentences", 1)),
        ).strip()

    def embed_financial_tables(
        self,
        excel_path: str,
        sheet_name: str,
        project_name: str,
        language: str,
        bs_is_results: Optional[Dict[str, Any]] = None,
    ):
        """Embed financial tables: BS to page 1, IS to page 5"""
        try:
            import pandas as pd
            from fdd_utils.workbook import extract_balance_sheet_and_income_statement
            
            logger.info("Embedding financial tables from %s, sheet: %s", excel_path, sheet_name)
            
            # Validate inputs
            if not excel_path or not sheet_name:
                logger.warning("Missing excel_path (%s) or sheet_name (%s), skipping table embedding", excel_path, sheet_name)
                return
            
            # Prefer raw (non-multiplied) values for the PPTX table so the
            # header unit label (CNY'000 / 人民币千元) matches what the cells
            # display. Upstream (process_workbook_data) extracts with
            # multiply_values=True for the AI pipeline. Try a fresh raw
            # extract here; if it fails or returns empty, fall back to the
            # precomputed results so the BS table still makes it into the
            # deck — we just have to compensate in the display step.
            precomputed_bs_is_results = bs_is_results
            bs_is_results = None
            try:
                logger.info("Re-extracting BS/IS in raw units (multiply_values=False) for table")
                bs_is_results = extract_balance_sheet_and_income_statement(
                    excel_path,
                    sheet_name,
                    debug=False,
                    multiply_values=False,
                )
            except Exception as exc:
                logger.warning(
                    "Raw BS/IS re-extraction failed (%s); will fall back to precomputed results",
                    exc,
                )

            values_pre_multiplied = False
            if not bs_is_results or (
                bs_is_results.get("balance_sheet") is None
                and bs_is_results.get("income_statement") is None
            ):
                if precomputed_bs_is_results:
                    logger.info(
                        "Falling back to precomputed BS/IS for the PPTX table; "
                        "values will be rescaled to match the header unit."
                    )
                    bs_is_results = precomputed_bs_is_results
                    values_pre_multiplied = True
                else:
                    logger.warning("No BS/IS data extracted")
                    return
            
            # Extract BS and IS DataFrames from results
            # The structure should have 'balance_sheet' and 'income_statement' keys with DataFrames
            bs_df = bs_is_results.get('balance_sheet')
            is_df = bs_is_results.get('income_statement')
            
            # Table titles follow the standard FDD phrasing regardless of what
            # the source Excel calls the sheet. Language-aware so Chinese decks
            # stay consistent with English decks.
            is_chinese_mode = str(language or "").strip().lower().startswith(("chi", "zh", "cn"))
            project_suffix = f" - {project_name}" if project_name else ""
            if is_chinese_mode:
                bs_table_name = f"示意性调整后资产负债表{project_suffix}"
                is_table_name = f"示意性调整后利润表{project_suffix}"
            else:
                bs_table_name = f"Indicative adjusted balance sheet{project_suffix}"
                is_table_name = f"Indicative adjusted income statement{project_suffix}"

            # Detect currency unit from the sheet so the header column can
            # show the correct scale (thousands vs millions).
            currency_unit = None
            try:
                excel_df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
                for idx, row in excel_df.iterrows():
                    row_text = ' '.join([str(cell) for cell in row if pd.notna(cell)])
                    # Order matters: check million markers before generic 000
                    # markers to avoid mislabelling (e.g. a table that literally
                    # reads "CNY million" shouldn't be called CNY'000).
                    if currency_unit is None:
                        if '人民币百万' in row_text:
                            currency_unit = '人民币百万'
                        elif "CNY'M" in row_text or 'CNY million' in row_text or 'CNY mn' in row_text.lower():
                            currency_unit = "CNY'M"
                        elif '人民币千元' in row_text:
                            currency_unit = '人民币千元'
                        elif "CNY'000" in row_text or "CNY 000" in row_text:
                            currency_unit = "CNY'000"
                    if currency_unit is not None:
                        break
            except Exception:
                pass

            logger.info("Extracted BS: %s, IS: %s", bs_df.shape if bs_df is not None else 'None', is_df.shape if is_df is not None else 'None')
            logger.info("Table names - BS: %s, IS: %s, Currency: %s", bs_table_name, is_table_name, currency_unit)

            # If the values came from the precomputed (multiply_values=True)
            # pipeline, rescale to the source unit so the cells match the
            # header. "CNY'000" / "人民币千元" → divide by 1000,
            # "CNY'M" / "人民币百万" → divide by 1_000_000.
            if values_pre_multiplied and currency_unit:
                rescale = None
                if "千" in currency_unit or "'000" in currency_unit or "000" in currency_unit:
                    rescale = 1000.0
                elif "百万" in currency_unit or "'M" in currency_unit or "million" in currency_unit.lower():
                    rescale = 1_000_000.0
                if rescale is not None:
                    logger.info("Rescaling precomputed values by 1/%s to match unit %s", int(rescale), currency_unit)
                    for _df in (bs_df, is_df):
                        if _df is None or _df.empty:
                            continue
                        for _col in _df.columns:
                            if pd.api.types.is_numeric_dtype(_df[_col]):
                                _df[_col] = _df[_col] / rescale
            
            # Embed BS table to slide 0 (page 1)
            if bs_df is not None and not bs_df.empty and len(self.presentation.slides) > 0:
                slide_0 = self.presentation.slides[0]
                logger.info("Looking for table shape on slide 1, available shapes: %s", [s.name if hasattr(s, 'name') else type(s).__name__ for s in slide_0.shapes])
                self._embed_statement_table(
                    slide_0,
                    bs_df,
                    "BS",
                    table_name=bs_table_name,
                    currency_unit=currency_unit,
                )
            else:
                logger.warning("BS DataFrame is None or empty, skipping BS table embedding")
            
            # Embed IS table to slide 5 (1-indexed) = slide index 4 (0-indexed)
            # IS slides are 5-8 (1-indexed) = indices 4-7 (0-indexed)
            # Just paste it to slide index 4 (page 5) - unused slides will be removed later
            if is_df is not None and not is_df.empty:
                is_slide_idx = 4  # First IS slide (page 5, 1-indexed)
                
                logger.info("Looking for IS table on slide %s (index %s), total slides: %s", is_slide_idx + 1, is_slide_idx, len(self.presentation.slides))
                
                if len(self.presentation.slides) > is_slide_idx:
                    is_slide = self.presentation.slides[is_slide_idx]
                    logger.info("Slide %s exists, searching for Table Placeholder...", is_slide_idx + 1)
                    self._embed_statement_table(
                        is_slide,
                        is_df,
                        "IS",
                        table_name=is_table_name,
                        currency_unit=currency_unit,
                    )
                else:
                    logger.error("Slide %s does not exist for IS table (only %s slides)", is_slide_idx + 1, len(self.presentation.slides))
                    logger.error("IS data should be on slide 5, but presentation only has %s slides", len(self.presentation.slides))
                    
        except Exception as e:
            logger.error("Error embedding financial tables: %s", e)
            logger.error(traceback.format_exc())

    def save(self, output_path: str):
        """Save the presentation"""
        if not self.presentation:
            raise ValueError("No presentation loaded")

        # Ensure output directory exists
        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        self.presentation.save(output_path)
        logger.info("Presentation saved to: %s", output_path)
# --- end pptx/generation.py ---
