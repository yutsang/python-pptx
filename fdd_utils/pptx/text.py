from __future__ import annotations

# re-added: bound by an import in another section of the pre-split module
from ..financial_common import (
    contains_chinese_text,
    contains_predominantly_chinese_text,
    get_pipeline_result_text,
    load_yaml_file,
    package_file_path,
)

from concurrent.futures import ThreadPoolExecutor, as_completed
import re
from typing import Optional

from pptx.util import Pt




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
    # Deck-wide typography: every commentary run, every slide, every
    # language renders at a single fixed size. We intentionally ignore the
    # text, base_size, and force_chinese_mode arguments — any caller that
    # asked for something else would reintroduce the size-variation bug.
    return Pt(9)


def get_font_name_for_text(text: str, default_font: str = "Arial") -> str:
    # Same philosophy: one font for the whole deck. Arial has CJK fallback
    # glyphs via the system's default font substitution, so Chinese content
    # still renders correctly without switching to Microsoft YaHei (which
    # would change glyph width / baseline on some slides).
    return "Arial"


def get_line_spacing_for_text(text: str, force_chinese_mode: bool = False) -> float:
    return 0.9 if detect_chinese_text(text, force_chinese_mode) else 1.0


def get_space_after_for_text(text: str, force_chinese_mode: bool = False) -> Pt:
    return Pt(6) if detect_chinese_text(text, force_chinese_mode) else Pt(4)


def get_space_before_for_text(text: str, force_chinese_mode: bool = False) -> Pt:
    return Pt(3) if detect_chinese_text(text, force_chinese_mode) else Pt(2)


# --- end pptx/text.py ---
