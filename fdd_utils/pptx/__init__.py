"""Split from the former single-file module. Re-exports every name
the flat module exposed, so `from fdd_utils.pptx import X` is unchanged."""
from __future__ import annotations

# namespace parity with the pre-split flat module
from concurrent.futures import ThreadPoolExecutor, as_completed
import re
from typing import Optional
from pptx.util import Pt
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
import copy
import logging
import os
import posixpath
import time
import traceback
from typing import Dict, List, Optional
from pptx import Presentation
from pptx.oxml.ns import qn
import threading
from typing import Any, Dict, List, Optional, Tuple
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE_TYPE  # noqa: F401


from .text import (  # noqa: F401
    clean_content_quotes,
    detect_chinese_text,
    get_font_name_for_text,
    get_font_size_for_text,
    get_line_spacing_for_text,
    get_space_after_for_text,
    get_space_before_for_text,
)
from .payloads import (  # noqa: F401
    PPTX_DEFAULT_SETTINGS,
    build_pptx_structured_payloads,
    shorten_company_names,
    _LEGAL_FORM_TAILS,
    _REGISTRATION_BRACKET,
    _TRADITIONAL_TO_SIMPLIFIED,
    _TRADITIONAL_TO_SIMPLIFIED_PAIRS,
    _build_statement_order,
    _extract_final_content,
    _extract_summary,
    _find_chinese_display_name,
    _has_significant_balance,
    _join_text_sentences,
    _load_pptx_settings,
    _looks_like_blocked_ai_content,
    _merge_nested_dict,
    _normalize_slide_commentary_text,
    _sentence_is_numeric_heavy,
    _split_text_sentences,
    _translate_statement_row_label,
)
from .exporters import (  # noqa: F401
    ReportGenerator,
    combine_presentations,
    export_pptx,
    export_pptx_from_structured_data,
    export_pptx_from_structured_data_combined,
    logger,
    merge_presentations,
    _copy_slide_into,
    _dedupe_part_name,
)
from .generation import (  # noqa: F401
    PowerPointGenerator,
    logger,
)
