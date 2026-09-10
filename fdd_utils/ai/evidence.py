"""One account's working memory: the facts it was graded against, and what it was shown.

Three verification defects on one real run had the same shape -- the checker
saw less than the model. A remainder the prompt computed; a period column that
lived only in the analysis frame; a movement the model subtracted from a table
it was handed. Each was patched by copying one more derived value from the
prompt side to the verifier side. The structural answer is to stop having two
sides: build the pool ONCE per account, grade every stage against that same
object, and write it to disk so a figure can be traced after the process ends.

`AccountEvidence` is that object. It is not a new derivation -- `SourceIndex`
already builds facts with provenance from the frame the Generator is shown --
it is the same facts, held once, given a record of what was rendered, and
persisted. Nothing here crosses a run boundary.

What it carries:

    facts   the SourceIndex fact records (value, kind, sheet, row, column),
            the pool every stage's verify_commentary grounds against
    dates   the allowed date facts, same provenance
    shown   what the Generator was rendered: prompt hash and size, which
            guidance sections were present, how many components the budget
            kept and dropped

What it does not do: change any verdict. `SourceIndex.from_evidence(ev)` yields
the same pool `from_df` would have built for the same frame at the same moment;
the point is that the moment is one and the pool is one.
"""
from __future__ import annotations

import hashlib
import json
import os
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

from .validator import SourceIndex
from .logging import coerce_plain

EVIDENCE_DIR = "evidence"

# Guidance section markers, matched on the rendered prompt so `shown` records
# what actually reached the model rather than what a builder intended.
_SECTION_MARKERS = (
    ("remainder", ("【余额差额已算好】", "[REMAINDER ALREADY COMPUTED]")),
    ("not_shown", ("【已省略的较小构成项】", "[SMALLER COMPONENTS NOT SHOWN]")),
    ("hierarchy", ("【本科目已核对的层级关系】", "[VERIFIED HIERARCHY")),
    ("component_nature", ("【构成项性质", "[COMPONENT NATURE")),
    ("material_movement", ("【重大变动提示】", "[MATERIAL MOVEMENT]")),
    ("data_insight", ("【数据洞察", "[DATA INSIGHT")),
    # The data blocks appear as markdown headings in one data_format and as
    # JSON keys in the other (config.example.yml ships json), so both spellings
    # are listed; a section is "shown" if either form reached the prompt.
    ("trend_summary", ("趋势摘要", "Trend summary", '"trend_summary"')),
    ("significant_movements", ("重大变动（", "Significant movements", '"significant_movements"')),
    ("table_remarks", ("表格关联备注", "Table context observations", '"table_context_observations"')),
    ("supporting_context", ("补充备注", "Supporting notes", '"supporting_context"')),
    ("cross_account", ("【跨科目", "[CROSS-ACCOUNT")),
)


@dataclass
class AccountEvidence:
    mapping_key: str
    language: str = ""
    statement_type: str = ""
    facts: List[Dict[str, Any]] = field(default_factory=list)
    dates: List[Dict[str, Any]] = field(default_factory=list)
    shown: Dict[str, Any] = field(default_factory=dict)

    # -- the pool ------------------------------------------------------------

    def source_index(self) -> SourceIndex:
        """The grounding pool, rebuilt from these records and nothing else."""
        index = SourceIndex(list(self.facts))
        index.date_facts = list(self.dates)
        return index

    @property
    def pool_size(self) -> int:
        return len(self.facts)

    # -- persistence ---------------------------------------------------------

    def to_dict(self) -> Dict[str, Any]:
        return coerce_plain({
            "mapping_key": self.mapping_key,
            "language": self.language,
            "statement_type": self.statement_type,
            "shown": dict(self.shown),
            "pool_size": self.pool_size,
            "facts": list(self.facts),
            "dates": list(self.dates),
        })

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "AccountEvidence":
        return cls(
            mapping_key=str(data.get("mapping_key") or ""),
            language=str(data.get("language") or ""),
            statement_type=str(data.get("statement_type") or ""),
            facts=list(data.get("facts") or []),
            dates=list(data.get("dates") or []),
            shown=dict(data.get("shown") or {}),
        )


# -- building ----------------------------------------------------------------

def describe_shown(mapping_key: str, df: Any, system_prompt: str, user_prompt: str) -> Dict[str, Any]:
    """What the Generator was rendered, as a record. Cheap; no frame walk."""
    from .client import AIClient
    text = (user_prompt or "")
    sections = [name for name, markers in _SECTION_MARKERS if any(m in text for m in markers)]
    attrs = getattr(df, "attrs", None) or {}
    analysis = attrs.get("prompt_analysis_df")
    components = list(getattr(analysis, "attrs", {}).get("component_descriptions") or []) if analysis is not None else []
    budget = attrs.get("prompt_budget") if isinstance(attrs.get("prompt_budget"), dict) else None
    dropped = list((budget or {}).get("components_dropped") or [])
    return {
        "prompt_sha256": hashlib.sha256(((system_prompt or "") + "\n" + text).encode("utf-8")).hexdigest(),
        "prompt_chars": len(system_prompt or "") + len(text),
        "prompt_tokens_est": AIClient._estimate_text_tokens(system_prompt) + AIClient._estimate_text_tokens(text),
        "sections": sections,
        "components_total": len(dict.fromkeys(components)),
        "components_dropped": len(dropped),
        "budget": ({k: budget.get(k) for k in ("limit", "headroom", "estimate", "fits")} if budget else None),
    }


def compile_account_evidence(
    mapping_key: str,
    df: Any,
    *,
    language: str = "",
    statement_type: str = "",
    sibling_dfs: Optional[List[Any]] = None,
    shown: Optional[Dict[str, Any]] = None,
) -> AccountEvidence:
    """Build the pool once, from the frame as it stands NOW.

    Call this after the Generator prompt has been rendered for the account:
    rendering is what stashes the computed remainder and the budget residual on
    the analysis frame, and those must be in the pool. Calling it earlier
    builds the pool the three defects were built on.
    """
    index = SourceIndex.from_df(df, sibling_dfs=sibling_dfs)
    return AccountEvidence(
        mapping_key=str(mapping_key),
        language=str(language or ""),
        statement_type=str(statement_type or ""),
        facts=list(index.facts),
        dates=list(index.date_facts),
        shown=dict(shown or {}),
    )


# -- files -------------------------------------------------------------------

def _safe_name(key: str) -> str:
    return re.sub(r"[^\w一-鿿-]+", "_", str(key)).strip("_") or "account"


def write_evidence(run_folder: str, evidence: AccountEvidence) -> str:
    folder = os.path.join(run_folder, EVIDENCE_DIR)
    os.makedirs(folder, exist_ok=True)
    path = os.path.join(folder, _safe_name(evidence.mapping_key) + ".json")
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(evidence.to_dict(), fh, ensure_ascii=False, indent=None)
    return path


def read_evidence(run_folder: str, mapping_key: str) -> Optional[AccountEvidence]:
    path = os.path.join(run_folder, EVIDENCE_DIR, _safe_name(mapping_key) + ".json")
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as fh:
        return AccountEvidence.from_dict(json.load(fh))


def list_evidence(run_folder: str) -> Dict[str, AccountEvidence]:
    folder = os.path.join(run_folder, EVIDENCE_DIR)
    out: Dict[str, AccountEvidence] = {}
    if not os.path.isdir(folder):
        return out
    for name in sorted(os.listdir(folder)):
        if not name.endswith(".json"):
            continue
        with open(os.path.join(folder, name), encoding="utf-8") as fh:
            ev = AccountEvidence.from_dict(json.load(fh))
        out[ev.mapping_key] = ev
    return out
