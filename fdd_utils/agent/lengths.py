"""How long finished team reports write an account, so the writer's LENGTH follows the team's practice for that account and not
how much evidence the account has.

The numbers are in fdd_utils/report_lengths.yml, one per finished report, built by ad-hoc/workbench/report_lengths.py from text
dumps of finished decks. The team's lengths do not follow the size of the balance: in the first report measured (2026-11), the
largest liability got 144 characters and a small deferred-cost account 265, because it carried a proposed adjustment; a P&L
account whose slide carries a detail table got a lead-in of 31-64. An account the table does not list keeps the evidence rule
in controller.length_target.
"""
from __future__ import annotations

import statistics
from functools import lru_cache
from pathlib import Path
from typing import Dict, List, Optional

import yaml

PATH = Path(__file__).resolve().parents[1] / "report_lengths.yml"


@lru_cache(maxsize=1)
def table() -> Dict[str, List[int]]:
    """{account: [characters in each finished report]}; empty when the file is missing or unreadable."""
    try:
        data = yaml.safe_load(PATH.read_text(encoding="utf-8")) or {}
    except (OSError, yaml.YAMLError):
        return {}
    out: Dict[str, List[int]] = {}
    for name, seen in (data.get("accounts") or {}).items():
        values = [int(n) for n in (seen or []) if isinstance(n, (int, float)) and n > 0]
        if values:
            out[str(name)] = values
    return out


def written_length(account: str) -> Optional[int]:
    """The median length finished reports gave this account, or None when none of them wrote it."""
    seen = table().get(str(account or "").strip())
    return int(statistics.median(seen)) if seen else None
