"""Does every number the Generator is SHOWN ground against the verifier's pool?

Three verification defects on one real portfolio run had one shape: the checker
saw less than the model. A remainder the prompt computed and told the model to
quote; a period column that lived only in the analysis frame; a movement the
model subtracted from a table it was handed. Each was flagged as a
hallucination, each cost a 5-10 minute real run to find, and one shipped a
corrected date into a deck before it was found.

This probe finds the whole class offline. For every account it renders the
Generator prompt, builds the SAME SourceIndex verify_commentary would build, and
asks whether each amount and each date in the prompt would pass. An amount the
prompt states that the pool cannot ground is a future false positive, before a
single token is spent.

    python ad-hoc/databook-probes/probe_prompt_grounds_itself.py <databook.xlsx> [entity]
    python ad-hoc/databook-probes/probe_prompt_grounds_itself.py <databook.xlsx> --account 固定资产

Exit status 1 when anything is ungrounded, so it can gate a commit. Free.

Bare numbers and percentages are reported separately and never counted: the
verifier deliberately does not treat them as amounts, so a wrong ratio is a
prompt-quality question, not a grounding one.

Output names accounts and prints figures; do not put it anywhere public.
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import re
from typing import Any, Dict, List, Tuple

from fdd_utils.ai.prompts import PromptEngine
from fdd_utils.ai.validator import (
    SourceIndex,
    _dates_in,
    _harvest_source_dates,
    extract_amount_spans,
)
from fdd_utils.workbook import process_workbook_data

RULE = "=" * 78
_PERCENT = re.compile(r"\d+(?:\.\d+)?\s*%")

# A figure inside a worked example is not data. The style guidance says things
# like 'e.g., "59.3 million", "54,950"' and '例如"1,062.8万元"', and those never
# ground because they are not about this account. They are reported apart from
# the data figures because they are a DIFFERENT known hazard -- a model copies
# a worked example's unit (a hardcoded 万元 in one example once shipped 3,091元
# as 3,091.0万元) -- and conflating the two would bury the grounding gaps this
# probe exists to surface under a pile of style-text hits.
_EXAMPLE_CUES = ("e.g.", "e.g,", "例如", "such as", "for example", "（如", "(如", "如'", '如"',
                 "写成", "写法", "Render with", "Never use", "must render", "不可写成", "错误写法",
                 "正确写法", "应写成", "should read", "instead of")


_THRESHOLD_CUES = ("<", ">", "≥", "≤", "不足", "低于", "超过", "以上", "以下", "达到", "under ", "over ", "below ", "above ")


def _is_example(prompt: str, start: int, lookback: int = 90) -> bool:
    window = prompt[max(0, start - lookback):start]
    if any(cue in window for cue in _EXAMPLE_CUES):
        return True
    # A materiality threshold ("Aggregate <CNY100K items", "金额达到1亿以上")
    # is a rule, not a figure about this account.
    tail = window[-12:]
    return any(cue in tail for cue in _THRESHOLD_CUES)


def _is_table_cell(prompt: str, start: int, end: int) -> bool:
    """A grouped number sitting in a markdown table row or a trend line is a
    SOURCE cell rendered in the section's display unit ('4,289.2' under a
    heading that says 人民币万元). extract_amount_spans reads it as 4,289 yuan
    because in commentary an unlabelled grouped number IS yuan -- but the model
    writes such a figure with its unit ('4,289.2万元'), which extracts as
    42,892,000 and grounds. The bare table cell is the pool itself, so testing
    it against the pool at the wrong scale measures nothing. Excluded, counted."""
    line_start = prompt.rfind("\n", 0, start) + 1
    line_end = prompt.find("\n", end)
    line = prompt[line_start:line_end if line_end != -1 else len(prompt)]
    stripped = line.lstrip()
    if stripped.startswith("|"):
        return True
    return any(tok in line for tok in ("delta=", "_value:", "start_value", "end_value", "net_change"))


def _user_prompt(pe: PromptEngine, language: str, key: str, df) -> str:
    _sys_p, usr_p = pe.render_prompt("subagent_1", language, key, df, data_format="markdown")
    return usr_p or ""


def _context(text: str, start: int, end: int, width: int = 28) -> str:
    lo, hi = max(0, start - width), min(len(text), end + width)
    snippet = text[lo:hi].replace("\n", "⏎")
    return f"…{snippet}…"


def probe_account(pe: PromptEngine, language: str, key: str, df) -> Dict[str, Any]:
    prompt = _user_prompt(pe, language, key, df)
    index = SourceIndex.from_df(df)
    allowed_dates = _harvest_source_dates(df)

    amounts = extract_amount_spans(prompt)
    ungrounded_amounts: List[Tuple[float, str]] = []
    example_amounts: List[Tuple[float, str]] = []
    table_cells = 0
    for value, start, end in amounts:
        if index.matches(value) is not None:
            continue
        if _is_table_cell(prompt, start, end):
            table_cells += 1
            continue
        if _is_example(prompt, start):
            example_amounts.append((value, _context(prompt, start, end)))
        else:
            ungrounded_amounts.append((value, _context(prompt, start, end)))

    dates_in_prompt = set()
    for m in re.finditer(r"\d{4}[-年/.]\d{1,2}[-月/.]\d{1,2}日?|\d{4}-\d{2}-\d{2}", prompt):
        for ymd in _dates_in(m.group(0)):
            dates_in_prompt.add((ymd, m.start(), m.end()))
    ungrounded_dates = [
        (f"{y:04d}-{mo:02d}-{d:02d}", _context(prompt, s, e))
        for (y, mo, d), s, e in sorted(dates_in_prompt)
        if (y, mo, d) not in allowed_dates
    ]

    return {
        "key": key,
        "n_amounts": len(amounts),
        "ungrounded_amounts": ungrounded_amounts,
        "example_amounts": example_amounts,
        "table_cells": table_cells,
        "n_dates": len(dates_in_prompt),
        "ungrounded_dates": ungrounded_dates,
        "n_percents": len(_PERCENT.findall(prompt)),
        "pool_size": len(getattr(index, "facts", []) or []),
    }


def main(path: str, entity: str = "", only: str = "") -> int:
    state = process_workbook_data(temp_path=path, entity_name=entity, selected_sheet=None)
    dfs = state["dfs"]
    language = state.get("language") or "Eng"
    pe = PromptEngine()

    keys = [k for k in dfs if not only or str(k).strip() == only.strip()]
    if not keys:
        print(f"no account named {only!r}. Present: {sorted(dfs)}")
        return 2

    print(f"{RULE}\n  PROMPT GROUNDS ITSELF — language={language}, {len(keys)} account(s)\n{RULE}")
    print(f"  {'account':<14}{'amounts':>8}{'ungrnd':>8}{'exmpl':>7}{'cells':>7}{'dates':>7}{'ungrnd':>8}{'%':>5}{'pool':>7}")

    rows = []
    for key in keys:
        try:
            r = probe_account(pe, language, key, dfs[key])
        except Exception as exc:  # noqa: BLE001
            print(f"  {str(key):<14}  RENDER FAILED: {type(exc).__name__}: {exc}")
            continue
        rows.append(r)
        ua, ud, ex = len(r["ungrounded_amounts"]), len(r["ungrounded_dates"]), len(r["example_amounts"])
        flag = "  <--" if (ua or ud) else ""
        print(f"  {str(key):<14}{r['n_amounts']:>8}{ua:>8}{ex:>7}{r['table_cells']:>7}{r['n_dates']:>7}{ud:>8}"
              f"{r['n_percents']:>5}{r['pool_size']:>7}{flag}")

    total_a = sum(r["n_amounts"] for r in rows)
    total_ua = sum(len(r["ungrounded_amounts"]) for r in rows)
    total_d = sum(r["n_dates"] for r in rows)
    total_ud = sum(len(r["ungrounded_dates"]) for r in rows)

    shown = 0
    for r in rows:
        if not (r["ungrounded_amounts"] or r["ungrounded_dates"]):
            continue
        print(f"\n  [{r['key']}]")
        for value, ctx in r["ungrounded_amounts"]:
            print(f"    amount {value:,.0f}   {ctx}")
            shown += 1
        for iso, ctx in r["ungrounded_dates"]:
            print(f"    date   {iso}   {ctx}")
            shown += 1
        if shown >= 60:
            print("\n  ... stopped listing at 60; the counts above are complete")
            break

    total_ex = sum(len(r["example_amounts"]) for r in rows)
    total_cells = sum(r["table_cells"] for r in rows)
    examples = sorted({v for r in rows for v, _c in r["example_amounts"]})
    checked = total_a - total_ex - total_cells
    print(f"\n{RULE}")
    print(f"  {checked - total_ua} of {checked} prompt DATA amounts ground; "
          f"{total_d - total_ud} of {total_d} prompt dates allowed.")
    if total_cells:
        print(f"  {total_cells} bare table cell(s) in the section's display unit excluded "
              f"(they ARE the pool; a model quotes them with the unit).")
    if total_ex:
        print(f"  {total_ex} worked-example figure(s) excluded from that count: "
              + ", ".join(f"{v:,.0f}" for v in examples[:8])
              + (" …" if len(examples) > 8 else ""))
        print("  (a model can copy an example's unit -- a separate, known hazard; see")
        print("   the 8442f36 note in memory. Not a grounding gap.)")
    if total_ua or total_ud:
        print("  Every item above is either a fact the verifier lacks (widen the pool from")
        print("  the SAME source the prompt used, never by re-deriving) or a figure the")
        print("  prompt should not be stating as a bare amount (a prompts.py defect).")
        print("  Classify each; do not change this probe to make it pass.")
    else:
        print("  Nothing the model is shown would be flagged. The invariant holds today.")
    print(RULE)
    return 1 if (total_ua or total_ud) else 0


if __name__ == "__main__":
    args = [a for a in sys.argv[1:]]
    only = ""
    if "--account" in args:
        i = args.index("--account")
        only = args[i + 1] if i + 1 < len(args) else ""
        args = [a for a in args if a not in ("--account", only)]
    if not args:
        sys.exit(__doc__)
    sys.exit(main(args[0], args[1] if len(args) > 1 else "", only))
