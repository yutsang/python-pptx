#!/usr/bin/env python3
"""Pairs each account's SOURCE material in a databook with the COMMENTARY
that actually came out in the exported deck, so the writing style can be
judged against what the model was given rather than in isolation.

inspect_is_variance.py looks only at the databook; inspect_pptx.py looks
only at the deck. Neither answers the question this one is for: given
these remarks, is the output using them, and how is it phrasing them?

Per account it prints the substantive source remarks, the generated
bullet, whether each remark visibly reached the output, and any style
flags. Then a summary covering remark uptake and the style issues found
across the whole deck.

Style flags (each is a worklist item, not an automatic fail):
  * source meta-reference -- '根据备注信息...' cites the databook's own
    remarks column; a consultant states the fact, or attributes it to
    management ('管理层表示...'), which is an established convention here.
  * repeated 人民币 in one sentence.
  * tie-out ('Check') text leaking into the commentary -- a working
    artefact that should never appear in a deliverable.
  * a bullet that reproduces a remark almost verbatim rather than
    summarising it into report prose.

Usage:
    python inspect_databook_to_pptx.py --databook "for_test/x.xlsx" --pptx "for_test/x.pptx"
    python inspect_databook_to_pptx.py --databook ... --pptx ... --account 管理费用
"""
import argparse
import difflib
import re
import sys
import warnings

warnings.filterwarnings("ignore")
sys.path.insert(0, ".")

from pptx import Presentation

from fdd_utils.workbook import process_workbook_data
from inspect_is_variance import (
    _entity_from_filename, _substantive_notes, _substantive_linked, _note_prose,
)

_SOURCE_META_RE = re.compile(
    r"(根据|依据|按照)\s*(补充)?备注(信息|内容|说明)?|备注(显示|表明|提到|中提及)|"
    r"according\s+to\s+the\s+(supplementary\s+)?(remarks?|notes?)\b|"
    r"(as|per)\s+(stated|noted)\s+in\s+the\s+(supplementary\s+)?remarks?\b",
    re.IGNORECASE,
)
_TIEOUT_LEAK_RE = re.compile(r"\bcheck\b|对账单|對賬單|差异数|recon(ciliation)?\s+row", re.IGNORECASE)
_CHI_SENTENCE_SPLIT_RE = re.compile(r"[。；;]")


def extract_bullets(pptx_path: str):
    """{account_label: full bullet text} for every '■ <account> - <text>'
    bullet in the deck. A long account can be split across slots with a
    '(cont'd)' marker, so fragments are stitched back together under one
    label -- otherwise a continuation looks like a separate, truncated
    account and its remark uptake is judged against half the text."""
    prs = Presentation(pptx_path)
    bullets = {}
    order = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for line in shape.text_frame.text.split("\n"):
                s = line.strip()
                if not s.startswith("■"):
                    continue
                body = s.lstrip("■").strip()
                if " - " in body:
                    label, text = body.split(" - ", 1)
                else:
                    label, text = body, ""
                label = label.strip()
                label = re.sub(r"\s*\((cont'd|续|續)\)\s*$", "", label).strip()
                if label in bullets:
                    bullets[label] += " " + text.strip()
                else:
                    bullets[label] = text.strip()
                    order.append(label)
    return bullets, order


def _distinctive_tokens(prose: str):
    """CJK runs and long latin words from a remark -- used to ask whether
    that remark visibly reached the commentary. Short tokens are dropped
    because generic words ('费用', 'the') match anything."""
    toks = re.findall(r"[一-鿿]{2,}|[A-Za-z]{4,}", prose or "")
    return [t for t in toks if len(t) >= 2]


def _remark_reached_output(prose: str, output: str) -> bool:
    toks = _distinctive_tokens(prose)
    if not toks:
        return False
    hits = sum(1 for t in toks if t in output)
    return hits >= max(1, len(toks) // 3)


def _style_flags(text: str):
    """[(kind, detail)] -- kind is a stable short label so the summary can
    group by issue type rather than by the varying quoted text."""
    flags = []
    for m in _SOURCE_META_RE.finditer(text):
        flags.append(("source meta-reference",
                      f"{m.group(0)!r} -- state the fact directly"))
    for sentence in _CHI_SENTENCE_SPLIT_RE.split(text):
        n = sentence.count("人民币")
        if n >= 2:
            flags.append(("repeated 人民币 in one sentence",
                          f"x{n}: {sentence.strip()[:90]!r}"))
    for m in _TIEOUT_LEAK_RE.finditer(text):
        flags.append(("tie-out artefact leaked into the deliverable", repr(m.group(0))))
    return flags


def _verbatim_ratio(remark_prose: str, output: str) -> float:
    """Longest common run between a remark and the commentary, as a share
    of the remark. High values mean the note was pasted rather than
    rewritten into report prose."""
    if not remark_prose or not output:
        return 0.0
    sm = difflib.SequenceMatcher(None, remark_prose, output)
    match = sm.find_longest_match(0, len(remark_prose), 0, len(output))
    return match.size / max(1, len(remark_prose))


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--databook", required=True, help="path to the source databook .xlsx")
    ap.add_argument("--pptx", required=True, help="path to the exported deck .pptx")
    ap.add_argument("--entity", default=None, help="entity name (default: derived from the filename)")
    ap.add_argument("--sheet", default=None, help="specific sheet, if needed")
    ap.add_argument("--account", default=None, help="only report this one account")
    ap.add_argument("--verbatim-threshold", type=float, default=0.5,
                     help="share of a remark appearing as one unbroken run in the output above "
                          "which it's called near-verbatim (default 0.5)")
    args = ap.parse_args()

    entity = args.entity or _entity_from_filename(args.databook)
    print(f"Databook : {args.databook!r} (entity={entity!r})")
    print(f"Deck     : {args.pptx!r}\n")

    result = process_workbook_data(temp_path=args.databook, entity_name=entity,
                                    selected_sheet=args.sheet)
    dfs = result["dfs"]
    print(f"{len(dfs)} account(s) in the databook. Language: {result.get('language')}")

    bullets, order = extract_bullets(args.pptx)
    print(f"{len(bullets)} account bullet(s) in the deck: {', '.join(order[:12])}"
          f"{' ...' if len(order) > 12 else ''}\n")

    # Deck labels are display names; databook keys are mapping keys. Match
    # exactly first, then by containment either way, before giving up --
    # a label like '应付账款' can appear in the deck while the key is 'AP'.
    def find_bullet(key: str, df):
        display = str(df.attrs.get("display_key") or "").strip()
        for cand in (key, display):
            if cand and cand in bullets:
                return bullets[cand], cand
        for label, text in bullets.items():
            if key and (key in label or label in key):
                return text, label
            if display and (display in label or label in display):
                return text, label
        return None, None

    total, matched, with_remarks, used_remarks = 0, 0, 0, 0
    all_flags, unused, verbatim_hits = [], [], []

    for key in sorted(dfs):
        if args.account and key != args.account:
            continue
        df = dfs[key]
        notes = _substantive_notes(df.attrs.get("supporting_notes") or [])
        linked = _substantive_linked(df.attrs.get("table_linked_remarks") or [])
        rhs = df.attrs.get("adjacent_detail_rows") or []
        total += 1

        output, label = find_bullet(key, df)
        print("=" * 78)
        print(f"{key}" + (f"   [deck label: {label!r}]" if label and label != key else ""))
        print("=" * 78)
        if output is None:
            print("  ⚠️  no matching bullet found in the deck (account not written up, or the")
            print("      deck label differs too much to match automatically)\n")
            continue
        matched += 1

        print(f"  SOURCE remarks ({len(notes)} note(s), {len(rhs)} RHS row(s), {len(linked)} linked):")
        if not notes and not rhs and not linked:
            print("    (none substantive -- commentary can only restate the figures)")
        for n in notes[:6]:
            print(f"    · {str(n)[:200]}")
        for r in rhs[:4]:
            desc = r.get("Description") if isinstance(r, dict) else None
            extra = [f"{k}={str(v)[:60]}" for k, v in (r.items() if isinstance(r, dict) else [])
                     if "Detail" in str(k) and str(v).strip()]
            if desc or extra:
                print(f"    · [RHS] {desc}: {'; '.join(extra[:2])}"[:200])

        print(f"\n  GENERATED commentary ({len(output)} chars):")
        for i in range(0, min(len(output), 900), 110):
            print(f"    {output[i:i + 110]}")
        if len(output) > 900:
            print(f"    ... (+{len(output) - 900} more chars)")

        if notes or rhs:
            with_remarks += 1
            reached, missed = [], []
            for n in notes:
                prose = _note_prose(n)
                (reached if _remark_reached_output(prose, output) else missed).append(prose[:70])
            if reached:
                used_remarks += 1
            print(f"\n  remark uptake: {len(reached)}/{len(notes)} note(s) visibly reached the output")
            for m in missed[:4]:
                print(f"    ✗ not visible: {m!r}")
                unused.append((key, m))
            for n in notes:
                prose = _note_prose(n)
                ratio = _verbatim_ratio(prose, output)
                if ratio >= args.verbatim_threshold and len(prose) >= 12:
                    print(f"    ⚠️ near-verbatim ({ratio:.0%} of the note is one unbroken run in the "
                          f"output) -- summarise rather than paste: {prose[:60]!r}")
                    verbatim_hits.append((key, prose[:60], ratio))

        flags = _style_flags(output)
        if flags:
            print(f"\n  STYLE FLAGS ({len(flags)}):")
            for kind, detail in flags:
                print(f"    ⚠️ {kind}: {detail}")
            all_flags.extend((key, kind) for kind, _detail in flags)
        print()

    print("=" * 78)
    print("SUMMARY")
    print("=" * 78)
    print(f"  accounts examined                 : {total}")
    print(f"  matched to a deck bullet          : {matched}")
    print(f"  had substantive source material   : {with_remarks}")
    print(f"  ... of which visibly used any     : {used_remarks}")
    if with_remarks:
        print(f"  => remark uptake rate: {used_remarks / with_remarks * 100:.0f}%")
    print(f"  style flags across the deck       : {len(all_flags)}")
    if all_flags:
        by_kind = {}
        for _key, kind in all_flags:
            by_kind[kind] = by_kind.get(kind, 0) + 1
        for kind, n in sorted(by_kind.items(), key=lambda kv: -kv[1]):
            print(f"    {n:>3d} x  {kind}")
    if verbatim_hits:
        print(f"\n  near-verbatim remark reuse: {len(verbatim_hits)}")
        for key, prose, ratio in verbatim_hits[:8]:
            print(f"    {ratio:.0%}  {key}: {prose!r}")
    if unused:
        print(f"\n  source remarks that did NOT visibly reach the deck: {len(unused)}")
        for key, m in unused[:10]:
            print(f"    {key}: {m!r}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
