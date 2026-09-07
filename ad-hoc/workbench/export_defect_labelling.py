#!/usr/bin/env python3
"""M8.6 — label DELIVERABLE defects, not flagged ones. Export sentences, summarise reasons.

Every code the plan proposes for `DEFECT_CODES` is named after a detector branch
that already exists in validator.py. That is a taxonomy of what is CHECKABLE, and
it was derived by reading the checker — so by construction it cannot contain a
defect nobody wrote a check for. Freezing a frozenset built that way locks the
blind spot in.

One such blind spot, measured over 247 archived run folders on 2026-09-07: of
4,219 archived final texts, **618 (14.6%) contain a percentage — 892 percentage
tokens in all — and not one is groundable**, because bare percentages are
deliberately not extracted as amounts. A sentence saying occupancy rose "from
0.9% in 2024 to 52.1%" passes verification untouched whatever the numbers are.

Two more, found while building this tool, neither with a code and neither
detectable by anything in the pipeline:

    "Financial expenses mainly comprised interest支出 of CNY7.6 million …"
    "… incurred during 1M26, which were accounted for on an accrual basis.
     5 million in FY24, and CNY14.8 million in FY25."

The first is a Chinese token welded into an English sentence; the second is a
sentence that lost its head somewhere in the pipeline. Both shipped. Both are
"supported" — every figure in them checks out.

So this tool asks the only question that finds those: give a reviewer real
archived text and have them mark **every sentence they would not send to a
client**, with a free-text reason in their own words. Then `--summarise` sorts
the reasons into three piles — covered by a code that exists, covered by a
candidate code this tool proposes from the evidence above, and covered by
nothing, printed verbatim for a human to name. The third pile is the output that
matters; it is what M1's frozenset is missing.

WHAT THE REVIEWER FILLS IN — two columns, per sentence:

    send_to_client   y = I would send this as it stands
                     n = I would not
                     ?  = cannot tell without the workbook open
    reason           free text, ONLY on the `n` rows, in your own words.
                     Do not try to use a code name. --summarise does the sorting,
                     and a reviewer reaching for the existing vocabulary is
                     exactly how a taxonomy stays blind.

The companion .md is the reading copy: each account's full text with its
sentences numbered to match the CSV, and the source figures underneath. Read
there, mark in the CSV.

WHAT IS SAMPLED
~10 final texts per statement type (BS / IS / untyped), each half from texts
carrying a percentage and half without, so the known blind spot is represented
without being the whole sample and `--summarise` can report a rate for each half.
DEAD runs are excluded: a run whose every LLM call returned 400 ships
deterministic fallback bullets as its commentary (five such runs are in this
archive — see list_archived_runs.py), and grading those teaches nothing about
what the model writes.

Only one account, `S&D expenses`, has no `type` in mappings.yml — that is a
mappings gap rather than a third real statement type, and it is why the untyped
stratum is small.

Usage:
    # export (free, ~10s, no tokens)
    PYTHONPATH=. python ad-hoc/workbench/export_defect_labelling.py \\
        "<your-databook>.xlsx" --out /tmp/defects.csv

    # ... reviewer reads /tmp/defects.csv.reading.md, fills /tmp/defects.csv ...

    # summarise (free, instant, needs no databook)
    PYTHONPATH=. python ad-hoc/workbench/export_defect_labelling.py \\
        --summarise /tmp/defects.csv

The CSV is written with a UTF-8 BOM so Excel opens the Chinese text correctly.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import random
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import pandas as pd
import yaml

REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import list_archived_runs as archive  # noqa: E402  (same folder; health verdicts)
from export_calibration_sample import (  # noqa: E402  (same folder; shared plumbing)
    build_dfs,
    paired_runs,
    resolve_runs,
    source_digest,
    _norm,
)

LOGS_DIR = REPO_ROOT / "fdd_utils" / "logs"

_PCT = re.compile(r"\d+(?:\.\d+)?\s*%|百分之")

# --------------------------------------------------------------------------
# the taxonomy under test
# --------------------------------------------------------------------------

# M1's frozenset landed at validator.py:40 while this tool was being written, so
# the import below is the live path and the literal is now only a guard for a
# checkout where M1 has been reverted. Keep both: grading a taxonomy against a
# copy of itself is the failure mode this whole milestone exists to avoid, and
# CODES_SOURCE prints which one was used so a stale literal cannot masquerade as
# the real list.
try:  # pragma: no cover - depends on whether M1 is present in this checkout
    from fdd_utils.ai.validator import DEFECT_CODES as EXISTING_CODES  # type: ignore
    CODES_SOURCE = "fdd_utils.ai.validator.DEFECT_CODES"
except ImportError:
    EXISTING_CODES = frozenset({
        "AMOUNT_UNSUPPORTED", "AMOUNT_SCALE_ERROR", "DATE_UNSUPPORTED",
        "DIRECTION_MISMATCH", "COMPOSITION_GAP", "COMPOSITION_UNIT_ERROR",
        "COMPOSITION_DOUBLE_COUNT", "LLM_UNSUPPORTED_CAUSE", "LLM_UNSUPPORTED_FACT",
        "UNCATEGORIZED", "CLAIM_MISSING",
    })
    CODES_SOURCE = "the plan's proposed M1 list (validator.DEFECT_CODES not yet in code)"

# Keyword routes from a reviewer's own words to a code that already exists.
# Bilingual because roughly half the archived commentary is Chinese and the
# reviewer will write the reason in whichever language the text was in.
EXISTING_ROUTES: Dict[str, Tuple[str, ...]] = {
    "AMOUNT_UNSUPPORTED": ("wrong number", "wrong figure", "wrong amount", "not in the source",
                           "not in source", "no such figure", "invented number", "made up",
                           "數字錯", "数字错", "金额错", "金額錯", "對不上", "对不上", "查無", "查无"),
    "AMOUNT_SCALE_ERROR": ("scale", "magnitude", "wrong unit", "off by", "10x", "100x", "1000x",
                           "decimal", "万元", "萬元", "亿", "億", "千元", "單位錯", "单位错", "量級", "量级"),
    "DATE_UNSUPPORTED": ("wrong date", "wrong period", "wrong year", "date is", "日期錯", "日期错",
                         "期間錯", "期间错", "年份"),
    "DIRECTION_MISMATCH": ("wrong direction", "opposite", "increase but", "decrease but",
                           "should be a decrease", "should be an increase", "方向", "升跌", "反了", "反咗"),
    "COMPOSITION_GAP": ("does not add", "doesn't add", "not add up", "components missing",
                        "missing component", "加唔埋", "合計不符", "合计不符", "明細不全", "明细不全"),
    "COMPOSITION_DOUBLE_COUNT": ("double count", "counted twice", "重複計", "重复计", "重覆"),
    "LLM_UNSUPPORTED_CAUSE": ("speculat", "guess", "no evidence for the cause", "invented reason",
                              "unsupported driver", "attributed", "推測", "推测", "臆測", "臆测",
                              "無根據", "无根据", "原因無"),
    "LLM_UNSUPPORTED_FACT": ("invented", "fabricat", "not stated anywhere", "management never",
                             "捏造", "杜撰", "未提及", "沒有講", "没有讲"),
}

# Candidate codes the taxonomy has NO branch for. Every one of these was seen in
# the archive before it was written down here — they are not speculation, and
# each carries the observation that produced it. --summarise reports them
# separately from the existing codes so the gap is legible rather than absorbed.
#
# Both route tables are a FIRST PASS and are expected to leave a large unrouted
# pile on the first real labelling round. That pile is the deliverable, not a
# defect in these tables: a matcher tuned until everything routes would have
# deleted exactly the finding the milestone is after. Extend them after reading
# the unrouted reasons — by adding phrasings a real reviewer used, never by
# broadening a keyword until it swallows the residue.
CANDIDATE_ROUTES: Dict[str, Tuple[str, ...]] = {
    # 892 percentage tokens across 618 finals, none groundable.
    "PERCENTAGE_UNGROUNDED": ("percent", "%", "百分", "佔比", "占比", "比率", "ratio", "occupancy rate"),
    # "interest支出", seen shipped in an English deck.
    "LANGUAGE_MIX": ("chinese word", "chinese character", "english word", "mixed language",
                     "mixed script", "untranslated", "not translated", "中英", "夾雜",
                     "夹杂", "冇譯", "没翻译", "未翻譯"),
    # "… on an accrual basis. 5 million in FY24, and CNY14.8 million in FY25."
    "SENTENCE_TRUNCATED": ("fragment", "incomplete sentence", "cut off", "no subject",
                           "does not parse", "garbled", "斷句", "断句", "唔完整", "不完整", "殘句"),
    "STYLE_REGISTER": ("style", "tone", "register", "wording", "phrasing", "awkward", "clumsy",
                       "reads badly", "rewrite", "hard to read", "文風", "文风", "用詞", "用词",
                       "語氣", "语气", "唔順", "不顺", "難讀", "难读"),
    "REDUNDANT": ("repeat", "redundan", "duplicate", "says it twice", "重複", "重复", "囉唆", "啰嗦"),
    "NO_ANALYSIS": ("no analysis", "just restates", "restate", "says what not why", "no insight",
                    "no explanation", "冇分析", "没分析", "只講數", "只讲数", "無解釋", "无解释"),
    "ENTITY_NAME": ("company name", "entity name", "full legal name", "not shortened",
                    "公司名", "全稱", "全称", "簡稱", "简称"),
    "CLIENT_CONFIDENTIAL": ("internal", "should not go to the client", "confidential",
                            "內部", "内部", "唔應該畀客", "不该给客户"),
}


# --------------------------------------------------------------------------
# sentence splitting
# --------------------------------------------------------------------------

# Abbreviations whose "." is not a sentence end. "Co., Ltd." is the one that
# matters — `segment_clauses` in validator.py breaks on it, which is where the
# one-token clause fragments in the calibration frame come from, and repeating
# that mistake here would hand the reviewer a row saying "Ltd." to judge.
_ABBREV = (
    "Co.", "Ltd.", "Inc.", "Corp.", "Pte.", "Sdn.", "Bhd.", "Pty.", "plc.",
    "No.", "Nos.", "approx.", "etc.", "e.g.", "i.e.", "vs.", "Mr.", "Ms.", "Mrs.",
    "Dr.", "St.", "Jan.", "Feb.", "Mar.", "Apr.", "Jun.", "Jul.", "Aug.",
    "Sept.", "Sep.", "Oct.", "Nov.", "Dec.",
)


def split_sentences(text: str) -> List[str]:
    """Coarse sentence split — deliberately coarser than `segment_clauses`.

    The verifier splits into clauses because it grounds figures, and a figure
    lives in a clause. A reviewer judging "would I send this" judges a sentence:
    "which were in line with the payment terms" is not something anyone can
    accept or reject on its own. So this splits on sentence punctuation only,
    protects the abbreviation dots and decimal points that would otherwise shatter
    a company name or an amount, and keeps a `➢` bullet or a hard line break as
    its own unit — the two-part subtable bullet convention (prompts.py:852-866)
    puts a load-bearing newline in the text and it is a real boundary.
    """
    if not text:
        return []
    work = text
    # Mask the dots that are not sentence ends, restore after splitting.
    for i, abbrev in enumerate(_ABBREV):
        work = work.replace(abbrev, abbrev[:-1] + f"\x00{i}\x00")
    work = re.sub(r"(?<=\d)\.(?=\d)", "\x01", work)

    pieces: List[str] = []
    for line in re.split(r"[\r\n]+", work):
        line = line.strip()
        if not line:
            continue
        for piece in re.split(r"(?<=[。！？；])|(?<=[.;!?])(?=\s)", line):
            piece = piece.strip()
            if piece:
                pieces.append(piece)

    out: List[str] = []
    for piece in pieces:
        piece = piece.replace("\x01", ".")
        for i, abbrev in enumerate(_ABBREV):
            piece = piece.replace(abbrev[:-1] + f"\x00{i}\x00", abbrev)
        # A trailing orphan of punctuation is not a sentence; glue it back.
        if out and len(piece) <= 2 and not re.search(r"[A-Za-z一-鿿0-9]", piece):
            out[-1] = out[-1] + piece
        else:
            out.append(piece)
    return out


# --------------------------------------------------------------------------
# sampling the texts
# --------------------------------------------------------------------------

def account_final_text(result: Dict[str, Any]) -> str:
    """What the deck shipped. `final` first, because that is the string the deck
    renders; the calibration tool prefers agent_4_validation.final_content instead
    because clause spans are defined against THAT string and the two differ on
    some accounts (the Validator may keep an older `final` deliberately,
    pipeline.py:1203-1206). Different questions, different string — this one asks
    what the client saw."""
    for key in ("final", "subagent_4", "subagent_2", "subagent_1"):
        value = result.get(key)
        if isinstance(value, str) and value.strip():
            return _norm(value)
    return ""


def collect_texts(run_dirs: Sequence[Path], skip_dead: bool = True) -> Tuple[List[Dict[str, Any]], Counter]:
    from fdd_utils.ai import get_prompt_engine

    prompt_manager = get_prompt_engine()
    seen: Dict[str, Dict[str, Any]] = {}
    totals: Counter = Counter()
    for run_dir in run_dirs:
        results_path = run_dir / "results.yml"
        if not results_path.exists():
            continue
        info = archive.scan_log(run_dir / "processing.log")
        if skip_dead and info.get("has_log") and info.get("successful_calls", 0) == 0:
            totals["skipped_dead_runs"] += 1
            continue
        try:
            data = yaml.safe_load(results_path.read_text(encoding="utf-8")) or {}
        except Exception:
            continue
        for account, result in data.items():
            if not isinstance(result, dict):
                continue
            text = account_final_text(result)
            if len(text) < 60:
                totals["skipped_too_short"] += 1
                continue
            totals["texts_seen"] += 1
            statement = prompt_manager.get_mapping_component(account, component="type") or "untyped"
            key = hashlib.sha1(f"{account}\x1f{text}".encode("utf-8")).hexdigest()[:12]
            if key in seen:
                seen[key]["occurrences"] += 1
                continue
            has_pct = bool(_PCT.search(text))
            totals[f"pop_{statement}_{'pct' if has_pct else 'nopct'}"] += 1
            seen[key] = {
                "text_id": key,
                "run": run_dir.name,
                "account": account,
                "statement": statement,
                "language": info.get("language", "?"),
                "has_pct": has_pct,
                "text": text,
                "occurrences": 1,
            }
    return list(seen.values()), totals


def sample_texts(texts: List[Dict[str, Any]], per_type: int, seed: int) -> List[Dict[str, Any]]:
    """Half with a percentage, half without, per statement type, account-spread.

    The percentage half is not there to make the sample representative — 14.6% of
    finals carry one, so a representative sample would have about one and a half
    per type. It is there because that is the measured blind spot this milestone
    is aimed at, and one and a half texts cannot show whether it produces
    defects. --summarise reports the two halves separately and prints the true
    14.6% base rate beside them, so nothing downstream mistakes the oversample
    for a population rate.
    """
    rng = random.Random(seed)
    chosen: List[Dict[str, Any]] = []
    by_bucket: Dict[Tuple[str, bool], Dict[str, List[Dict[str, Any]]]] = defaultdict(lambda: defaultdict(list))
    for entry in texts:
        by_bucket[(entry["statement"], entry["has_pct"])][entry["account"]].append(entry)

    statements = sorted({e["statement"] for e in texts})
    for statement in statements:
        for has_pct, want in ((True, per_type // 2), (False, per_type - per_type // 2)):
            buckets = by_bucket.get((statement, has_pct), {})
            for rows in buckets.values():
                rng.shuffle(rows)
            accounts = sorted(buckets)
            rng.shuffle(accounts)
            taken: List[Dict[str, Any]] = []
            cursor = 0
            while len(taken) < want and any(buckets[a] for a in accounts):
                account = accounts[cursor % len(accounts)]
                cursor += 1
                if buckets[account]:
                    taken.append(buckets[account].pop())
            chosen.extend(taken)
    chosen.sort(key=lambda e: (e["statement"], e["account"], e["text_id"]))
    return chosen


# --------------------------------------------------------------------------
# export
# --------------------------------------------------------------------------

COLUMNS = [
    "send_to_client", "reason",            # the two the reviewer touches, first
    "row_id", "text_id", "statement", "account", "run", "language",
    "sentence_no", "sentence", "has_pct",
]


def do_export(args: argparse.Namespace) -> None:
    run_dirs = resolve_runs(args.run)
    print(f"Scanning {len(run_dirs)} run folder(s) under {LOGS_DIR} …")

    dfs = build_dfs(args.databook, args.entity, args.sheet)
    if not dfs:
        sys.exit(f"❌ No dfs built from {args.databook} — wrong sheet? (pass --sheet)")
    kept, dropped = paired_runs(dfs, run_dirs)
    print(f"  {len(dfs)} accounts in {Path(args.databook).name}; "
          f"{len(kept)} run(s) pair with it, {dropped} dropped as another book's.")
    if not kept:
        sys.exit("❌ No archived run pairs with this databook. Nothing to sample.")

    texts, totals = collect_texts(kept)
    chosen = sample_texts(texts, args.per_type, args.seed)
    if not chosen:
        sys.exit("❌ No texts survived sampling.")

    rows: List[Dict[str, Any]] = []
    for entry in chosen:
        sentences = split_sentences(entry["text"])
        entry["sentences"] = sentences
        for i, sentence in enumerate(sentences, start=1):
            rows.append({
                "send_to_client": "",
                "reason": "",
                "row_id": f"{entry['text_id']}-{i:02d}",
                "text_id": entry["text_id"],
                "statement": entry["statement"],
                "account": entry["account"],
                "run": entry["run"],
                "language": entry["language"],
                "sentence_no": i,
                "sentence": sentence,
                "has_pct": "y" if _PCT.search(sentence) else "",
            })

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    with open(out, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=COLUMNS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)

    reading = write_reading_copy(out, chosen, dfs)

    # The population base rates --summarise needs, for the same reason the
    # calibration tool ships a sidecar: they cannot live in a column the reviewer
    # would have to ignore on every row.
    pct_pop = sum(v for k, v in totals.items() if k.startswith("pop_") and k.endswith("_pct"))
    nopct_pop = sum(v for k, v in totals.items() if k.startswith("pop_") and k.endswith("_nopct"))
    meta = {
        "databook": Path(args.databook).name,
        "runs_paired": len(kept),
        "runs_dropped": dropped,
        "dead_runs_skipped": totals["skipped_dead_runs"],
        "seed": args.seed,
        "per_type": args.per_type,
        "texts_seen": totals["texts_seen"],
        "distinct_texts": len(texts),
        "population_with_pct": pct_pop,
        "population_without_pct": nopct_pop,
        "sampled_texts": len(chosen),
        "sampled_sentences": len(rows),
        "codes_source": CODES_SOURCE,
    }
    out.with_suffix(out.suffix + ".frame.yml").write_text(
        yaml.safe_dump(meta, allow_unicode=True, sort_keys=True), encoding="utf-8")

    print("\n--- SAMPLE ---")
    by_statement = Counter(e["statement"] for e in chosen)
    for statement in sorted(by_statement):
        with_pct = sum(1 for e in chosen if e["statement"] == statement and e["has_pct"])
        sents = sum(len(e["sentences"]) for e in chosen if e["statement"] == statement)
        print(f"  {statement:<10} {by_statement[statement]:>3} texts "
              f"({with_pct} with a percentage), {sents:>4} sentences")
    print(f"  {'TOTAL':<10} {len(chosen):>3} texts, {len(rows):>4} sentences")
    if pct_pop + nopct_pop:
        print(f"  population base rate for percentages: "
              f"{pct_pop}/{pct_pop + nopct_pop} = {pct_pop / (pct_pop + nopct_pop):.1%} "
              f"(the sample deliberately oversamples them)")

    minutes, parts = estimate_minutes(chosen, rows)
    print("\n--- REVIEWER TIME (estimated from the actual content) ---")
    for name, value in parts.items():
        print(f"  {name:<34} {value:6.1f} min")
    print(f"  {'TOTAL':<34} {minutes:6.1f} min")
    if minutes > 95:
        print(f"  ⚠️  over the 90-minute budget — re-run with --per-type "
              f"{max(2, int(args.per_type * 90 / minutes))}")
    elif minutes < 65:
        # The default of 10 per type is the plan's own figure, and the texts turn
        # out shorter than the budget assumed (2.8 sentences each at the median).
        # Say so rather than quietly under-spending an hour that was already set
        # aside: more texts is the cheapest way to widen a taxonomy search.
        print(f"  {90 - minutes:.0f} min of the 90-minute budget unused — "
              f"--per-type {int(args.per_type * 90 / max(minutes, 1))} would fill it")

    print(f"\n✅ wrote {len(rows)} sentence rows → {out}")
    print(f"   reading copy → {reading}")
    print(f"   frame sidecar → {out.with_suffix(out.suffix + '.frame.yml')}")
    print("\n   Read the .md; in the CSV fill send_to_client (y/n/?) on every row and")
    print("   reason (free text, your own words) on the n rows.")
    print(f"   Then: PYTHONPATH=. python {Path(__file__).relative_to(REPO_ROOT)} --summarise {out}")


def estimate_minutes(texts: List[Dict[str, Any]], rows: List[Dict[str, Any]]) -> Tuple[float, Dict[str, float]]:
    """Reviewer minutes, from the real content, with the constants exposed.

    Different shape from the calibration tool's: there the unit is a row, here it
    is a TEXT — the reviewer orients on the account's source card once and then
    rules on that account's sentences in a batch, so the orientation cost is paid
    per text rather than per sentence.

      * 20 seconds per text to read its source card and get the account in mind
      * text read at 14 chars/second (the same skim-and-judge rate)
      * 7 seconds per sentence to decide y/n
      * 25 seconds to type a reason on the sentences rejected. The rejection rate
        is unknown before the exercise — that is what the exercise measures — so
        this assumes 20%, and the figure moves if the real rate does.
    """
    per_text_orient, read_rate, per_sentence, per_reason, reject_rate = 20.0, 14.0, 7.0, 25.0, 0.20
    chars = sum(len(t["text"]) for t in texts)
    seconds = (
        per_text_orient * len(texts)
        + chars / read_rate
        + per_sentence * len(rows)
        + per_reason * reject_rate * len(rows)
    )
    return seconds / 60, {
        "orient on each account": per_text_orient * len(texts) / 60,
        "read the texts": chars / read_rate / 60,
        "decide y/n per sentence": per_sentence * len(rows) / 60,
        "type reasons (assumes 20% rejected)": per_reason * reject_rate * len(rows) / 60,
    }


def write_reading_copy(out: Path, chosen: List[Dict[str, Any]], dfs: Dict[str, Any]) -> Path:
    """The reading copy. A CSV is where you mark, not where you read: a whole
    account text pasted into every one of its sentence rows would be the biggest
    cell in the file repeated four times, and reading a text one spreadsheet row
    at a time destroys the coherence the reviewer is being asked to judge."""
    path = out.with_suffix(out.suffix + ".reading.md")
    lines = [
        "# Defect labelling — reading copy",
        "",
        "Read a text whole, then mark its sentences in the CSV by `row_id`.",
        "Ask of every sentence: **would I send this to a client as it stands?**",
        "If not, say why in your own words — do not reach for a code name.",
        "",
    ]
    for entry in chosen:
        lines += [
            f"## {entry['account']}  ·  {entry['statement']}  ·  {entry['language']}"
            f"  ·  `{entry['text_id']}`",
            f"<sub>{entry['run']}, seen in {entry['occurrences']} archived run(s)</sub>",
            "",
        ]
        for i, sentence in enumerate(entry.get("sentences") or [], start=1):
            lines.append(f"{i}. {sentence}")
        lines.append("")
        df = dfs.get(entry["account"])
        if df is not None:
            figures, notes = source_digest(df, max_rows_chars=10_000, max_notes_chars=4_000)
            if figures:
                lines += ["<details><summary>source figures</summary>", "", "```",
                          figures.replace(" | ", "\n"), "```", "</details>", ""]
            if notes:
                lines += ["<details><summary>supporting notes</summary>", "", "```",
                          notes.replace(" || ", "\n"), "```", "</details>", ""]
        lines.append("---")
        lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")
    return path


# --------------------------------------------------------------------------
# summarise
# --------------------------------------------------------------------------

def route_reason(reason: str) -> Tuple[Optional[str], Optional[str]]:
    """(existing_code, candidate_code) a reason's own words fall under.

    Keyword routing, not clustering — no tokens are spent here and none should
    be. The point of the exercise is the reasons that route to NEITHER; a
    cleverer matcher that forced every reason into a bucket would delete the
    finding. So the routes are deliberately literal and the residue is printed
    verbatim for a human to read and name.
    """
    text = (reason or "").lower()
    if not text.strip():
        return None, None
    existing = next((code for code, keys in EXISTING_ROUTES.items()
                     if any(k.lower() in text for k in keys)), None)
    candidate = next((code for code, keys in CANDIDATE_ROUTES.items()
                      if any(k.lower() in text for k in keys)), None)
    return existing, candidate


_MIXED_SCRIPT = re.compile(r"[A-Za-z][一-鿿]|[一-鿿][A-Za-z]")


def auto_signals(sentence: str) -> List[str]:
    """Deterministic markers on a rejected sentence, independent of the reason text.

    A reviewer who writes "reads badly" has told you nothing routable; the
    sentence itself may still be carrying an obvious signature. These three cost
    nothing and were all observed in the archive, so the coverage report has a
    second, reason-independent leg to stand on.
    """
    signals = []
    if _PCT.search(sentence):
        signals.append("contains a percentage (ungroundable by construction)")
    if _MIXED_SCRIPT.search(sentence):
        signals.append("mixed Latin/CJK inside a token")
    stripped = sentence.strip()
    if re.match(r"^[\d,.]+\s", stripped) or re.match(r"^(and|但|及|以及)\b", stripped, re.I):
        signals.append("opens mid-clause — likely a truncated sentence")
    return signals


def do_summarise(args: argparse.Namespace) -> None:
    rows: List[Dict[str, str]] = []
    meta: Dict[str, Any] = {}
    for p in args.summarise:
        path = Path(p)
        if not path.exists():
            sys.exit(f"❌ No such file: {path}")
        with open(path, encoding="utf-8-sig", newline="") as fh:
            rows.extend(dict(r) for r in csv.DictReader(fh))
        sidecar = path.with_suffix(path.suffix + ".frame.yml")
        if sidecar.exists():
            meta.update(yaml.safe_load(sidecar.read_text(encoding="utf-8")) or {})
    if not rows:
        sys.exit("❌ No rows read.")

    marks = Counter((r.get("send_to_client") or "").strip().lower() for r in rows)
    rejected = [r for r in rows if (r.get("send_to_client") or "").strip().lower() == "n"]
    judged = marks["y"] + marks["n"]
    print("--- COVERAGE ---")
    print(f"  {len(rows)} sentences, {judged} judged, {marks['n']} rejected, "
          f"{marks['?']} unclear, {marks['']} blank")
    if not judged:
        sys.exit("\n❌ Nothing marked yet — fill the send_to_client column first.")
    print(f"  reject rate {marks['n']}/{judged} = {marks['n'] / judged:.0%}")

    by_statement: Dict[str, Counter] = defaultdict(Counter)
    for r in rows:
        mark = (r.get("send_to_client") or "").strip().lower()
        if mark in ("y", "n"):
            by_statement[r.get("statement", "?")][mark] += 1
    print("\n  by statement type:")
    for statement in sorted(by_statement):
        c = by_statement[statement]
        total = c["y"] + c["n"]
        print(f"    {statement:<10} {c['n']}/{total} rejected = {c['n'] / total:.0%}")

    pct_rows = [r for r in rows if (r.get("has_pct") or "").strip() == "y"
                and (r.get("send_to_client") or "").strip().lower() in ("y", "n")]
    other_rows = [r for r in rows if not (r.get("has_pct") or "").strip()
                  and (r.get("send_to_client") or "").strip().lower() in ("y", "n")]
    if pct_rows:
        pn = sum(1 for r in pct_rows if r["send_to_client"].strip().lower() == "n")
        on = sum(1 for r in other_rows if r["send_to_client"].strip().lower() == "n")
        print(f"\n  THE MEASURED BLIND SPOT — sentences carrying a percentage:")
        print(f"    with a percentage:    {pn}/{len(pct_rows)} rejected = "
              f"{pn / len(pct_rows):.0%}")
        if other_rows:
            print(f"    without:              {on}/{len(other_rows)} rejected = "
                  f"{on / len(other_rows):.0%}")
        print("    Every percentage in the archive is unverified by construction, so a")
        print("    reject rate here that is no worse than the baseline is itself the")
        print("    finding: the gap is not costing defects and need not be closed first.")

    print(f"\n--- REASONS ROUTED (existing codes, from {CODES_SOURCE}) ---")
    existing_hits: Counter = Counter()
    candidate_hits: Counter = Counter()
    unrouted: List[Dict[str, str]] = []
    for r in rejected:
        existing, candidate = route_reason(r.get("reason", ""))
        if existing:
            existing_hits[existing] += 1
        if candidate:
            candidate_hits[candidate] += 1
        if not existing and not candidate:
            unrouted.append(r)

    if existing_hits:
        for code, n in existing_hits.most_common():
            print(f"  {code:<28} {n:>4}")
    else:
        print("  (none — no rejected sentence's reason matched a code that exists)")
    unused = sorted(EXISTING_CODES - set(existing_hits))
    if unused:
        print(f"\n  Codes that exist and NOTHING routed to: {', '.join(unused)}")
        print("  A code with no reviewer reason behind it is a detector branch with no")
        print("  deliverable defect behind it. That is what M8.6 was looking for.")

    print("\n--- REASONS THE TAXONOMY HAS NO CODE FOR (candidates) ---")
    if candidate_hits:
        for code, n in candidate_hits.most_common():
            print(f"  {code:<28} {n:>4}   ← not in {CODES_SOURCE}")
    else:
        print("  (none matched a candidate)")

    print(f"\n--- UNROUTED: {len(unrouted)} reason(s) matching nothing ---")
    print("  Read these. They are the part no keyword table anticipated, and naming")
    print("  them is the human half of this milestone — that is why they are printed")
    print("  verbatim instead of being forced into a bucket.")
    for r in unrouted[: args.show]:
        print(f"\n  [{r.get('row_id')}] {r.get('account')} · {r.get('statement')}")
        print(f"    sentence: {(r.get('sentence') or '')[:200]}")
        print(f"    reason:   {(r.get('reason') or '')[:300]}")
    if len(unrouted) > args.show:
        print(f"\n  … {len(unrouted) - args.show} more (raise --show)")

    print("\n--- AUTOMATIC SIGNALS ON REJECTED SENTENCES (no reason text needed) ---")
    signal_hits: Counter = Counter()
    for r in rejected:
        for signal in auto_signals(r.get("sentence", "")):
            signal_hits[signal] += 1
    if signal_hits:
        for signal, n in signal_hits.most_common():
            print(f"  {n:>4}  {signal}")
        print("\n  These are read off the sentence, not the reason, so they hold even")
        print("  where a reviewer's wording routed nowhere. A signal that fires often on")
        print("  rejected sentences and rarely elsewhere is a detector worth writing.")
    else:
        print("  (none fired)")

    if meta.get("population_with_pct") is not None:
        pop_pct = meta["population_with_pct"]
        pop_all = pop_pct + meta.get("population_without_pct", 0)
        if pop_all:
            print(f"\n  Reminder from the frame sidecar: percentages are {pop_pct}/{pop_all} = "
                  f"{pop_pct / pop_all:.1%} of the archived population, and this sample")
            print("  oversamples them on purpose. Do not read the sample's split as a rate.")


# --------------------------------------------------------------------------

def main() -> None:
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("databook", nargs="?",
                    help="path to the databook these runs were generated from "
                         "(required for export; not needed with --summarise)")
    ap.add_argument("--run", action="append", default=[],
                    help="run folder, name or timestamp; repeatable. Default: every "
                         "run under fdd_utils/logs that pairs with the databook.")
    ap.add_argument("--out", help="CSV to write (export mode)")
    ap.add_argument("--per-type", type=int, default=10,
                    help="final texts per statement type (default 10)")
    ap.add_argument("--seed", type=int, default=11, help="sampling seed (default 11)")
    ap.add_argument("--entity", default=None, help="entity name passed to process_workbook_data")
    ap.add_argument("--sheet", default=None, help="Financials sheet (default: auto-resolve)")
    ap.add_argument("--summarise", action="append", default=[],
                    help="filled CSV to summarise; repeatable")
    ap.add_argument("--show", type=int, default=25,
                    help="how many unrouted reasons to print in full (default 25)")
    args = ap.parse_args()

    if args.summarise:
        do_summarise(args)
        return
    if not args.databook or not args.out:
        ap.error("export mode needs a databook and --out (or use --summarise)")
    do_export(args)


if __name__ == "__main__":
    main()
