#!/usr/bin/env python3
"""M8.5 — calibrate the verifier against a human. Export a sample, score it back.

Nothing in this repo has ever compared a verifier verdict to a human one. There
is not even a field where a human verdict could be written: `grep -rn
'human_label|ground_truth|reviewer_verdict'` over the whole tree returns nothing.
So every downstream decision that leans on the unsupported flag — M3's promotion
step, M5 graduating a detector from measurement to enforcement, whether the
`hallucination` category can be trusted at all — is currently ungrounded.

The sample frame already exists and needs no tokens. Measured over 247 archived
run folders on 2026-09-07: **25,136 clause reviews, 418 unsupported** —
hallucination 222, reasoning 171, category `None` 25. Deduplicated on
(account, clause, category) that is 327 distinct flagged clauses and 11,534
distinct supported ones, of which 8,030 are substantive (a fragment like "," or
"Ltd." is a segmentation artefact of `_CLAUSE_END_CHARS`, not a claim, and
asking a human to rule on one wastes the budget this tool exists to protect).

WHAT THE REVIEWER FILLS IN — one column, `human_label`, four allowed values:

    fine            nothing wrong with this clause; I would send it to a client
    hallucination   states a fact or a figure that is not in the source
    reasoning       an inference that goes beyond the source, but nothing invented
    ?               cannot tell from what is on the row (counted, never guessed)

One column, because the verdict is derivable from it: on a row the run flagged,
`fine` is a false positive and anything else is a true positive; on a row the run
passed, `fine` is a true negative and anything else is a miss. That also gives
the per-category confusion matrix for free, without a second column to fill.
`human_note` is optional and only read by a human.

WHY BOTH FLAGGED AND SUPPORTED ROWS ARE SAMPLED
Precision alone is the cheap half and the misleading one: a verifier that flags
almost nothing scores beautifully on it. Recall needs rows the run passed, so
the sample carries three supported strata. Their population is 8,030 against 327
flagged, so they are sampled far below their weight and the score mode reweights
by population — it prints the weights it used, because a recall computed from
~30 supported rows at a low base rate is an indication with a wide interval, not
a measurement. The Wilson intervals are printed for exactly that reason. Read
them before quoting a number.

The supported strata are split by whether the clause carries a figure, and
percentages get their own stratum on purpose: bare percentages are deliberately
not extracted as amounts (M8.6's measured blind spot), so a percentage clause is
supported *by construction*, never by evidence. 332 of the 8,030 substantive
supported clauses contain one.

WHY THE DATABOOK IS A REQUIRED POSITIONAL ARGUMENT
An archived run folder records no workbook identity — no path, no entity, no
sheet hash (see replay_verification.py, same folder, same trap). This tool needs
the databook for a second reason beyond identity: a reason string like
"Amount(s) 121,000,000 not found in source data within tolerance" cannot be
judged without seeing the source figures, and a reviewer who has to open Excel
for it will not finish the sample in 75 minutes. So each row carries a compact
digest of its account's rows and supporting notes, built from the databook.
Runs whose account keys are not a subset of this databook's dfs are dropped from
the frame with a count, never silently mispaired. Measured coverage: the widest
book at the repo root pairs with 193 of the 255 run folders and 329 of the 418
unsupported reviews (79%). To cover the rest, run this once per databook and
pass every filled file to --score at once.

Usage:
    # export (free, ~10s, no tokens)
    PYTHONPATH=. python ad-hoc/workbench/export_calibration_sample.py \\
        "<your-databook>.xlsx" --out /tmp/calibration.csv

    # ... the reviewer fills the human_label column in Excel/Numbers ...

    # score (free, instant, needs no databook)
    PYTHONPATH=. python ad-hoc/workbench/export_calibration_sample.py \\
        --score /tmp/calibration.csv [--score /tmp/other-book.csv ...]

The CSV is written with a UTF-8 BOM so Excel opens the Chinese clauses correctly.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import math
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

LOGS_DIR = REPO_ROOT / "fdd_utils" / "logs"

# The four values the reviewer may write. "?" is kept as a first-class value
# rather than left blank so an unreadable row is DISTINGUISHABLE from a row the
# reviewer never reached — the two mean opposite things when you divide by n.
LABELS = ("fine", "hallucination", "reasoning", "?")

# Default sample sizes per stratum. They sum to 100, NOT the plan's 150.
#
# The plan paired "~150 clauses" with "~75 reviewer-minutes" as one estimate; on
# the real rows those two do not hold together. Measured on the exported sample:
# a row is 45 seconds — 250 characters of context plus a 149-character reason to
# read, a source digest to scan on the rows where a figure decides it, and the
# decide-and-type cost. 150 of those is 112 minutes, and the half-hour of
# overrun lands at the end, where a tiring reviewer's labels are worth least.
# 100 rows is 75 minutes and the number the tool defaults to. `--n 150` restores
# the plan's size for anyone who wants the tighter interval and has the time;
# estimate_minutes() prints the cost either way, so the trade is visible.
DEFAULT_STRATA = {
    "flagged_hallucination": 34,
    "flagged_reasoning": 27,
    "flagged_uncategorised": 10,
    "supported_pct": 13,
    "supported_numeric": 10,
    "supported_prose": 6,
}

_PCT = re.compile(r"\d+(?:\.\d+)?\s*%|百分之")
_DIGIT = re.compile(r"\d")


# --------------------------------------------------------------------------
# frame
# --------------------------------------------------------------------------

def _norm(text: Any) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip()


def _substantive(clause: str) -> bool:
    """Is this clause a claim, or a segmentation artefact?

    `segment_clauses` breaks on every "." including the one inside "Co., Ltd.",
    so the review list is full of one-token fragments that are trivially
    supported and carry no claim. They are 3,504 of the 11,534 distinct supported
    clauses. Excluding them shrinks the supported population this tool reweights
    against, which is why the exclusion is applied to the population count too,
    not only to the sample.
    """
    words = len(re.findall(r"[A-Za-z]+", clause)) + len(re.findall(r"[一-鿿]", clause))
    return len(clause) >= 25 and words >= 5


def stratum_of(supported: bool, category: Optional[str], clause: str) -> Optional[str]:
    if not supported:
        if category == "hallucination":
            return "flagged_hallucination"
        if category == "reasoning":
            return "flagged_reasoning"
        return "flagged_uncategorised"
    if not _substantive(clause):
        return None
    if _PCT.search(clause):
        return "supported_pct"
    if _DIGIT.search(clause):
        return "supported_numeric"
    return "supported_prose"


def row_id(account: str, clause: str, category: Optional[str]) -> str:
    raw = f"{account}\x1f{clause}\x1f{category}".encode("utf-8")
    return hashlib.sha1(raw).hexdigest()[:12]


def scan_frame(run_dirs: Sequence[Path]) -> Tuple[Dict[str, Dict[str, Any]], Counter]:
    """Deduplicate every clause review in the given runs into the sample frame.

    Keyed on (account, clause, category) so the same sentence re-flagged in 40
    repeat runs of one databook costs the reviewer one row, not forty. The
    occurrence count is kept because the population weights the score mode uses
    are per-occurrence — "418 of 25,136" is an occurrence statement.
    """
    frame: Dict[str, Dict[str, Any]] = {}
    totals: Counter = Counter()
    for run_dir in run_dirs:
        results_path = run_dir / "results.yml"
        if not results_path.exists():
            continue
        try:
            data = yaml.safe_load(results_path.read_text(encoding="utf-8")) or {}
        except Exception:
            totals["unreadable_results"] += 1
            continue
        for account, result in data.items():
            if not isinstance(result, dict):
                continue
            validation = result.get("agent_4_validation")
            if not isinstance(validation, dict):
                continue
            text = _norm(validation.get("final_content") or result.get("final") or "")
            for review in validation.get("clause_reviews") or []:
                if not isinstance(review, dict):
                    continue
                clause = _norm(review.get("clause"))
                if not clause:
                    continue
                supported = bool(review.get("supported"))
                category = review.get("category")
                totals["reviews"] += 1
                if not supported:
                    totals["unsupported"] += 1
                stratum = stratum_of(supported, category, clause)
                if stratum is None:
                    totals["excluded_fragment"] += 1
                    continue
                totals[f"pop_{stratum}"] += 1
                rid = row_id(account, clause, category)
                entry = frame.get(rid)
                if entry is None:
                    frame[rid] = {
                        "row_id": rid,
                        "stratum": stratum,
                        "account": account,
                        "run": run_dir.name,
                        "clause": clause,
                        "context": context_window(text, clause),
                        "stored_verdict": "supported" if supported else "unsupported",
                        "stored_category": category if category is not None else "",
                        "stored_reason": _norm(review.get("reason"))[:600],
                        "occurrences": 1,
                    }
                else:
                    entry["occurrences"] += 1
                    if not entry["context"]:
                        entry["context"] = context_window(text, clause)
    return frame, totals


def context_window(text: str, clause: str, radius: int = 90) -> str:
    """The clause's neighbourhood in the account's own final text.

    A clause is judged as part of a sentence — "which were in line with the
    payment terms" is unjudgeable alone — but a whole account text is 285 chars
    at the median and 663 at p90, and pasting all of it into every row is what
    turns a 30-second row into a 2-minute one. A fixed window either side is the
    compromise; ellipses mark that it is a window, so a reviewer knows when the
    sentence they need is off the edge.
    """
    if not text or not clause:
        return ""
    pos = text.find(clause)
    if pos < 0:
        head = clause[:40]
        pos = text.find(head) if head else -1
        if pos < 0:
            return ""
    start = max(0, pos - radius)
    end = min(len(text), pos + len(clause) + radius)
    out = text[start:end]
    if start > 0:
        out = "…" + out
    if end < len(text):
        out = out + "…"
    return out


# --------------------------------------------------------------------------
# source digest — what makes an amount reason judgeable in 30 seconds
# --------------------------------------------------------------------------

def source_digest(df: Any, max_rows_chars: int = 700, max_notes_chars: int = 700) -> Tuple[str, str]:
    """(figures, notes) for one account, both truncated hard.

    Deliberately NOT the grounding pool. `SourceIndex` also pools column totals,
    2-4-row window sums, annualized variants and sibling accounts — reproducing
    that here would hand the reviewer the verifier's own answer and turn a
    calibration into a tautology. What the reviewer gets is the account's own
    rows and the human notes attached to them, which is the evidence a consultant
    would check the sentence against. If a figure is absent from this digest and
    the verifier still passed it, that is a finding, not a bug in this function.
    """
    if df is None:
        return "", ""
    try:
        label_col = df.columns[0]
        period_cols = [
            c for c in df.columns
            if not str(c).endswith("_formatted")
            and str(c) not in (str(label_col), "__source_row_idx")
        ]
        parts: List[str] = []
        for _, row in df.iterrows():
            vals = "; ".join(
                f"{c}={row[c]:,.0f}" for c in period_cols
                if pd.notna(row[c]) and isinstance(row[c], (int, float))
            )
            parts.append(f"{row[label_col]}: {vals}" if vals else str(row[label_col]))
        figures = " | ".join(parts)
    except Exception:
        figures = ""
    notes_raw = df.attrs.get("supporting_notes") if hasattr(df, "attrs") else None
    if isinstance(notes_raw, (list, tuple)):
        notes = " || ".join(_norm(n) for n in notes_raw)
    else:
        notes = _norm(notes_raw)
    return _truncate(figures, max_rows_chars), _truncate(notes, max_notes_chars)


_NEEDS_FIGURES = re.compile(
    r"Amount\(s\)|Date\(s\)|not found in source|within tolerance|"
    r"composition|components sum|UNIT error|未在数据中|与数据不符|金额"
)


def needs_source_figures(reason: str, clause: str) -> bool:
    """Does this row need the numbers pasted in, or is its reason self-contained?

    Measured on the exported sample, source_figures is the single biggest cell
    (459 chars at the median) and on a `reasoning` row it is dead weight: "the
    characterization … is a logical inference" is judgeable from the reason
    alone, and a reviewer who scans a 460-character digest anyway is spending ten
    seconds per row on nothing. That is fifteen minutes over a 150-row sample.
    So the digest is attached only where a verdict actually turns on a figure —
    a deterministic amount/date/composition reason, or a clause carrying a number
    the reviewer would otherwise have to trust. The full card is in the
    companion .sources.md for every account in the sample either way, so nothing
    is unreachable; it is one file away instead of in the row.
    """
    return bool(_NEEDS_FIGURES.search(reason or "")) or bool(_DIGIT.search(clause or ""))


def _truncate(text: str, limit: int) -> str:
    text = _norm(text)
    return text if len(text) <= limit else text[:limit] + " …[truncated]"


def build_dfs(databook: str, entity: Optional[str], sheet: Optional[str]) -> Dict[str, Any]:
    """Same construction replay_verification.py uses — the detail_analysis variant
    that production hands the pipeline. Kept as a copy rather than imported so
    this tool does not drag in the replay module's LLM-adjacent imports."""
    from fdd_utils.workbook import process_workbook_data

    if sheet is None:
        import inspect_databook

        found = inspect_databook._resolve_financials_sheets(pd.ExcelFile(databook))
        sheet = found[0] if found else None
    state = process_workbook_data(
        temp_path=databook, entity_name=entity, selected_sheet=sheet, debug=False,
    )
    return state.get("dfs") or {}


def paired_runs(dfs: Dict[str, Any], run_dirs: Sequence[Path]) -> Tuple[List[Path], int]:
    """Runs whose account keys are a subset of this databook's dfs.

    The same necessary-but-not-sufficient test replay_verification.py makes, and
    the same warning applies: two workbooks built from one template share an
    account vocabulary, so a subset match is evidence of pairing, not proof. What
    it does reliably prevent is the silent case — grounding one project's
    commentary against another project's numbers and reporting a digest that
    means nothing.
    """
    keys = set(dfs)
    kept, dropped = [], 0
    for run_dir in run_dirs:
        results_path = run_dir / "results.yml"
        if not results_path.exists():
            dropped += 1
            continue
        try:
            data = yaml.safe_load(results_path.read_text(encoding="utf-8")) or {}
        except Exception:
            dropped += 1
            continue
        accounts = {k for k, v in data.items() if isinstance(v, dict)}
        if accounts and accounts <= keys:
            kept.append(run_dir)
        else:
            dropped += 1
    return kept, dropped


# --------------------------------------------------------------------------
# sampling
# --------------------------------------------------------------------------

def sample_frame(frame: Dict[str, Dict[str, Any]], sizes: Dict[str, int],
                 seed: int) -> List[Dict[str, Any]]:
    """Stratified sample, account-spread within each stratum.

    Plain random sampling inside a stratum clusters on whichever account produced
    the most flags — Financial expenses and Operating income between them own a
    large share of the hallucination stratum — and a reviewer who sees the same
    account twenty times in a row starts pattern-matching instead of judging. So
    each stratum is drawn round-robin over accounts, shuffled once at the end.
    """
    rng = random.Random(seed)
    by_stratum: Dict[str, Dict[str, List[Dict[str, Any]]]] = defaultdict(lambda: defaultdict(list))
    for entry in frame.values():
        by_stratum[entry["stratum"]][entry["account"]].append(entry)

    for buckets in by_stratum.values():
        for rows in buckets.values():
            rng.shuffle(rows)

    def draw(stratum: str, want: int) -> List[Dict[str, Any]]:
        buckets = by_stratum.get(stratum, {})
        accounts = sorted(buckets)
        rng.shuffle(accounts)
        taken: List[Dict[str, Any]] = []
        cursor = 0
        while len(taken) < want and any(buckets[a] for a in accounts):
            account = accounts[cursor % len(accounts)]
            cursor += 1
            if buckets[account]:
                taken.append(buckets[account].pop())
        return taken

    chosen: List[Dict[str, Any]] = []
    shortfall = {"flagged": 0, "supported": 0}
    for stratum, want in sizes.items():
        taken = draw(stratum, want)
        chosen.extend(taken)
        shortfall[stratum.split("_")[0]] += want - len(taken)

    # A stratum can be empty for a given databook without the frame being thin:
    # all 25 archive-wide `category: None` reviews come from two March runs of a
    # different book, so pairing against any other book leaves
    # flagged_uncategorised at zero. Letting that shrink the sample would spend
    # 15 fewer reviewer-minutes on a budget that was already agreed, so the
    # shortfall is redistributed inside its own family — flagged shortfalls top
    # up flagged strata, supported top up supported. Never across the two: moving
    # rows between families would silently change the precision/recall split the
    # weights are computed against.
    for family, missing in shortfall.items():
        candidates = [s for s in sizes if s.startswith(family)]
        while missing > 0:
            drawn_this_pass = 0
            for stratum in candidates:
                if missing <= 0:
                    break
                extra = draw(stratum, 1)
                if extra:
                    chosen.extend(extra)
                    missing -= 1
                    drawn_this_pass += 1
            if drawn_this_pass == 0:
                break  # every stratum in this family is exhausted

    rng.shuffle(chosen)
    return chosen


# --------------------------------------------------------------------------
# export
# --------------------------------------------------------------------------

COLUMNS = [
    "human_label", "human_note",           # the two the reviewer touches, first
    "row_id", "stratum", "account", "run", "occurrences",
    "clause", "context",
    "stored_verdict", "stored_category", "stored_reason",
    "source_figures", "source_notes",
]


def write_csv(path: Path, rows: List[Dict[str, Any]], meta: Dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=COLUMNS, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow({c: row.get(c, "") for c in COLUMNS})
    # The population weights the score mode needs cannot live in the CSV without
    # a column the reviewer would have to leave alone on every row, so they go in
    # a sidecar keyed by the CSV's name. --score works without it (unweighted,
    # and it says so), which is why this is a sidecar and not a hard dependency.
    sidecar = path.with_suffix(path.suffix + ".frame.yml")
    sidecar.write_text(yaml.safe_dump(meta, allow_unicode=True, sort_keys=True), encoding="utf-8")


def write_source_cards(path: Path, rows: List[Dict[str, Any]],
                       dfs: Dict[str, Any]) -> Optional[Path]:
    """Full, untruncated source cards for the accounts in the sample.

    The CSV cells are capped so the spreadsheet stays readable; this is where a
    reviewer goes when a truncated digest was not enough. One file, read-only,
    account-ordered — cheaper to consult than reopening the databook.
    """
    accounts = sorted({r["account"] for r in rows})
    if not accounts or not dfs:
        return None
    out = path.with_suffix(path.suffix + ".sources.md")
    lines = ["# Source cards for the calibration sample", ""]
    for account in accounts:
        df = dfs.get(account)
        if df is None:
            continue
        figures, notes = source_digest(df, max_rows_chars=10_000, max_notes_chars=10_000)
        lines += [f"## {account}", "", "```", figures.replace(" | ", "\n"), "```", ""]
        if notes:
            lines += ["Supporting notes:", "", "```", notes.replace(" || ", "\n"), "```", ""]
    out.write_text("\n".join(lines), encoding="utf-8")
    return out


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

    frame, totals = scan_frame(kept)
    sizes = scale_sizes(DEFAULT_STRATA, args.n)
    rows = sample_frame(frame, sizes, args.seed)

    for row in rows:
        if needs_source_figures(row["stored_reason"], row["clause"]):
            figures, notes = source_digest(dfs.get(row["account"]))
        else:
            figures, notes = "", ""
        row["source_figures"] = figures
        row["source_notes"] = notes
        row["human_label"] = ""
        row["human_note"] = ""

    meta = {
        "databook": Path(args.databook).name,
        "runs_paired": len(kept),
        "runs_dropped": dropped,
        "seed": args.seed,
        "reviews_seen": totals["reviews"],
        "unsupported_seen": totals["unsupported"],
        "excluded_fragments": totals["excluded_fragment"],
        "population_occurrences": {
            s: totals[f"pop_{s}"] for s in DEFAULT_STRATA
        },
        "population_distinct": {
            s: sum(1 for e in frame.values() if e["stratum"] == s) for s in DEFAULT_STRATA
        },
        "sampled": {s: sum(1 for r in rows if r["stratum"] == s) for s in DEFAULT_STRATA},
        "labels_allowed": list(LABELS),
    }

    out = Path(args.out)
    write_csv(out, rows, meta)
    cards = write_source_cards(out, rows, dfs)

    print(f"\n--- FRAME (occurrences / distinct clauses) ---")
    for stratum in DEFAULT_STRATA:
        print(f"  {stratum:<24} pop {meta['population_occurrences'][stratum]:>6} "
              f"/ {meta['population_distinct'][stratum]:>6} distinct "
              f"→ sampled {meta['sampled'][stratum]:>4}")
    print(f"  {'TOTAL':<24} pop {sum(meta['population_occurrences'].values()):>6} "
          f"/ {sum(meta['population_distinct'].values()):>6} distinct "
          f"→ sampled {len(rows):>4}")
    minutes, parts = estimate_minutes(rows)
    print(f"\n--- REVIEWER TIME (estimated from the actual row content) ---")
    for name, value in parts.items():
        print(f"  {name:<34} {value:6.1f} min")
    print(f"  {'TOTAL':<34} {minutes:6.1f} min   "
          f"({minutes * 60 / max(1, len(rows)):.0f}s per row)")
    if minutes > 80:
        print(f"  ⚠️  over the 75-minute budget — re-run with --n {int(len(rows) * 75 / minutes)}")

    print(f"\n✅ wrote {len(rows)} rows → {out}")
    print(f"   frame sidecar → {out.with_suffix(out.suffix + '.frame.yml')}")
    if cards:
        print(f"   source cards → {cards}")
    print(f"\n   Fill ONE column: human_label ∈ {{{', '.join(LABELS)}}}. "
          f"human_note is optional and free text.")
    print(f"   Then: PYTHONPATH=. python {Path(__file__).relative_to(REPO_ROOT)} --score {out}")


def estimate_minutes(rows: List[Dict[str, Any]]) -> Tuple[float, Dict[str, float]]:
    """Reviewer minutes, from the actual characters in the actual rows.

    This is an estimate with visible assumptions, not a measurement, and it is
    printed with its constants so they can be argued with:

      * context (which contains the clause) + reason are READ, at 14 chars/second. That is fast for
        financial prose but this is skim-and-judge, not proofreading, and the
        mix is roughly half Chinese where a character carries more.
      * source_figures + source_notes are SCANNED for one number, at 45
        chars/second — nobody reads a row digest end to end.
      * 8 seconds per row of fixed cost: read the label options, decide, type.

    If the total comes out over the budget, the lever is --n, not a smaller
    context window: a row too thin to judge produces a "?" and costs the same.
    """
    read_rate, scan_rate, fixed = 14.0, 45.0, 8.0
    # NOT clause + context + reason: the context column CONTAINS the clause, so
    # charging both double-counts the longest text on the row and inflated an
    # earlier version of this estimate by roughly a third. The clause column is a
    # pointer into the context, not a second thing to read.
    read = sum(max(len(r.get("context", "")), len(r.get("clause", ""))) + len(r.get("stored_reason", ""))
               for r in rows)
    scan = sum(len(r.get("source_figures", "")) + len(r.get("source_notes", "")) for r in rows)
    seconds = read / read_rate + scan / scan_rate + fixed * len(rows)
    parts = {
        "reading (clause+context+reason)": read / read_rate / 60,
        "scanning (source digest)": scan / scan_rate / 60,
        "decide + type": fixed * len(rows) / 60,
    }
    return seconds / 60, parts


def scale_sizes(base: Dict[str, int], n: Optional[int]) -> Dict[str, int]:
    if not n:
        return dict(base)
    total = sum(base.values())
    return {k: max(1, round(v * n / total)) for k, v in base.items()}


def resolve_runs(run_args: Sequence[str]) -> List[Path]:
    if not run_args or list(run_args) == ["all"]:
        return sorted(p for p in LOGS_DIR.glob("run_*") if p.is_dir())
    out: List[Path] = []
    for arg in run_args:
        for candidate in (Path(arg), LOGS_DIR / arg, LOGS_DIR / f"run_{arg}"):
            if candidate.is_dir():
                out.append(candidate)
                break
        else:
            sys.exit(f"❌ No such run folder: {arg!r} (looked in {LOGS_DIR})")
    return out


# --------------------------------------------------------------------------
# scoring
# --------------------------------------------------------------------------

def wilson(hits: int, n: int, z: float = 1.96) -> Tuple[float, float]:
    """95% Wilson interval. Printed on every rate because the supported strata
    are small and a bare point estimate from 20 rows invites being quoted as if
    it were measured."""
    if n == 0:
        return (0.0, 1.0)
    p = hits / n
    denom = 1 + z * z / n
    centre = (p + z * z / (2 * n)) / denom
    half = z * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / denom
    return (max(0.0, centre - half), min(1.0, centre + half))


def read_scored(paths: Sequence[str]) -> Tuple[List[Dict[str, str]], Dict[str, Dict[str, int]]]:
    rows: List[Dict[str, str]] = []
    weights: Dict[str, Dict[str, int]] = {}
    for p in paths:
        path = Path(p)
        if not path.exists():
            sys.exit(f"❌ No such file: {path}")
        with open(path, encoding="utf-8-sig", newline="") as fh:
            rows.extend(dict(r) for r in csv.DictReader(fh))
        sidecar = path.with_suffix(path.suffix + ".frame.yml")
        if sidecar.exists():
            meta = yaml.safe_load(sidecar.read_text(encoding="utf-8")) or {}
            for stratum, pop in (meta.get("population_occurrences") or {}).items():
                weights.setdefault(stratum, {"pop": 0})["pop"] += int(pop)
    return rows, weights


def do_score(args: argparse.Namespace) -> None:
    rows, weights = read_scored(args.score)
    if not rows:
        sys.exit("❌ No rows read.")

    bad = sorted({(r.get("human_label") or "").strip() for r in rows} - set(LABELS) - {""})
    if bad:
        print(f"⚠️  Unrecognised human_label value(s), treated as unlabelled: {bad}")
        print(f"    Allowed: {', '.join(LABELS)}\n")

    per: Dict[str, Dict[str, int]] = defaultdict(Counter)
    confusion: Dict[Tuple[str, str], int] = Counter()
    for r in rows:
        stratum = r.get("stratum", "?")
        label = (r.get("human_label") or "").strip()
        per[stratum]["n"] += 1
        if label not in LABELS:
            per[stratum]["unlabelled"] += 1
            continue
        if label == "?":
            per[stratum]["unclear"] += 1
            continue
        per[stratum]["labelled"] += 1
        if label == "fine":
            per[stratum]["fine"] += 1
        else:
            per[stratum]["defect"] += 1
        stored = r.get("stored_category") or ("supported" if r.get("stored_verdict") == "supported" else "none")
        if r.get("stored_verdict") == "supported":
            stored = "supported"
        confusion[(stored, label)] += 1

    print("--- COVERAGE ---")
    header = f"{'stratum':<24} {'n':>5} {'labelled':>9} {'unclear':>8} {'blank':>7}"
    print(header)
    print("-" * len(header))
    for stratum in sorted(per):
        c = per[stratum]
        print(f"{stratum:<24} {c['n']:>5} {c['labelled']:>9} {c['unclear']:>8} {c['unlabelled']:>7}")
    total_labelled = sum(per[s]["labelled"] for s in per)
    if total_labelled == 0:
        sys.exit("\n❌ Nothing labelled yet — fill the human_label column first.")

    flagged = [s for s in per if s.startswith("flagged_")]
    supported = [s for s in per if s.startswith("supported_")]

    print("\n--- PER-STRATUM RATES (point estimate, 95% Wilson) ---")
    for stratum in flagged:
        c = per[stratum]
        if not c["labelled"]:
            continue
        lo, hi = wilson(c["defect"], c["labelled"])
        print(f"  {stratum:<24} precision {c['defect']}/{c['labelled']} = "
              f"{c['defect'] / c['labelled']:.0%}  [{lo:.0%}, {hi:.0%}]")
    for stratum in supported:
        c = per[stratum]
        if not c["labelled"]:
            continue
        lo, hi = wilson(c["defect"], c["labelled"])
        print(f"  {stratum:<24} miss rate {c['defect']}/{c['labelled']} = "
              f"{c['defect'] / c['labelled']:.0%}  [{lo:.0%}, {hi:.0%}]")

    print("\n--- POPULATION-WEIGHTED ---")
    if not weights:
        print("  ⚠️  No .frame.yml sidecar found next to the CSV(s) — the stratum")
        print("      population weights are unknown, so the numbers below are the")
        print("      UNWEIGHTED sample rates. They over-state the miss rate badly,")
        print("      because the supported strata are sampled far below their weight.")
    est_tp = est_fp = est_fn = 0.0
    for stratum in flagged + supported:
        c = per[stratum]
        if not c["labelled"]:
            continue
        pop = weights.get(stratum, {}).get("pop", c["n"])
        rate = c["defect"] / c["labelled"]
        if stratum.startswith("flagged_"):
            est_tp += rate * pop
            est_fp += (1 - rate) * pop
        else:
            est_fn += rate * pop
        print(f"  {stratum:<24} weight {pop:>6} occ × defect-rate {rate:.0%}")

    precision = est_tp / (est_tp + est_fp) if (est_tp + est_fp) else 0.0
    recall = est_tp / (est_tp + est_fn) if (est_tp + est_fn) else 0.0
    f1 = 2 * precision * recall / (precision + recall) if (precision + recall) else 0.0
    print(f"\n  estimated true positives  {est_tp:9.1f} occurrences")
    print(f"  estimated false positives {est_fp:9.1f}")
    print(f"  estimated misses (FN)     {est_fn:9.1f}")
    print(f"\n  PRECISION {precision:.1%}   RECALL {recall:.1%}   F1 {f1:.1%}")

    # The point estimates above are the honest centre; on their own they are also
    # the most quotable and the most misleading number this tool produces. A
    # supported stratum weighing 2,657 occurrences is being extrapolated from six
    # labelled rows: one row changing hands moves the estimated miss count by
    # ~440, which moves recall by tens of points. So the envelope is printed
    # beside it, propagated from the same Wilson intervals shown per stratum —
    # worst case takes every flagged rate at its lower bound and every supported
    # rate at its upper, and vice versa.
    lo_tp = lo_fp = lo_fn = hi_tp = hi_fp = hi_fn = 0.0
    for stratum in flagged + supported:
        c = per[stratum]
        if not c["labelled"]:
            continue
        pop = weights.get(stratum, {}).get("pop", c["n"])
        lo, hi = wilson(c["defect"], c["labelled"])
        if stratum.startswith("flagged_"):
            lo_tp += lo * pop
            lo_fp += (1 - lo) * pop
            hi_tp += hi * pop
            hi_fp += (1 - hi) * pop
        else:
            lo_fn += hi * pop   # worst case for recall: more misses
            hi_fn += lo * pop
    worst_p = lo_tp / (lo_tp + lo_fp) if (lo_tp + lo_fp) else 0.0
    best_p = hi_tp / (hi_tp + hi_fp) if (hi_tp + hi_fp) else 0.0
    worst_r = lo_tp / (lo_tp + lo_fn) if (lo_tp + lo_fn) else 0.0
    best_r = hi_tp / (hi_tp + hi_fn) if (hi_tp + hi_fn) else 0.0
    print(f"  envelope from the per-stratum 95% intervals: "
          f"precision {worst_p:.0%}–{best_p:.0%}, recall {worst_r:.0%}–{best_r:.0%}")
    if best_r - worst_r > 0.25:
        thin = [s for s in supported if 0 < per[s]["labelled"] < 15]
        print(f"  ⚠️  the recall envelope spans {best_r - worst_r:.0%} — it is not a measurement yet.")
        if thin:
            print(f"      Thinnest strata carrying the most weight: {', '.join(thin)}.")
            print(f"      Re-export with a larger --n, or with the supported strata")
            print(f"      widened, before quoting a recall figure to anyone.")
    unclear = sum(per[s]["unclear"] for s in per)
    if unclear:
        labelled_all = sum(per[s]["labelled"] for s in per)
        print(f"\n  {unclear} row(s) marked '?' are excluded from every rate above "
              f"({unclear / (unclear + labelled_all):.0%} of what was reviewed).")
        print("  A high share means the rows did not carry enough to judge on — that is")
        print("  a finding about this export, not about the verifier.")

    print("\n--- CONFUSION MATRIX (stored category × human label) ---")
    stored_keys = ["hallucination", "reasoning", "none", "", "supported"]
    stored_keys = [k for k in stored_keys if any(s == k for s, _ in confusion)]
    label_keys = [l for l in LABELS if l != "?"]
    head = f"{'stored ↓ human →':<20}" + "".join(f"{l:>16}" for l in label_keys) + f"{'row total':>12}"
    print(head)
    print("-" * len(head))
    for stored in stored_keys:
        cells = [confusion.get((stored, l), 0) for l in label_keys]
        name = stored or "(blank)"
        print(f"{name:<20}" + "".join(f"{v:>16}" for v in cells) + f"{sum(cells):>12}")
    col_total = [sum(confusion.get((s, l), 0) for s in stored_keys) for l in label_keys]
    print(f"{'column total':<20}" + "".join(f"{v:>16}" for v in col_total) + f"{sum(col_total):>12}")

    agree = sum(confusion.get((c, c), 0) for c in ("hallucination", "reasoning"))
    typed = sum(v for (s, l), v in confusion.items()
                if s in ("hallucination", "reasoning") and l in ("hallucination", "reasoning"))
    if typed:
        print(f"\n  Of the {typed} flagged rows the reviewer also calls a defect, the stored")
        print(f"  CATEGORY matches the reviewer's in {agree}/{typed} = {agree / typed:.0%}.")
        print("  A low number here does not move precision — it means the category, not")
        print("  the flag, is what cannot be routed on.")


# --------------------------------------------------------------------------

def main() -> None:
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("databook", nargs="?",
                    help="path to the databook these runs were generated from "
                         "(required for export; not needed with --score)")
    ap.add_argument("--run", action="append", default=[],
                    help="run folder, name or timestamp; repeatable. Default: every "
                         "run under fdd_utils/logs that pairs with the databook.")
    ap.add_argument("--out", help="CSV to write (export mode)")
    ap.add_argument("--n", type=int, default=None,
                    help="total sample size; strata scale proportionally (default 100 — see DEFAULT_STRATA for why that is not the plan's 150)")
    ap.add_argument("--seed", type=int, default=7, help="sampling seed (default 7)")
    ap.add_argument("--entity", default=None, help="entity name passed to process_workbook_data")
    ap.add_argument("--sheet", default=None, help="Financials sheet (default: auto-resolve)")
    ap.add_argument("--score", action="append", default=[],
                    help="filled CSV to score; repeatable, to pool several databooks")
    args = ap.parse_args()

    if args.score:
        do_score(args)
        return
    if not args.databook or not args.out:
        ap.error("export mode needs a databook and --out (or use --score)")
    do_export(args)


if __name__ == "__main__":
    main()
