"""How much of a finished report each trial arm states, from a trial's review file and the report's text. No model call; free.

    python ad-hoc/workbench/score_reference.py --reference <report text> [--trial <trial_quality_*.md>] [--verbose]

The reference is a text dump of the team's finished deck (`inspect_pptx.py <deck> --dump-text`, saved to a file); an account's
paragraph runs from its name followed by a dash ("应收账款 –") to the next such header. For every account the trial wrote, each
amount and percentage of the reference paragraph is looked for in each arm's paragraph and in the legacy run's (same
tolerance as the verifier: half the last written digit). Three readings per arm:

  accepted   only paragraphs the gate accepted (what a clean deck would carry)
  in deck    accepted plus the drafts carried with their rejected sentences marked (a person corrects those)
  all text   every paragraph the arm wrote, accepted or not

and the length of each paragraph against the reference's, since a longer paragraph states more figures and does not fit the
slot it was sized for. Recall says what is covered; it does not say a sentence is right (the review form in the trial file
does), and a figure can be stated in the right paragraph for the wrong thing.
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Dict, List, Tuple

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))


def _read(path: str) -> str:
    raw = Path(path).read_bytes()
    try:
        return raw.decode("utf-8-sig")
    except UnicodeDecodeError:                 # a file written by `>` in a Windows console is in its code page
        import locale
        return raw.decode(locale.getpreferredencoding(False) or "gbk", errors="replace")


def _paragraphs(text: str, accounts: List[str]) -> Dict[str, str]:
    heads = sorted((m.start(), a) for a in accounts for m in re.finditer(re.escape(a) + r"(?:（续）)?\s*[–—-]", text))
    out: Dict[str, str] = {}
    for i, (pos, a) in enumerate(heads):
        end = heads[i + 1][0] if i + 1 < len(heads) else len(text)
        body = re.split(r"\n\s*\[|\n--- Slide|资料来源", text[pos:end])[0]      # stop at the next shape, slide or source line
        out[a] = out.get(a, "") + body
    return out


def _arm_paragraph(block: str, arm: str) -> Tuple[str, str, int]:
    """(text, end state, verified sentences) of one arm in one case block; legacy has no end state."""
    if arm == "legacy":
        m = re.search(r"\*\*legacy\*\* \(\d+ chars\)\s*\n(.*?)\n\s*\n\*\*", block, re.S)
        return (m.group(1).strip(), "done", 1) if m else ("", "missing", 0)
    m = re.search(r"\n\*\*" + re.escape(arm) + r"\*\*: end `([^`]+)`; mechanical (\{[^}]*\})[^\n]*\n\s*\n(.*?)\n\s*\n", block, re.S)
    if not m:
        return "", "missing", 0
    verified = sum(int(x) for x in re.findall(r"'(?:verified|inference)': (\d+)", m.group(2)))
    return m.group(3).strip(), m.group(1), verified


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--trial", default="", help="trial_quality_<stamp>.md (default: the newest in fdd_utils/output/agent_trial)")
    ap.add_argument("--reference", required=True, help="the finished report's text")
    ap.add_argument("--verbose", action="store_true", help="list the reference figures each arm misses")
    args = ap.parse_args()
    from fdd_utils.ai import claims as cl

    if not args.trial:
        found = sorted((ROOT / "fdd_utils" / "output" / "agent_trial").glob("trial_quality_*.md"), key=lambda f: f.stat().st_mtime)
        if not found:
            print("no trial_quality_*.md in fdd_utils/output/agent_trial; pass --trial")
            return 2
        args.trial = str(found[-1])
    print(f"trial: {Path(args.trial).name}")
    md, ref = _read(args.trial), _read(args.reference)
    blocks = {}
    for b in re.split(r"\n## ", md)[1:]:
        m = re.match(r"case \d+: (.+?) \(", b)
        if m:
            blocks[m.group(1)] = b
    arms = re.search(r"Arms: ([^.\n]+)\.", md)
    arms = [a.strip() for a in arms.group(1).split(",")] if arms else ["pro", "flash"]
    names = arms + ["legacy"]
    paras = _paragraphs(ref, list(blocks))

    def amounts(t: str):
        return [(abs(s["value"]), s["written"], s["tolerance"]) for s in cl.text_amounts(t) if s["value"]]

    def pcts(t: str):
        return [(abs(s["value"]), s["written"]) for s in cl.text_percents(t) if not s.get("points")]

    def length(t: str) -> int:
        return len(re.sub(r"\s", "", t))

    tot = {a: {"accepted": 0, "in deck": 0, "all text": 0, "chars": 0, "n": 0} for a in names}
    ref_chars = 0
    print(f"{'account':12s} {'ref':>4s} " + " ".join(f"{a[:7]:>9s}" for a in names) + "   (figures stated; * not accepted, "
          "+ in the deck marked; chars in brackets)")
    for acc, block in blocks.items():
        tp = paras.get(acc, "")
        figs, ps = amounts(tp), pcts(tp)
        n = len(figs) + len(ps)
        if not n:
            continue
        ref_chars += length(tp)
        cells, missed = [], {}
        for arm in names:
            text, end, verified = _arm_paragraph(block, arm)
            mine, myp = amounts(text), pcts(text)
            hit = [w for v, w, t in figs if any(abs(v - mv) <= max(t, mt) for mv, _w, mt in mine)]
            hit += [w for v, w in ps if any(abs(v - x) <= 0.05 for x, _ in myp)]
            accepted = end == "done"
            in_deck = accepted or (bool(text) and verified > 0)
            tot[arm]["accepted"] += len(hit) if accepted else 0
            tot[arm]["in deck"] += len(hit) if in_deck else 0
            tot[arm]["all text"] += len(hit)
            tot[arm]["chars"] += length(text) if in_deck else 0
            tot[arm]["n"] += n
            cells.append(f"{len(hit):2d}{'' if accepted else ('+' if in_deck else '*')}[{length(text):4d}]")
            missed[arm] = [w for v, w, t in figs if w not in hit] + [w for v, w in ps if w not in hit]
        print(f"{acc[:12]:12s} {n:4d} " + " ".join(f"{c:>9s}" for c in cells))
        if args.verbose:
            for arm in names:
                if missed[arm]:
                    print(f"    {arm} misses: {', '.join(missed[arm])}")
    print(f"\nreference paragraphs: {ref_chars:,} characters")
    for arm in names:
        t = tot[arm]
        n = max(t["n"], 1)
        print(f"{arm:8s} accepted {t['accepted']}/{t['n']} = {t['accepted'] / n:.0%}   in deck {t['in deck']}/{t['n']} = "
              f"{t['in deck'] / n:.0%}   all text {t['all text']}/{t['n']} = {t['all text'] / n:.0%}   "
              f"length in deck {t['chars']:,} ({t['chars'] / max(ref_chars, 1):.1f}x the reference)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
