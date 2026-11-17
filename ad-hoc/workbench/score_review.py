"""Free; no model call. Lay a reviewer's list of problem sentences against the marks the deck showed, by account and sentence number.

    python ad-hoc/workbench/score_review.py --stamp 20261117_150703 --stamp 20261117_150858 ...

For each stamp it reads, from fdd_utils/output/agent_trial:
    trial_<arm>_<stamp>.agent-session.json     the drafts and the marks (the deck's colours, as trial_digest.py reads them)
    review_gpt_<stamp>.txt                     the model reviewer's reply to the blind file (blind_review.py)
    review_human_<stamp>.txt                   a person's list, same line format (optional)
A missing review file is reported, not an error. The reply may hold other text and a code fence; only these lines are read, with
a half-width or full-width bar:

    REVIEWED|account,account,...
    FINDING|account|sentence number|category|重大 or 輕微|確定 or 爭議|reason

The two reviewers are reported separately and never merged. "Unmarked" means the sentence the reviewer flagged has no red or orange
mark in the deck. The categories are the reviewer's own; a cause that spans categories is judged by a person, so every unmarked
major finding is listed with its sentence. The output holds client sentences: it stays in the chat it is pasted to.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter
from pathlib import Path
from types import SimpleNamespace
from typing import Any, Dict, List, Optional

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ROOT))

TRIAL_DIR = ROOT / "fdd_utils" / "output" / "agent_trial"
_BAR = r"[|｜]"
_MAJOR, _MINOR = ("重大", "major"), ("輕微", "轻微", "minor")
_SURE, _DISPUTED = ("確定", "确定", "sure", "confirmed"), ("爭議", "争议", "disputed", "uncertain")


def parse(text: str) -> Dict[str, Any]:
    """{'reviewed': [account], 'findings': [dict], 'bad': [line]} from a reviewer's reply."""
    reviewed: List[str] = []
    findings: List[Dict[str, Any]] = []
    bad: List[str] = []
    for raw in str(text or "").splitlines():
        line = raw.strip().strip("`").strip()
        head = re.match(rf"(REVIEWED|FINDING)\s*{_BAR}\s*(.*)$", line)
        if not head:
            continue
        fields = [f.strip() for f in re.split(_BAR, head.group(2))]
        if head.group(1) == "REVIEWED":
            reviewed.extend(a.strip() for a in re.split(r"[,，、]", " ".join(fields)) if a.strip())
            continue
        number = re.search(r"\d+", fields[1]) if len(fields) > 1 else None
        severity = _word(fields[3], _MAJOR, _MINOR) if len(fields) > 3 else None
        certainty = _word(fields[4], _SURE, _DISPUTED) if len(fields) > 4 else None
        if len(fields) < 6 or not number or severity is None or certainty is None:
            bad.append(line[:160])
            continue
        findings.append({"account": fields[0], "n": int(number.group(0)), "category": fields[2], "severity": severity,
                         "certainty": certainty, "reason": " ".join(fields[5:])})
    return {"reviewed": reviewed, "findings": findings, "bad": bad}


def _word(value: str, first: Any, second: Any) -> Optional[str]:
    """'first' or 'second' by which word list the field names; 不确定 reads as disputed, not sure."""
    v = value.strip().lower()
    if first is _SURE and v.startswith("不"):
        return "second"
    if any(w in v for w in first):
        return "first"
    if any(w in v for w in second):
        return "second"
    return None


def marks_by_account(session: Dict[str, Any]) -> Dict[str, Any]:
    """{account: {'sentences': [text], 'marks': {sentence number: (colour, why)}}} as the deck coloured them."""
    from entity_trial import deck_marks
    out: Dict[str, Any] = {}
    for key, c in (session.get("candidates") or {}).items():
        claims = list(c.get("claims") or [])
        if not claims:
            continue
        marks = deck_marks(key, SimpleNamespace(claims=claims, outcome=c.get("outcome") or {}, flags=c.get("flags") or []))
        out[key] = {"sentences": [str(x.get("text") or "") for x in claims], "marks": {m["claim"]: (m["colour"], m["why"]) for m in marks}}
    return out


def score_one(books: Dict[str, Any], reply: Optional[str], width: int = 110) -> Dict[str, Any]:
    """The counts and the unmarked major findings of one reviewer's reply against one book."""
    if reply is None:
        return {"missing": True}
    got = parse(reply)
    known = set(books)
    reviewed = [a for a in got["reviewed"] if a in known]
    unknown_reviewed = [a for a in got["reviewed"] if a not in known]
    rows, invalid = [], []
    for f in got["findings"]:
        book = books.get(f["account"])
        if book is None or not 0 <= f["n"] < len(book["sentences"]):
            invalid.append(f)
            continue
        mark = book["marks"].get(f["n"])
        rows.append({**f, "major": f["severity"] == "first", "sure": f["certainty"] == "first",
                     "mark": mark[0] if mark else None, "sentence": book["sentences"][f["n"]]})
    unmarked_major_sure = [r for r in rows if r["major"] and r["sure"] and not r["mark"]]
    unmarked_major_disputed = [r for r in rows if r["major"] and not r["sure"] and not r["mark"]]
    flagged = {(r["account"], r["n"]) for r in rows}
    unflagged_marks = [(a, n, m[0], m[1]) for a, b in books.items() if a in reviewed for n, m in sorted(b["marks"].items())
                       if (a, n) not in flagged]
    return {"missing": False, "reviewed": reviewed, "not_reviewed": [a for a in books if a not in reviewed],
            "unknown_reviewed": unknown_reviewed, "rows": rows, "invalid": invalid, "bad": got["bad"],
            "unmarked_major_sure": unmarked_major_sure, "unmarked_major_disputed": unmarked_major_disputed,
            "unflagged_marks": unflagged_marks,
            "sentences_reviewed": sum(len(books[a]["sentences"]) for a in reviewed), "width": width}


def lines_for(name: str, res: Dict[str, Any], books: Dict[str, Any]) -> List[str]:
    if res.get("missing"):
        return [f" {name}: no review file"]
    rows = res["rows"]
    out = [f" {name}: reviewed {len(res['reviewed'])}/{len(books)} accounts ({res['sentences_reviewed']} sentences); findings {len(rows)} "
           f"(major {sum(r['major'] for r in rows)}: sure {sum(r['major'] and r['sure'] for r in rows)}, disputed "
           f"{sum(r['major'] and not r['sure'] for r in rows)}; minor {sum(not r['major'] for r in rows)}); "
           f"flagged and marked {sum(1 for r in rows if r['mark'])}; UNMARKED major sure {len(res['unmarked_major_sure'])}, "
           f"unmarked major disputed {len(res['unmarked_major_disputed'])}; invalid references {len(res['invalid'])}; "
           f"unreadable lines {len(res['bad'])}"]
    if res["not_reviewed"]:
        out.append(f"   not reviewed: {', '.join(res['not_reviewed'])}")
    if res["unknown_reviewed"]:
        out.append(f"   REVIEWED names that match no account: {', '.join(res['unknown_reviewed'])}")
    width = res["width"]
    for title, items in (("unmarked major, sure", res["unmarked_major_sure"]), ("unmarked major, disputed", res["unmarked_major_disputed"])):
        if items:
            out.append(f"   {title}:")
            for r in items:
                out.append(f"     {r['account']}[{r['n']}] {r['category']}: {r['reason'][:width]}")
                out.append(f"         {r['sentence'][:width]}")
    for r in res["invalid"]:
        out.append(f"   invalid reference (no such account or sentence): {r['account']}[{r['n']}] {r['category']}")
    for line in res["bad"][:5]:
        out.append(f"   unreadable: {line}")
    if res["unflagged_marks"]:
        out.append(f"   marked sentences this reviewer did not flag ({len(res['unflagged_marks'])}; on accounts it reviewed): "
                   + "; ".join(f"{a}[{n}] {c[0].upper()}" for a, n, c, _w in res["unflagged_marks"][:12]))
    return out


def find(stamp: str, folder: Path) -> Dict[str, Any]:
    sessions = sorted(folder.glob(f"trial_*_{stamp}.agent-session.json"))
    return {"session": sessions[0] if sessions else None,
            "gpt": folder / f"review_gpt_{stamp}.txt", "human": folder / f"review_human_{stamp}.txt"}


def _read(path: Path) -> Optional[str]:
    if not path.exists():
        return None
    raw = path.read_bytes()
    try:
        return raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        return raw.decode("gbk", errors="replace")


def score_all(stamps: List[str], folder: Path = TRIAL_DIR) -> List[str]:
    out: List[str] = []
    totals: Dict[str, Counter] = {"gpt": Counter(), "human": Counter()}
    by_category: Dict[str, Counter] = {"gpt": Counter(), "human": Counter()}
    for stamp in stamps:
        files = find(stamp, folder)
        if files["session"] is None:
            out.append(f"===== {stamp}: no saved session in {folder}")
            continue
        session = json.loads(files["session"].read_text(encoding="utf-8"))
        books = marks_by_account(session)
        every = [m for b in books.values() for m in b["marks"].values()]
        out.append(f"===== trial {stamp} ({files['session'].name.split('_')[1]}): {len(books)} accounts, "
                   f"{sum(len(b['sentences']) for b in books.values())} sentences; deck marks {sum(c == 'red' for c, _w in every)} red "
                   f"{sum(c == 'orange' for c, _w in every)} orange")
        for who in ("gpt", "human"):
            res = score_one(books, _read(files[who]))
            out.extend(lines_for(who, res, books))
            if not res.get("missing"):
                totals[who].update(accounts=len(res["reviewed"]), of=len(books), findings=len(res["rows"]),
                                   unmarked_sure=len(res["unmarked_major_sure"]), unmarked_disputed=len(res["unmarked_major_disputed"]),
                                   invalid=len(res["invalid"]), books=1)
                by_category[who].update(r["category"] for r in res["unmarked_major_sure"])
    out.append("===== ALL BOOKS (each reviewer on its own; never added together)")
    for who in ("gpt", "human"):
        t = totals[who]
        if not t["books"]:
            out.append(f" {who}: no review file read")
            continue
        out.append(f" {who}: {t['books']} book(s), accounts reviewed {t['accounts']}/{t['of']}, findings {t['findings']}, "
                   f"UNMARKED major sure {t['unmarked_sure']} (registered ceiling 12 over four books for the model reviewer), "
                   f"unmarked major disputed {t['unmarked_disputed']}, invalid references {t['invalid']}")
        if by_category[who]:
            out.append("   unmarked major sure by the reviewer's category (a cause may span categories; the ceiling is 3 from one cause): "
                       + ", ".join(f"{c} {n}" for c, n in by_category[who].most_common()))
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--stamp", action="append", required=True, help="a trial stamp, e.g. 20261117_150703; repeat for each book")
    ap.add_argument("--dir", default=str(TRIAL_DIR))
    args = ap.parse_args()
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    print("\n".join(score_all(args.stamp, Path(args.dir))))
    return 0


if __name__ == "__main__":
    sys.exit(main())
