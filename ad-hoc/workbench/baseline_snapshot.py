"""Free; no model call. Proof that a change to the code left what the agent path does exactly as it was.

    python ad-hoc/workbench/baseline_snapshot.py save  <label>      write the snapshot of the code as it is now
    python ad-hoc/workbench/baseline_snapshot.py check <label>      read it again with the code as it is now; any difference exits 1
    python ad-hoc/workbench/baseline_snapshot.py check <label> --books "Project Gold"      only the books whose name contains this
    python ad-hoc/workbench/baseline_snapshot.py save  <label> --dir databooks --books A   books in another folder (the paid sessions found for them are read too)

Per local databook (every .xlsx in the repo root whose sessions or contracts can be read), three things are hashed:
    CONTRACTS   every account's evidence contract, as extracted now (what Python reads from the workbook)
    GENERATED   the sentences claim_eval.py and claim_shapes.py write from every account's own figures (the right ones, and each broken in
                a named way, in the shapes a writer uses) through the verifier: verdict, each finding's severity and code, a hash of its message
    SESSIONS    every claim of every saved paid session of that workbook (a scripted run's session says nothing about a writer and is left out;
                this machine has none, a machine that ran the paid books has many), read again by the verifier the same way
    DECK        one scripted whole-book run (no model call): which accounts finish, and per slide the text, geometry, font size, bold and
                colour of every shape and table cell, and the delivery decision

A snapshot is only as wide as these inputs. It is local (fdd_utils/output/baseline, git-ignored: the sessions hold client sentences) and
holds hashes plus the codes, never a sentence. Run `check` after each step of a cleanup; a difference is named by book, account and claim.
"""
from __future__ import annotations

import argparse
import glob
import hashlib
import json
import shutil
import sys
import time
import warnings
from pathlib import Path
from typing import Any, Dict, List, Optional

warnings.filterwarnings("ignore")
HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ROOT))

OUT = ROOT / "fdd_utils" / "output" / "baseline"
TRIALS = ROOT / "fdd_utils" / "output" / "agent_trial"


def sha(obj: Any) -> str:
    return hashlib.sha256(json.dumps(obj, ensure_ascii=False, sort_keys=True, default=str).encode("utf-8")).hexdigest()[:16]


BOOKS_DIR = ROOT


def local_books() -> Dict[str, str]:
    return {hashlib.sha256(Path(b).read_bytes()).hexdigest(): b for b in sorted(glob.glob(str(BOOKS_DIR / "*.xlsx")))
            if "reconciliation_report" not in b and not Path(b).name.startswith("~$")}


def open_work(book: str):
    from fdd_utils.agent.workspace import WorkSession
    return WorkSession.open(book, model=lambda *a, **k: {}, model_name="baseline")


def contracts_part(work: Any) -> Dict[str, Any]:
    items = {k: sha(c.to_dict()) for k, c in sorted(work.source.contracts.items())}
    return {"n": len(items), "sha": sha(items), "items": items}


def sessions_part(work: Any, book: str, sha_book: str, files: Optional[List[str]] = None) -> Dict[str, Any]:
    """`files`: the session files a snapshot was taken over (a check reads those and no others, so a paid run made since does not count
    as a difference). Without it, every saved session of this workbook that a person's run wrote: a scripted run or this tool's own
    session says nothing about how a writer writes and is left out."""
    from fdd_utils.ai import claims as cl
    items: Dict[str, Any] = {}
    used: List[str] = []
    for f in (files if files is not None else sorted(glob.glob(str(ROOT / "fdd_utils" / "output" / "**" / "*.agent-session.json"), recursive=True))):
        try:
            saved = json.loads(Path(f).read_text(encoding="utf-8"))
        except Exception:
            continue
        if (((saved.get("source") or {}).get("binding") or {}).get("workbook_sha256")) != sha_book:
            continue
        model = saved.get("model")
        name_of_model = str((model or {}).get("name") if isinstance(model, dict) else model or "")
        if files is None and name_of_model.startswith(("scripted", "baseline")):
            continue
        used.append(f)
        name = Path(f).name
        for key, cand in sorted((saved.get("candidates") or {}).items()):
            claims = [dict(c, account=key, source_version=work.source.source_version) for c in cand.get("claims") or []]
            if not claims or key not in work.source.contracts:
                continue
            outcome = cl.verify_candidate(claims, work.source.contracts, default_account=key)
            for n, v in enumerate(outcome["verdicts"]):
                findings = sorted((f_.get("severity"), f_.get("code"), sha(f_.get("message"))[:8]) for f_ in v.get("findings") or [])
                items[f"{name}|{key}|{n}"] = [v.get("verdict"), findings]
    return {"sessions": len(used), "files": used, "claims": len(items), "sha": sha(items), "items": items}


def sig(verdict: Dict[str, Any]) -> List[Any]:
    return [verdict.get("verdict"), sorted((f.get("severity"), f.get("code"), sha(f.get("message"))[:8]) for f in verdict.get("findings") or [])]


TEXTS: Dict[str, str] = {}      # the sentence behind each key of this run, for naming a difference; never written to the snapshot


def generated_part(contracts: Dict[str, Any], book: str) -> Dict[str, Any]:
    import claim_eval as ce
    import claim_shapes as csh
    from fdd_utils.ai import claims as cl
    items: Dict[str, Any] = {}
    for key, contract in sorted(contracts.items()):
        for name, claim, _ in ce.correct_claims(contract, cl):
            k = f"eval|{key}|C:{name}"
            items[k] = sig(cl.verify_claim(claim, {key: contract}))
            TEXTS[f"{book}|{k}"] = claim["text"]
            kinds = ce.kind_sweep(claim, key, contract, cl, False)
            items[f"eval|{key}|K:{name}"] = [kinds[x] for x in sorted(kinds)]
            for n, (mname, broken) in enumerate(ce.mutations(name, claim, contract, cl)):
                k = f"eval|{key}|B:{name}:{mname}:{n}"
                items[k] = sig(cl.verify_claim(broken, {key: contract}))
                TEXTS[f"{book}|{k}"] = broken["text"]
        for n, r in enumerate(csh.evaluate(contract, cl, contracts)):
            k = f"shape|{key}|{n}:{r['kind']}:{r['shape']}"
            items[k] = [r["ok"], r["codes"], sha(r["text"])[:8]]
            TEXTS[f"{book}|{k}"] = r["text"]
    return {"claims": len(items), "sha": sha(items), "items": items}


def deck_digest(path: str) -> Dict[str, Any]:
    from pptx import Presentation
    from pptx.util import Emu

    def run_sig(r: Any) -> List[Any]:
        f = r.font
        try:
            colour = str(f.color.rgb) if f.color is not None and f.color.type is not None and f.color.rgb is not None else None
        except Exception:
            colour = None
        return [r.text, f.size.pt if f.size else None, f.bold, colour]

    def frame_sig(tf: Any) -> List[Any]:
        return [[run_sig(r) for r in p.runs] + [p.level] for p in tf.paragraphs]

    def cell_fill(cell: Any) -> Optional[str]:
        try:
            return str(cell.fill.fore_color.rgb) if cell.fill.type == 1 else None
        except Exception:
            return None

    prs = Presentation(path)
    slides = []
    for s in prs.slides:
        shapes = []
        for sh in s.shapes:
            sig: Dict[str, Any] = {"name": sh.name, "box": [sh.left, sh.top, sh.width, sh.height]}
            if sh.has_text_frame:
                sig["text"] = frame_sig(sh.text_frame)
            if getattr(sh, "has_table", False) and sh.has_table:
                t = sh.table
                sig["table"] = [[frame_sig(c.text_frame) + [cell_fill(c)] for c in row.cells] for row in t.rows]
                sig["rows"] = [r.height for r in t.rows]
                sig["cols"] = [c.width for c in t.columns]
            shapes.append(sig)
        slides.append(sha(shapes))
    return {"slides": len(slides), "per_slide": slides, "sha": sha(slides)}


def scripted_part(book: str) -> Dict[str, Any]:
    import entity_trial
    started = time.time()
    entity_trial.main(["run", "--databook", book, "--arms", "flash", "--scripted", "--skip-gate"])
    reports = [p for p in TRIALS.glob("trial_2*.json") if p.stat().st_mtime >= started - 1 and ".progress" not in p.name]
    report = json.loads(max(reports, key=lambda p: p.stat().st_mtime).read_text(encoding="utf-8"))
    arm = report["arms"]["flash"]
    out: Dict[str, Any] = {"accounts": arm.get("accounts"), "ok": arm.get("ok"), "delivery": arm.get("delivery"),
                           "findings": arm.get("findings"), "marked": arm.get("marked"), "export_error": arm.get("export_error")}
    deck = ROOT / (arm.get("deck") or "")
    out["deck"] = deck_digest(str(deck)) if arm.get("deck") and deck.exists() else None
    session = ROOT / (arm.get("session_file") or "")
    if arm.get("session_file") and session.exists():
        saved = json.loads(session.read_text(encoding="utf-8"))
        out["session"] = sha({k: [c.get("claims"), (c.get("outcome") or {}).get("by_code"), c.get("accepted")]
                              for k, c in sorted((saved.get("candidates") or {}).items())})
    return out


def book_part(book: str) -> Dict[str, Any]:
    """The same scripted flow through the product path (fdd_utils/agent/book.py, which the app and fdd_agent.py book use), in the
    fields scripted_part reads from the trial's report: it must equal it, for the app and the trial write the same deck."""
    import tempfile
    import entity_trial
    from fdd_utils.agent import book as B
    from fdd_utils.agent.flag import agent_budget
    with tempfile.TemporaryDirectory(prefix="baseline_book_") as tmp:
        # the trial works on a copy named after its arm and the deck's title is the workbook's name: the same name here
        copy = Path(tmp) / "flash.xlsx"
        shutil.copyfile(book, copy)
        work = B.open_book(str(copy), model=entity_trial.ScriptedWriter("flash"), model_name="scripted-flash")
        result = B.run_book(work, issues_dir=None, out_path=str(Path(tmp) / "book.pptx"))
        data = result.deck
        findings: Dict[str, int] = {}
        for f in (data.get("delivery") or {}).get("findings") or []:
            findings[f.get("code")] = findings.get(f.get("code"), 0) + 1
        from fdd_utils.agent import tools as T
        from fdd_utils.ai import claims as cl
        marked = {"accounts": 0, "red": 0, "orange": 0}
        for key in T.drafts_to_mark(work.source):
            marked["accounts"] += 1
            for v in work.source.candidates[key].outcome.get("verdicts") or []:
                if v.get("verdict") in ("failed", "unverified"):
                    red = any(f.get("severity") == "failed" and f.get("dimension") in cl._RED_DIMENSIONS for f in v.get("findings") or [])
                    marked["red" if red else "orange"] += 1
        out: Dict[str, Any] = {"accounts": len(result.accounts), "ok": sum(1 for r in result.rows if r.get("ok")),
                               "delivery": (data.get("delivery") or {}).get("status"), "findings": findings, "marked": marked,
                               "export_error": result.error.get("code"),
                               "deck": deck_digest(result.deck_path) if result.deck_path and Path(result.deck_path).exists() else None}
        saved = work.to_dict()
        out["session"] = sha({k: [c.get("claims"), (c.get("outcome") or {}).get("by_code"), c.get("accepted")]
                              for k, c in sorted(((saved.get("source") or saved).get("candidates") or saved.get("candidates") or {}).items())})
    return out


def snapshot(only: str = "", was: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    books: Dict[str, Any] = {}
    for sha_book, book in local_books().items():
        name = Path(book).name
        if only and only not in name:
            continue
        print(f"  {name} ...", flush=True)
        work = open_work(book)
        files = ((was or {}).get("books", {}).get(name, {}).get("sessions", {}).get("files")) if was else None
        books[name] = {"contracts": contracts_part(work), "generated": generated_part(work.source.contracts, name),
                       "sessions": sessions_part(work, book, sha_book, files), "scripted": scripted_part(book),
                       "book": book_part(book)}
    import entity_trial
    return {"code": entity_trial.code_fingerprint()["short"], "taken": time.strftime("%Y-%m-%d %H:%M:%S"), "books": books}


def diff(a: Dict[str, Any], b: Dict[str, Any], limit: int = 12) -> List[str]:
    lines: List[str] = []
    for name in sorted(set(a["books"]) | set(b["books"])):
        x, y = a["books"].get(name), b["books"].get(name)
        if x is None or y is None:
            lines.append(f"{name}: only in {'the snapshot' if y is None else 'the check'}")
            continue
        for part in ("contracts", "generated", "sessions"):
            if x[part]["sha"] != y[part]["sha"]:
                ix, iy = x[part]["items"], y[part]["items"]
                keys = sorted(k for k in set(ix) | set(iy) if ix.get(k) != iy.get(k))
                lines.append(f"{name}: {part.upper()} differ in {len(keys)} of {max(len(ix), len(iy))}")
                for k in keys[:limit]:
                    lines.append(f"    {k}: was {ix.get(k)}  now {iy.get(k)}")
                    if f"{name}|{k}" in TEXTS:
                        lines.append(f"        sentence: {TEXTS[f'{name}|{k}'][:140]}")
        sx, sy = x["scripted"], y["scripted"]
        for field in ("accounts", "ok", "delivery", "findings", "marked", "export_error", "session"):
            if sx.get(field) != sy.get(field):
                lines.append(f"{name}: SCRIPTED {field} was {sx.get(field)}  now {sy.get(field)}")
        # the product path (book.py) writes what the trial writes: the fresh one is held to the saved trial's, field by field
        sb = y.get("book") or {}
        for field in ("accounts", "ok", "delivery", "findings", "marked", "export_error", "session"):
            if sb and sb.get(field) != sx.get(field):
                lines.append(f"{name}: BOOK PATH {field} is {sb.get(field)}, the saved trial's is {sx.get(field)}")
        if sb and (sb.get("deck") or {}).get("sha") != (sx.get("deck") or {}).get("sha"):
            lines.append(f"{name}: BOOK PATH deck differs from the saved trial's deck")
        dx, dy = sx.get("deck"), sy.get("deck")
        if (dx or {}).get("sha") != (dy or {}).get("sha"):
            slides = [i + 1 for i, (p, q) in enumerate(zip((dx or {}).get("per_slide", []), (dy or {}).get("per_slide", []))) if p != q]
            lines.append(f"{name}: DECK differs: slides {(dx or {}).get('slides')} -> {(dy or {}).get('slides')}, changed slides {slides[:limit]}")
    return lines


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("action", choices=["save", "check"])
    ap.add_argument("label")
    ap.add_argument("--books", default="", help="only the books whose file name contains this")
    ap.add_argument("--dir", default="", help="the folder of databooks (default: the repo root; the company machine's is databooks)")
    args = ap.parse_args(argv)
    global BOOKS_DIR
    BOOKS_DIR = Path(args.dir).resolve() if args.dir else ROOT
    path = OUT / f"{args.label}.json"
    print(f"BASELINE {args.action} {args.label}")
    saved_snapshot = json.loads(path.read_text(encoding="utf-8")) if args.action == "check" and path.exists() else None
    snap = json.loads(json.dumps(snapshot(args.books, saved_snapshot), default=str))      # as a saved one reads back: lists, not tuples
    if args.action == "save":
        OUT.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(snap, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"saved {path.relative_to(ROOT)}")
    else:
        if not path.exists():
            print(f"no snapshot {path}")
            return 2
        was = json.loads(path.read_text(encoding="utf-8"))
        if args.books:
            was["books"] = {k: v for k, v in was["books"].items() if args.books in k}
        lines = diff(was, snap)
        print(f"snapshot code {was['code']} ({was['taken']})  now {snap['code']}")
        print("IDENTICAL" if not lines else f"{len(lines)} DIFFERENCE LINE(S)")
        for ln in lines:
            print("  " + ln)
        if lines:
            return 1
    for name, b in snap["books"].items():
        s = b["scripted"]
        print(f"  {name[:30]:30s} contracts {b['contracts']['n']:3d} {b['contracts']['sha'][:8]}  sentences {b['generated']['claims']:5d} "
              f"{b['generated']['sha'][:8]}  paid sessions {b['sessions']['sessions']:2d}/{b['sessions']['claims']:4d} {b['sessions']['sha'][:8]}  scripted {s.get('ok')}/{s.get('accounts')} "
              f"deck {((s.get('deck') or {}).get('slides'))} slides {((s.get('deck') or {}).get('sha') or '-')[:8]}")
    return 0


if __name__ == "__main__":
    import os
    if os.environ.get("PYTHONHASHSEED") != "0":
        # the deck's summary band takes its accounts in an order that follows the process's string hashing: pin it, or two runs of
        # the same code differ (found 2026-11-18); a snapshot compares what the code does, not which seed it drew.
        # Not os.execve: on Windows that starts the new process and ends this one at once, so the shell prompt came back
        # and the output of the real run arrived after it (found 2026-11-19 on the company machine). Wait for the child.
        import subprocess
        raise SystemExit(subprocess.run([sys.executable] + sys.argv, env=dict(os.environ, PYTHONHASHSEED="0")).returncode)
    raise SystemExit(main())
