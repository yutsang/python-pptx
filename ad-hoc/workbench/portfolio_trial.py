"""One comparison across a portfolio folder: one entity per portfolio, each read against its own Financials sheet in that
portfolio's roll-up (主表), as the UI's portfolio flow reads it, and each run with the same arms by entity_trial.py. The
offline gate runs once for all of them.

    python ad-hoc/workbench/portfolio_trial.py --dir databooks                                       free: what would run
    python ad-hoc/workbench/portfolio_trial.py --dir databooks --arms checked,pro --max-cost-per-book 12 \\
        --flash <provider::model> --pro <provider::model> --i-accept-paid-calls                      PAID

    python ad-hoc/workbench/portfolio_trial.py --dir databooks --audit [--where] --entities A,B     free: date audit and gap count of each book, read as the trial reads it;
                                                                                                     --where: what the flagged months are (the row's name beside each)
    python ad-hoc/workbench/portfolio_trial.py --dir databooks --blind --entities A,B                free: a review file with the sentences and the workbook text they cite and nothing
                                                                                                     the verifier said (blind_review.py), from each book's newest paid trial
    python ad-hoc/workbench/portfolio_trial.py --dir databooks --recheck                             free: the verifier as it is
                                                                                                     now, over those runs' drafts

--untried picks only entities no paid trial has run on: the books a frozen version is accepted on.

Which entity: of those whose sheet the roll-up's matcher finds (inspect_databook.resolve_auto_rollup_targets, the one the UI
uses), the file with the most visible sheets (the most accounts to write), ties by name. --entities picks by name instead.
Hidden sheets are neither counted nor read. Each book writes its own trial files and prints its own DIGEST.
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ROOT))


def paid_trials(out_dir: Optional[Path] = None) -> List[Any]:
    """(stamp, report) of every paid trial written to out_dir, oldest first."""
    import json
    import re
    import entity_trial as E
    found = []
    for f in Path(out_dir or E.OUT_DIR).glob("trial_*.json"):
        m = re.fullmatch(r"trial_(\d{8}_\d{6})\.json", f.name)
        if m:
            try:
                data = json.loads(f.read_text(encoding="utf-8"))
            except ValueError:
                continue
            if not data.get("scripted"):
                found.append((m.group(1), data))
    return sorted(found, key=lambda t: t[0])


def survey(folder: str, entities: Optional[List[str]] = None, untried: bool = False,
           out_dir: Optional[Path] = None) -> List[Dict[str, Any]]:
    """Per portfolio: the roll-up, its visible Financials-like sheets, every entity with the sheet it resolves to, and the pick.
    With untried, only entities no paid trial has run on are picked: the books a frozen version is accepted on."""
    import inspect_databook as I
    from entity_trial import _sha
    from fdd_utils.workbook.visibility import visible_sheet_names
    done = [str(t[1].get("databook_sha256") or "") for t in paid_trials(out_dir)]
    files = sorted(Path(folder).glob("*.xlsx"))
    groups = I.group_portfolio_files(files)
    targets = I.resolve_auto_rollup_targets(files)
    out = []
    for name, group in sorted(groups.items(), key=lambda kv: len(kv[0])):          # I, II, III
        rollup = group.get("rollup")
        rows = []
        for entity, path in sorted(group["entities"].items()):
            visible = visible_sheet_names(str(path)) or []
            sha = _sha(str(path))[:8]
            rows.append({"entity": entity, "path": path, "sheet": (targets.get(path) or (None, None))[1], "sheets": len(visible),
                         "sha": sha, "tried": any(d.startswith(sha) for d in done)})
        fin = [s for s in (visible_sheet_names(str(rollup)) or []) if "financ" in s.lower()] if rollup else []
        chosen = [r for r in rows if r["entity"] in entities] if entities else \
            sorted([r for r in rows if r["sheet"] and not (untried and r["tried"])], key=lambda r: (-r["sheets"], r["entity"]))[:1]
        out.append({"portfolio": name, "rollup": rollup, "financials": fin, "entities": rows, "pick": chosen})
    return out


def show(found: List[Dict[str, Any]]) -> List[str]:
    out = []
    for g in found:
        out.append(f"Portfolio {g['portfolio']}: roll-up {g['rollup'].name if g['rollup'] else 'NONE'}; "
                   f"Financials sheets {', '.join(g['financials']) or 'none'}")
        picked = {r["entity"] for r in g["pick"]}
        for r in g["entities"]:
            out.append(f"  {'*' if r['entity'] in picked else ' '} {r['entity']}: {r['sheets']} visible sheets, sha {r['sha']}, "
                       f"Financials {r['sheet'] or 'NOT FOUND in the roll-up'}" + ("; tried before" if r.get("tried") else ""))
    picks = [r for g in found for r in g["pick"]]
    out.append(f"will run {len(picks)} book(s): " + (", ".join(r["entity"] for r in picks) or "none"))
    return out


def run(found: List[Dict[str, Any]], args: argparse.Namespace, trial: Optional[Callable[[List[str]], int]] = None,
        gate: Optional[Callable[[], str]] = None) -> int:
    import entity_trial as E
    trial, gate = trial or E.main, gate or E._gate
    problem = gate()
    if problem:
        print("NOT STARTED (nothing was sent):", problem)
        return 2
    codes = []
    for g in found:
        for r in g["pick"]:
            print(f"\n===== Portfolio {g['portfolio']}: {r['entity']} =====", flush=True)
            argv = ["run", "--databook", str(r["path"]), "--arms", args.arms, "--expect-sha", r["sha"],
                    "--max-cost", str(args.max_cost_per_book), "--overview", args.overview, "--skip-gate", "--i-accept-paid-calls"]
            if r["sheet"]:
                argv += ["--financials-from", str(g["rollup"]), "--financials-sheet", r["sheet"]]
            for flag, value in (("--flash", args.flash), ("--pro", args.pro)):
                if value:
                    argv += [flag, value]
            codes.append(trial(argv))
    print(f"\nPORTFOLIO TRIAL: {len(codes)} book(s), exit codes {codes}")
    return 0 if codes and all(c == 0 for c in codes) else 1


def recheck_all(found: List[Dict[str, Any]], out_dir: Optional[Path] = None, check: Optional[Callable[..., Any]] = None) -> List[str]:
    """For each pick, the newest paid trial on that workbook (by its sha), and every arm's saved drafts in it: the marks the
    verifier gives now against those it gave then (recheck_session.py). No model call."""
    import entity_trial as E
    import recheck_session as RS
    out_dir, check = out_dir or E.OUT_DIR, check or RS.recheck
    trials = paid_trials(out_dir)
    out = []
    for g in found:
        for r in g["pick"]:
            mine = sorted((t for t in trials if str(t[1].get("databook_sha256") or "").startswith(r["sha"])), key=lambda t: t[0])
            if not mine:
                out.append(f"===== {r['entity']}: no paid trial of this workbook in {out_dir}")
                continue
            stamp, data = mine[-1]
            for arm in data.get("arms") or {}:
                session = Path(out_dir) / f"trial_{arm}_{stamp}.agent-session.json"
                out.append(f"===== {r['entity']} {arm} (trial {stamp})")
                out.extend(RS.report(check(str(r["path"]), str(session)), 160) if session.exists() else [f"no session file {session.name}"])
    return out


def blind_all(found: List[Dict[str, Any]], out_dir: Optional[Path] = None, build: Optional[Callable[..., str]] = None) -> List[str]:
    """For each pick, the newest paid trial on that workbook and every arm's saved drafts in it, written as a blind review file
    (blind_review.py: sentences and the workbook text they cite, no marks, refusals or flags). No model call. One file per book
    and arm, named by the trial's stamp; the file holds client text and stays local."""
    import entity_trial as E
    import blind_review as BR
    out_dir, build = Path(out_dir or E.OUT_DIR), build or BR.blind
    trials = paid_trials(out_dir)
    out = []
    for g in found:
        for r in g["pick"]:
            mine = sorted((t for t in trials if str(t[1].get("databook_sha256") or "").startswith(r["sha"])), key=lambda t: t[0])
            if not mine:
                out.append(f"===== {r['entity']}: no paid trial of this workbook in {out_dir}")
                continue
            stamp, data = mine[-1]
            for arm in data.get("arms") or {}:
                session = out_dir / f"trial_{arm}_{stamp}.agent-session.json"
                if not session.exists():
                    out.append(f"===== {r['entity']} {arm}: no session file {session.name}")
                    continue
                target = out_dir / f"trial_blind_{arm}_{stamp}.md"
                text = build(str(r["path"]), str(session))
                target.write_text(text, encoding="utf-8")
                out.append(f"===== {r['entity']} {arm} (trial {stamp}): {text.count(chr(10) + '## ')} account(s) -> {target}")
    return out


def packet_all(found: List[Dict[str, Any]], out_dir: Optional[Path] = None, build: Optional[Callable[..., str]] = None,
               every: int = 3, everything: bool = False) -> List[str]:
    """One compact review packet over the newest paid trial of each pick (blind_review.packet): the sentences of every
    `every`-th account and the facts each cites, no marks. Written to review_packet.txt in the trial folder, a section per book
    headed 'BOOK|<trial stamp>' (the stamp score_review.py takes). No model call; the file holds client text."""
    import entity_trial as E
    import blind_review as BR
    out_dir, build = Path(out_dir or E.OUT_DIR), build or BR.packet
    trials = paid_trials(out_dir)
    parts, out = [], []
    for g in found:
        for r in g["pick"]:
            mine = sorted((t for t in trials if str(t[1].get("databook_sha256") or "").startswith(r["sha"])), key=lambda t: t[0])
            if not mine:
                out.append(f"===== {r['entity']}: no paid trial of this workbook in {out_dir}")
                continue
            stamp, data = mine[-1]
            for arm in data.get("arms") or {}:
                session = out_dir / f"trial_{arm}_{stamp}.agent-session.json"
                if not session.exists():
                    out.append(f"===== {r['entity']} {arm}: no session file {session.name}")
                    continue
                body = build(str(r["path"]), str(session), every=every, everything=everything)
                parts.append(f"BOOK|{stamp}\n{body}")
                out.append(f"===== {r['entity']} {arm} (trial {stamp}): {body.count(chr(10) + '### ') + body.startswith('### ')} account(s), {len(body):,} characters")
    if parts:
        target = out_dir / "review_packet.txt"
        target.write_text("\n".join(parts), encoding="utf-8")
        out.append(f"written: {target} ({sum(len(x) for x in parts):,} characters in all)")
    return out


def flagged_months(lines: List[str]) -> List[Any]:
    """[(sheets, months)] from an audit's lines: the sheets of a 'N sheet(s) to check' group and the months it says are written
    often there and used by no account."""
    out, names = [], []
    for line in lines:
        head = re.match(r"\d+ sheet\(s\) to check: (.+?)(?: \(\+\d+ more\))?$", line)
        if head:
            names = [n.strip() for n in head.group(1).split(", ") if n.strip()]
        elif "months written 2+ times in header rows that no account uses" in line and names:
            out.append((names, re.findall(r"\d{4}-\d{2}", line)))
    return out


def audit_all(found: List[Dict[str, Any]], load: Optional[Callable[..., Any]] = None, audit: Optional[Callable[..., List[str]]] = None,
              gaps: Optional[Callable[..., List[str]]] = None, where: Optional[Callable[..., List[str]]] = None) -> List[str]:
    """Free, no model call: each pick read exactly as the paid trial reads it (same workbook, same Financials sheet), then the
    date audit (period_audit.py: does every account's period label agree with the dates written on its own sheet) and how many
    accounts have something their evidence does not say (gap_probe.py). The paid run on a book whose dates are mislabelled is
    wasted: the 2026-11-10 acceptance lost 13 of 16 confirmed errors to one such cluster."""
    import gap_probe as GP
    import period_audit as PA
    load = load or PA.load_session
    audit = audit or (lambda session, params: PA.audit(session, params))
    gaps = gaps or (lambda session: GP.gaps(session, 56))
    out = []
    for g in found:
        for r in g["pick"]:
            out.append(f"===== {r['entity']} (Portfolio {g['portfolio']})")
            args = argparse.Namespace(session="", dir="", databook=str(r["path"]), entity="", sheet="", verbose=False, probe="",
                                      financials_from=str(g["rollup"]) if r["sheet"] else "", financials_sheet=r["sheet"] or "")
            try:
                session, params = load(args)
                lines = audit(session, params)
                out.extend(lines)
                out.append(gaps(session)[0])
                for names, months in (flagged_months(lines) if where else []):
                    for name in names:
                        out.extend(where(params["workbook"], name, months))
            except Exception as exc:                                              # noqa: BLE001
                out.append(f"could not be read: {type(exc).__name__}: {exc}")
    return out


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--dir", required=True, help="the folder of portfolio databooks")
    ap.add_argument("--entities", default="", help="comma list of entity names to run instead of one pick per portfolio")
    ap.add_argument("--arms", default="checked,pro")
    ap.add_argument("--max-cost-per-book", type=float, default=12.0, help="yuan; one ceiling over every arm of one book")
    ap.add_argument("--overview", default="none", choices=("both", "pro", "flash", "none"))
    ap.add_argument("--flash", default="")
    ap.add_argument("--pro", default="")
    ap.add_argument("--i-accept-paid-calls", action="store_true")
    ap.add_argument("--recheck", action="store_true", help="free: re-verify the newest paid trial of each pick with the code as it is now")
    ap.add_argument("--audit", action="store_true", help="free: before a paid run, the date audit and the gap count of each pick, read as the trial reads it")
    ap.add_argument("--blind", action="store_true", help="free: write a blind review file (no marks) from the newest paid trial of each pick")
    ap.add_argument("--packet", action="store_true", help="free: one compact review packet (sentences and the facts they cite, no marks) from the newest paid trial of each pick")
    ap.add_argument("--every", type=int, default=3, help="with --packet: every n-th account in the blind file's numbering (default 3)")
    ap.add_argument("--all-accounts", action="store_true", help="with --packet: every account")
    ap.add_argument("--where", action="store_true", help="with --audit: for each sheet it flags, the rows that write the months no account uses, with the row's name")
    ap.add_argument("--untried", action="store_true", help="pick only entities no paid trial has run on (acceptance books)")
    args = ap.parse_args(argv)
    import agent_acceptance as A
    A._utf8()
    found = survey(args.dir, [e.strip() for e in args.entities.split(",") if e.strip()] or None, untried=args.untried)
    print("\n".join(show(found)), flush=True)
    if args.recheck:
        print("\n".join(recheck_all(found)))
        return 0
    if args.packet:
        print("\n".join(packet_all(found, every=args.every, everything=args.all_accounts)))
        return 0
    if args.blind:
        print("\n".join(blind_all(found)))
        return 0
    if args.audit:
        import warnings
        warnings.filterwarnings("ignore")
        import period_audit as PA
        print("\n".join(audit_all(found, where=PA.where_months if args.where else None)))
        return 0
    if not args.i_accept_paid_calls:
        print("free listing only; add --i-accept-paid-calls to run the books above")
        return 0
    return run(found, args)


if __name__ == "__main__":
    sys.exit(main())
