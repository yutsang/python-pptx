"""One comparison across a portfolio folder: one entity per portfolio, each read against its own Financials sheet in that
portfolio's roll-up (主表), as the UI's portfolio flow reads it, and each run with the same arms by entity_trial.py. The
offline gate runs once for all of them.

    python ad-hoc/workbench/portfolio_trial.py --dir databooks                                       free: what would run
    python ad-hoc/workbench/portfolio_trial.py --dir databooks --arms checked,pro --max-cost-per-book 12 \\
        --flash <provider::model> --pro <provider::model> --i-accept-paid-calls                      PAID

Which entity: of those whose sheet the roll-up's matcher finds (inspect_databook.resolve_auto_rollup_targets, the one the UI
uses), the file with the most visible sheets (the most accounts to write), ties by name. --entities picks by name instead.
Hidden sheets are neither counted nor read. Each book writes its own trial files and prints its own DIGEST.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ROOT))


def survey(folder: str, entities: Optional[List[str]] = None) -> List[Dict[str, Any]]:
    """Per portfolio: the roll-up, its visible Financials-like sheets, every entity with the sheet it resolves to, and the pick."""
    import inspect_databook as I
    from entity_trial import _sha
    from fdd_utils.workbook.visibility import visible_sheet_names
    files = sorted(Path(folder).glob("*.xlsx"))
    groups = I.group_portfolio_files(files)
    targets = I.resolve_auto_rollup_targets(files)
    out = []
    for name, group in sorted(groups.items(), key=lambda kv: len(kv[0])):          # I, II, III
        rollup = group.get("rollup")
        rows = []
        for entity, path in sorted(group["entities"].items()):
            visible = visible_sheet_names(str(path)) or []
            rows.append({"entity": entity, "path": path, "sheet": (targets.get(path) or (None, None))[1], "sheets": len(visible),
                         "sha": _sha(str(path))[:8]})
        fin = [s for s in (visible_sheet_names(str(rollup)) or []) if "financ" in s.lower()] if rollup else []
        chosen = [r for r in rows if r["entity"] in entities] if entities else \
            sorted([r for r in rows if r["sheet"]], key=lambda r: (-r["sheets"], r["entity"]))[:1]
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
                       f"Financials {r['sheet'] or 'NOT FOUND in the roll-up'}")
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
    args = ap.parse_args(argv)
    import agent_acceptance as A
    A._utf8()
    found = survey(args.dir, [e.strip() for e in args.entities.split(",") if e.strip()] or None)
    print("\n".join(show(found)), flush=True)
    if not args.i_accept_paid_calls:
        print("free listing only; add --i-accept-paid-calls to run the books above")
        return 0
    return run(found, args)


if __name__ == "__main__":
    sys.exit(main())
