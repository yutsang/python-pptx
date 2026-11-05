#!/usr/bin/env python3
"""Pack what a baseline comparison needs, or say exactly what is missing. Free: no model call, nothing leaves the machine.

A baseline is a databook, the archived run that describes THAT databook, the deck that run produced (if one was accepted),
and the usage-ledger lines of the run. This tool finds them, proves they belong together, and copies them into one folder
under a git-ignored path. It never pairs by file name: a run belongs to a databook when the run's identity names the
databook's sha256, or (a run with no identity file) when every account's data fingerprint matches the databook as it reads now.
It never accepts a databook that was rebuilt from a text export, a copy that was saved again, or a demo copy in place of the
original: those have other bytes and so no run names them.

    # what would be packed, and what is missing (writes nothing)
    python ad-hoc/workbench/baseline_pack.py --databook for_test/<databook>.xlsx

    # the same, and write the pack
    python ad-hoc/workbench/baseline_pack.py --databook for_test/<databook>.xlsx --out for_test/baseline-pack

    # check a pack again later (hashes, secrets, that the databook is still the one the run names)
    python ad-hoc/workbench/baseline_pack.py --verify for_test/baseline-pack

Exit status: 0 complete, 1 usable with gaps (the gaps are listed), 2 not a baseline (nothing is written; the reasons are listed).

What goes in the pack, and what never does
  in      the databook as it is; the run's results, identity, per-account fingerprints, evidence, digest, section summaries,
          selection, summary; the deck and its lineage / delivery sidecars; the ledger lines of the run (numbers and model names)
  out     config.yml or anything like it; processing.log (error lines can carry gateway addresses); the stage prompts and replies
          (data.yml, audit.jsonl) unless --include-prompts; the ledger's error text; any other run, file or entity's ledger line
  checked every file that is written is searched for API keys, bearer tokens, the values of the key / address / token fields of
          the config in use, and addresses; a hit refuses the pack and names the file and the kind of hit, never the value

What is printed leaves out file names, entity and account names (use --names to show them); the manifest keeps them, for local use.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import time
from collections import Counter
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

PACK_VERSION = 1
MANIFEST = "MANIFEST.json"
RUN_FILES = ("results.yml", "run_identity.json", "checkpoint.jsonl", "digest.json", "section_summaries.json", "selection.json", "summary.txt")
PROMPT_FILES = ("data.yml", "audit.jsonl")                     # stage prompts and replies: client text, large; only on request
DECK_SIDECARS = (".lineage.json", ".delivery.json")
LEDGER_FIELDS = ("time", "stage", "client_agent", "provider", "model", "ok", "duration_s", "prompt_tokens", "completion_tokens", "total_tokens",
                 "cached_prompt_tokens", "reasoning_tokens", "estimated_prompt_tokens", "estimated_output_tokens", "token_source",
                 "saved_prompt_tokens", "saved_completion_tokens", "account", "session", "file", "entity")
#: how long after the last file of a run was written its executive-summary and mapping calls may still be in the ledger
LEDGER_SLACK_S = 900

SECRET_PATTERNS = {
    "an OpenAI-style key": re.compile(r"\bsk-[A-Za-z0-9_\-]{16,}"),
    "a bearer token": re.compile(r"\bBearer\s+[A-Za-z0-9._\-]{16,}"),
    "an api key assignment": re.compile(r"(?i)\bapi[_-]?key\"?\s*[:=]\s*\"?[A-Za-z0-9._\-]{8,}"),
    "a secret / token / password assignment": re.compile(r"(?i)\b(?:secret|password|passwd|auth[_-]?token|access[_-]?token)\"?\s*[:=]\s*\"?[^\s\",}{]{8,}"),
    "an https address": re.compile(r"https?://[A-Za-z0-9.\-]+\.[A-Za-z]{2,}(?::\d+)?/[^\s\"']*"),
}
SECRET_FIELD = re.compile(r"(?i)(key|secret|token|password|passwd|charge_code|api_base|base_url|endpoint|host)")


# ---------------------------------------------------------------------------------------------------------------------------
# small helpers
# ---------------------------------------------------------------------------------------------------------------------------

def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read_json(path: Path) -> Optional[Dict[str, Any]]:
    try:
        with open(path, encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else None
    except Exception:                                           # noqa: BLE001
        return None


def read_jsonl(path: Path) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    try:
        with open(path, encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if line:
                    try:
                        obj = json.loads(line)
                        if isinstance(obj, dict):
                            rows.append(obj)
                    except json.JSONDecodeError:
                        continue
    except OSError:
        pass
    return rows


def parse_run_time(folder: Path) -> Optional[float]:
    m = re.match(r"run_(\d{8})_(\d{6})", folder.name)
    if not m:
        return None
    try:
        return time.mktime(time.strptime(m.group(1) + m.group(2), "%Y%m%d%H%M%S"))
    except ValueError:
        return None


def run_window(folder: Path) -> Tuple[float, float]:
    """(start, end) in epoch seconds: the folder is created when the run starts, and its newest file is its end."""
    start = parse_run_time(folder) or min((p.stat().st_mtime for p in folder.rglob("*") if p.is_file()), default=0.0)
    end = max((p.stat().st_mtime for p in folder.rglob("*") if p.is_file()), default=start)
    return start, max(end, start)


def iso_to_epoch(text: Any) -> Optional[float]:
    try:
        return time.mktime(time.strptime(str(text)[:19], "%Y-%m-%dT%H:%M:%S"))
    except ValueError:
        return None


def rebuilt_marker(path: Path) -> Optional[Dict[str, Any]]:
    """databook_dump.py marks a workbook it rebuilt from a text dump: a `.rebuilt.json` beside it and a line inside its properties.
    A marked file is a rebuild whatever it is called, and its bytes are not the original's."""
    side = Path(str(path) + ".rebuilt.json")
    if side.exists():
        data = read_json(side)
        return data or {"rebuilt_from_a_text_dump": True}
    try:
        import zipfile
        with zipfile.ZipFile(path) as z:
            if "REBUILT FROM A TEXT DUMP" in z.read("docProps/core.xml").decode("utf-8", errors="ignore"):
                return {"rebuilt_from_a_text_dump": True}
    except Exception:                                           # noqa: BLE001
        pass
    return None


_DERIVED_NAME = re.compile(r"(?i)(rebuil[dt]|demo|anonymi[sz]ed|redacted|sample|synthetic|\bcopy\b|副本)")


def derived_name_hint(name: str) -> bool:
    """The file's own name says it is not the client's original. Weak: the bytes alone cannot tell (a rebuild saved by Excel carries Excel's
    application tag), so this only ever adds a warning; a file with an innocent name is not thereby the original."""
    return bool(_DERIVED_NAME.search(name))


def check_out_path(path: Path) -> None:
    """A pack holds a client's figures: refuse a location git would track."""
    target = path.expanduser().resolve()
    try:
        target.relative_to(ROOT)
    except ValueError:
        return
    probe = target / "MANIFEST.json"
    under_for_test = str(target).startswith(str(ROOT / "for_test"))
    try:
        # 0 = ignored, 1 = not ignored; any other code (128 in a copied or downloaded tree that is not a git clone) is no answer from git,
        # and then only the documented ignored folder, for_test/, is accepted.
        code = subprocess.run(["git", "check-ignore", "-q", str(probe)], cwd=str(ROOT), capture_output=True, timeout=10).returncode
        ignored = (code == 0) if code in (0, 1) else under_for_test
    except Exception:                                           # noqa: BLE001
        ignored = under_for_test
    if not ignored:
        raise SystemExit(f"{path} is inside the repo and not git-ignored. A pack holds a client's figures: use for_test/<name> (ignored) or a path outside the repo.")


# ---------------------------------------------------------------------------------------------------------------------------
# runs
# ---------------------------------------------------------------------------------------------------------------------------

def list_runs(logs_dir: Path) -> List[Path]:
    return sorted(p for p in logs_dir.glob("run_*") if (p / "results.yml").exists())


def load_results(folder: Path) -> Dict[str, Any]:
    import yaml
    with open(folder / "results.yml", encoding="utf-8") as fh:
        return yaml.safe_load(fh) or {}


def account_texts(results: Dict[str, Any]) -> List[str]:
    from fdd_utils.financial_common import get_pipeline_result_text
    return [str(get_pipeline_result_text(v) or "") for k, v in results.items() if not str(k).startswith("__") and isinstance(v, dict)]


def run_kind(folder: Path, ledger_in_window: List[Dict[str, Any]]) -> str:
    """'real' when the ledger has provider-reported calls in the run's window; 'stub-like' when three or more accounts carry the very same
    commentary (a canned model answers every account alike); otherwise 'unknown'. A run's own 'calls ok' line does not tell them apart:
    a stub counts as a call that succeeded."""
    if any(r.get("ok") and r.get("provider") and r.get("provider") != "heuristic" and r.get("prompt_tokens") for r in ledger_in_window):
        return "real"
    try:
        texts = [t for t in account_texts(load_results(folder)) if len(t) > 20]
    except Exception:                                           # noqa: BLE001
        return "unknown"
    if len(texts) >= 3 and len(set(texts)) == 1:
        return "stub-like"
    return "unknown"


def ledger_rows(path: Path, file_names: Iterable[str], start: float, end: float) -> List[Dict[str, Any]]:
    """The ledger's lines for this databook inside this run's window, and only the whitelisted fields; the error text is cut to its type.
    `file_names` is every name the databook went by: the one on disk and the one the run's identity recorded (an upload keeps its own)."""
    names = {n for n in file_names if n}
    out = []
    for row in read_jsonl(path):
        when = iso_to_epoch(row.get("time"))
        if when is None or not (start - 60 <= when <= end + LEDGER_SLACK_S):
            continue
        if row.get("file") not in names:
            continue                                          # another file's line, or a line no upload or CLI run claims: not shown to belong to this one
        clean = {k: row.get(k) for k in LEDGER_FIELDS if k in row}
        if row.get("error"):
            clean["error_type"] = str(row["error"]).split(":", 1)[0][:60]
        out.append(clean)
    return out


def summarise_ledger(rows: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Provider-reported tokens, estimates, provider cache and response-cache hits, kept apart (never added together)."""
    sent = [r for r in rows if r.get("token_source") != "response_cache"]
    hits = [r for r in rows if r.get("token_source") == "response_cache"]
    reported = [r for r in sent if r.get("ok") and r.get("prompt_tokens") is not None and r.get("token_source") not in ("estimate", "estimated")]
    by_stage: Dict[str, Dict[str, int]] = {}
    for r in sent:
        row = by_stage.setdefault(str(r.get("stage") or "?"), {"calls": 0, "failed": 0, "prompt": 0, "completion": 0})
        row["calls"] += 1
        row["failed"] += 0 if r.get("ok") else 1
        if r in reported:
            row["prompt"] += int(r.get("prompt_tokens") or 0)
            row["completion"] += int(r.get("completion_tokens") or 0)
    return {
        "rows": len(rows), "calls_sent": len(sent), "calls_failed": sum(not r.get("ok") for r in sent),
        "provider_reported": {"calls": len(reported), "prompt": sum(int(r.get("prompt_tokens") or 0) for r in reported),
                              "completion": sum(int(r.get("completion_tokens") or 0) for r in reported),
                              "cached_prompt": sum(int(r.get("cached_prompt_tokens") or 0) for r in reported),
                              "reasoning": sum(int(r.get("reasoning_tokens") or 0) for r in reported)},
        "calls_without_provider_usage": len(sent) - len(reported),
        "estimated_prompt_all_calls": sum(int(r.get("estimated_prompt_tokens") or 0) for r in sent),
        "response_cache_hits": len(hits),
        "response_cache_saved": {"prompt": sum(int(r.get("saved_prompt_tokens") or 0) for r in hits),
                                 "completion": sum(int(r.get("saved_completion_tokens") or 0) for r in hits)},
        "models": dict(Counter(str(r.get("model") or "?") for r in sent)), "by_stage": by_stage,
        "first": min((str(r.get("time")) for r in rows), default=None), "last": max((str(r.get("time")) for r in rows), default=None),
        "unmeasured": "retries made inside the provider SDK or gateway never reach the ledger",
    }


def total_provider_tokens(summary: Dict[str, Any]) -> int:
    p = summary.get("provider_reported") or {}
    return int(p.get("prompt") or 0) + int(p.get("completion") or 0)


# ---------------------------------------------------------------------------------------------------------------------------
# proof that a run describes a databook
# ---------------------------------------------------------------------------------------------------------------------------

def identity_matches(folder: Path, databook_sha: str, financials_sha: Optional[str]) -> Tuple[bool, Optional[Dict[str, Any]], List[str]]:
    identity = read_json(folder / "run_identity.json")
    if not identity:
        return False, None, ["the run has no identity file"]
    why = []
    if identity.get("workbook_sha256") != databook_sha:
        why.append("the run was made from a databook with other bytes")
    if identity.get("financials_from") and identity.get("financials_sha256") != financials_sha:
        why.append("the run's Financials file is not the one given (or none was given)")
    return not why, identity, why


def deep_check(databook: str, folder: Path, identity: Optional[Dict[str, Any]], entity: str = "", sheet: Optional[str] = None,
               financials_from: Optional[str] = None, financials_sheet: Optional[str] = None) -> Dict[str, Any]:
    """The agent path's own proof (session.check_base_run): the identity, and every account's data fingerprint against the workbook as it
    reads now. Reads the workbook once; no model."""
    try:
        from fdd_utils.agent.session import SourceSession
        from fdd_utils.run_identity import load_run
        ident = identity or {}
        session = SourceSession.load(databook, entity=entity or ident.get("entity") or "", sheet=sheet or ident.get("sheet") or None,
                                     financials_from=financials_from or None, financials_sheet=financials_sheet or None, base_run=str(folder))
        check = dict(session.base_run_check or {})
        check["accounts_in_workbook"] = len(session.contracts)
        return check
    except Exception as exc:                                    # noqa: BLE001
        return {"status": "error", "reasons": [f"the proof could not be run: {type(exc).__name__}: {str(exc)[:200]}"], "accounts": {}}


# ---------------------------------------------------------------------------------------------------------------------------
# secrets
# ---------------------------------------------------------------------------------------------------------------------------

def config_secret_values(config: Optional[Dict[str, Any]]) -> List[str]:
    """Values of the config's key / token / address fields, so the exact strings in use are searched for, not only things that look like keys."""
    found: List[str] = []

    def walk(node: Any, field: str = "") -> None:
        if isinstance(node, dict):
            for k, v in node.items():
                walk(v, str(k))
        elif isinstance(node, list):
            for v in node:
                walk(v, field)
        elif isinstance(node, str) and SECRET_FIELD.search(field) and len(node.strip()) >= 8 and node.strip().lower() not in ("not-needed", "bearer", "workbench"):
            found.append(node.strip())
    walk(config or {})
    return sorted(set(found))


def load_config() -> Dict[str, Any]:
    try:
        from fdd_utils.ai import FDDConfig
        return FDDConfig().config or {}
    except Exception:                                           # noqa: BLE001
        return {}


#: set by --allow-urls after the person has looked at what the address hits are (a client's own text can hold a link)
ALLOW_ADDRESSES = False


def scan_text_for_secrets(text: str, values: List[str]) -> List[str]:
    hits = [kind for kind, pattern in SECRET_PATTERNS.items() if pattern.search(text) and not (ALLOW_ADDRESSES and kind == "an https address")]
    if any(v in text for v in values):
        hits.append("a value taken from the config in use")
    return hits


def scan_files(paths: Iterable[Path], values: List[str], root: Optional[Path] = None) -> List[Tuple[str, str]]:
    """[(relative path, kind of hit)] over the text files; binaries (the databook, the deck) are not searched: they are the client's own file."""
    out = []
    for path in paths:
        if path.suffix.lower() in (".xlsx", ".xlsm", ".pptx", ".png", ".jpg", ".zip"):
            continue
        try:
            text = path.read_text(encoding="utf-8", errors="ignore")
        except OSError:
            continue
        for kind in scan_text_for_secrets(text, values):
            out.append((path.relative_to(root).as_posix() if root else path.name, kind))
    return out


# ---------------------------------------------------------------------------------------------------------------------------
# the plan: what was found, what is missing
# ---------------------------------------------------------------------------------------------------------------------------

def find_decks(databook_sha: str, folders: List[Path]) -> Tuple[List[Path], int]:
    """Decks whose lineage sidecar names this databook's hash, and how many decks were seen without a sidecar that could say (not paired)."""
    paired, unknown = [], 0
    for folder in folders:
        if not folder.is_dir():
            continue
        for deck in sorted(folder.glob("*.pptx")):
            side = read_json(Path(str(deck) + ".lineage.json"))
            sha = ((side or {}).get("binding") or {}).get("workbook_sha256")
            if sha == databook_sha:
                paired.append(deck)
            elif not sha:
                unknown += 1
    return paired, unknown


def build_plan(args: argparse.Namespace, *, config: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    databook = Path(args.databook)
    logs_dir = Path(args.logs_dir)
    plan: Dict[str, Any] = {"pieces": [], "missing": [], "problems": [], "status": "not_a_baseline", "proof": None, "run": None,
                            "alternatives": [], "deck": None, "ledger": None}
    pieces, missing, problems = plan["pieces"], plan["missing"], plan["problems"]

    def need(name: str, why: str, how: str) -> None:
        missing.append({"piece": name, "why_it_matters": why, "how_to_get": how})

    if not databook.is_file():
        problems.append("the databook file was not found")
        return plan
    db_sha = sha256_file(databook)
    fin_sha = sha256_file(Path(args.financials_from)) if getattr(args, "financials_from", None) and Path(args.financials_from).is_file() else None
    marker = rebuilt_marker(databook)
    plan["databook"] = {"path": str(databook), "sha256": db_sha, "bytes": databook.stat().st_size, "name_suggests_a_copy": derived_name_hint(databook.name),
                        "rebuilt": marker}
    #: the hash the runs are searched for under: a rebuild is looked up by the file it was rebuilt from, which only FINDS candidates;
    #: whether one describes the rebuild is the per-account proof's call, never the hash's
    search_sha = (marker or {}).get("source_sha256") or db_sha
    pieces.append({"piece": "databook", "path": str(databook), "required": True})

    runs = list_runs(logs_dir) if logs_dir.is_dir() else []
    chosen: Optional[Path] = None
    identity: Optional[Dict[str, Any]] = None
    if args.run:
        candidate = Path(args.run) if Path(args.run).is_dir() else logs_dir / (args.run if str(args.run).startswith("run_") else f"run_{args.run}")
        if (candidate / "results.yml").exists():
            chosen = candidate
        else:
            problems.append("the run named on the command line has no results.yml")
    else:
        paired = []
        for folder in runs:
            ok, ident, _ = identity_matches(folder, search_sha, fin_sha)
            if ok:
                start, end = run_window(folder)
                rows = ledger_rows(Path(args.ledger), {databook.name, (ident or {}).get("workbook")}, start, end) if Path(args.ledger).exists() else []
                kind = run_kind(folder, rows)
                try:
                    n = len([k for k in load_results(folder) if not str(k).startswith("__")])
                except Exception:                               # noqa: BLE001
                    n = 0
                paired.append({"folder": folder, "identity": ident, "kind": kind, "accounts": n, "time": parse_run_time(folder) or 0.0})
        # a real run first, then the one with the most accounts, then the newest
        order = {"real": 0, "unknown": 1, "stub-like": 2}
        paired.sort(key=lambda p: (order.get(p["kind"], 1), -p["accounts"], -p["time"]))
        plan["alternatives"] = [{"run": p["folder"].name, "kind": p["kind"], "accounts": p["accounts"]} for p in paired[1:]]
        if paired:
            chosen, identity = paired[0]["folder"], paired[0]["identity"]
            if paired[0]["kind"] == "stub-like":
                problems.append("the only run that names this databook looks like a stub run (every account carries the same commentary); "
                                "it is not a model baseline")
        else:
            problems.append(f"no archived run names this databook's hash ({search_sha[:8]}): {len(runs)} run folder(s) were looked at. A run made from a "
                            "different copy of the file (re-saved, rebuilt from a text export, a demo copy) cannot stand in for it")
    if chosen is None:
        need("the archived run for this databook", "the commentary, the evidence and the fingerprints that say which workbook it describes",
             "to see what each archived run can prove about this file: add --explain-runs (writes nothing). On the machine that ran the AI pass: "
             "python ad-hoc/workbench/list_archived_runs.py --healthy --min-accounts 20 , then give --run <RUN_ID>; or copy that run's folder "
             "from fdd_utils/logs/ to this machine's fdd_utils/logs/")
        return plan
    identity = identity or read_json(chosen / "run_identity.json")
    plan["run"] = {"folder": str(chosen), "id": chosen.name}
    ok, _, why = identity_matches(chosen, db_sha, fin_sha)
    # the proof: identity first (cheap), then the agent path's own check against the workbook as it reads now
    if marker and not args.deep:
        why = ["this workbook is a rebuild from a text dump: its bytes differ from the original's, so only the per-account fingerprint proof "
               "(run without --no-deep) can show a run describes it"] + why
        ok = False
    if args.deep:
        proof = deep_check(str(databook), chosen, identity, args.entity or "", args.sheet, args.financials_from, args.financials_sheet)
    else:
        proof = {"status": "verified" if ok else "mismatch", "reasons": why, "accounts": {}, "identity_file": bool(identity), "checked": "identity hash only"}
    if identity and ok:
        proof["kind"] = "identity"
    elif proof.get("status") == "verified":
        proof["kind"] = "fingerprints"
        if identity and not ok:
            proof["note"] = "the workbook's bytes differ from the ones the run's identity names; every account's data fingerprint matches"
    else:
        proof["kind"] = "identity" if identity else "none"
    plan["proof"] = proof
    if proof.get("status") != "verified":
        problems.append(f"the run does not provably describe this databook ({proof.get('status')}): " + "; ".join(proof.get("reasons") or ["no reason given"])[:600])
        need("proof that this run was written from this databook",
             "a run's commentary is a baseline only for the workbook it was written from; a run that merely resembles one (figures that match) is not that",
             "run_identity.json (written by inspect_databook.py --run-ai) or checkpoint.jsonl with a data fingerprint for every account. A run that has "
             "neither cannot be tied to a databook afterwards; make a new run from the original file, or use the run that has the identity file")
        return plan
    for name in RUN_FILES:
        present = (chosen / name).exists()
        pieces.append({"piece": f"run/{name}", "path": str(chosen / name), "required": name == "results.yml", "present": present})
    if (chosen / "evidence").is_dir() and any((chosen / "evidence").iterdir()):
        pieces.append({"piece": "run/evidence", "path": str(chosen / "evidence"), "required": False, "present": True})
    else:
        need("run/evidence/", "the per-account evidence the run's verdicts were checked against (replay_verification.py reads it)",
             "it is written by a run made with the evidence contract; a run from before that has none, and cannot be replayed")
    if not (chosen / "section_summaries.json").exists():
        need("run/section_summaries.json", "the executive summary the deck shipped with (a re-export without it has no summary)",
             "written by every export since the offline re-export landed; an older run has none")
    if not (chosen / "checkpoint.jsonl").exists():
        need("run/checkpoint.jsonl", "the per-account data fingerprints that prove which accounts a later edit of the workbook touched",
             "written during the run; a run that was not checkpointed has none")
    if args.include_prompts:
        for name in PROMPT_FILES:
            if (chosen / name).exists():
                pieces.append({"piece": f"run/{name}", "path": str(chosen / name), "required": False, "present": True, "prompts": True})
    # the deck
    if args.deck:
        decks = [Path(args.deck)] if Path(args.deck).is_file() else []
        if not decks:
            problems.append("the deck named on the command line was not found")
        unknown = 0
    else:
        # --deck-dir replaces the default places to look (this repo's pptx_previews/ and fdd_utils/output/); the databook's folder and the run's are always looked in
        searched = [Path(p) for p in args.deck_dir] if args.deck_dir else [ROOT / "pptx_previews", ROOT / "fdd_utils" / "output"]
        folders = searched + [databook.parent, chosen]
        decks, unknown = find_decks(search_sha, folders)
    plan["decks_seen_without_lineage"] = unknown
    if decks:
        deck = max(decks, key=lambda p: p.stat().st_mtime)
        plan["deck"] = {"path": str(deck), "sha256": sha256_file(deck), "others_paired": len(decks) - 1}
        pieces.append({"piece": "deck", "path": str(deck), "required": False, "present": True})
        for suffix in DECK_SIDECARS:
            if Path(str(deck) + suffix).exists():
                pieces.append({"piece": f"deck{suffix}", "path": str(deck) + suffix, "required": False, "present": True})
            else:
                need(f"deck{suffix}", "the paragraph-level lineage / delivery decision that returned-deck tests read",
                     "it is written beside every deck exported since the lineage landed; for an older deck, export it again with --from-run")
    else:
        need("the accepted PPTX with its lineage", "the returned-deck tests (manual edits, reflow, tables, cross-page accounts) need the deck as it was accepted",
             "give --deck <path>; the deck is found by itself only when its .lineage.json sidecar names this databook's hash"
             + (f" ({unknown} deck(s) were seen with no sidecar that could say)" if unknown else ""))
    # the ledger
    start, end = run_window(chosen)
    ledger_path = Path(args.ledger)
    if ledger_path.exists():
        rows = ledger_rows(ledger_path, {databook.name, (identity or {}).get("workbook")}, start, end)
        summary = summarise_ledger(rows)
        plan["ledger"] = {"rows": rows, "summary": summary, "path": str(ledger_path)}
        if rows:
            pieces.append({"piece": "usage/usage_slice.jsonl", "required": False, "present": True})
        else:
            need("the ledger lines of the run", "the only complete record of what the run cost (the run folder keeps one attempt per account and stage)",
                 "no line of usage.jsonl names this file inside the run's time window: take usage.jsonl from the machine that ran it "
                 "(fdd_utils/logs/usage.jsonl), or the admin usage download for that day")
    else:
        need("fdd_utils/logs/usage.jsonl", "the only complete record of what the run cost; without it the historical token figure cannot be confirmed",
             "copy it from the machine that ran the AI pass (or pass --ledger <path>)")
    start_kind = run_kind(chosen, (plan["ledger"] or {}).get("rows") or [])
    plan["run"]["kind"] = start_kind
    if start_kind == "stub-like":
        problems.append("this run looks like a stub run (every account carries the same commentary): it is not a model baseline")
        return plan
    if start_kind != "real" and (plan["ledger"] or {}).get("rows"):
        need("ledger lines that show a model was called", "a run's own 'calls ok' line counts a stub as a success; only provider-reported ledger lines show a model answered",
             "the ledger lines found for this run report no provider tokens: use the run from a real model pass, or the ledger that holds it")
    if marker or proof.get("kind") == "fingerprints" and identity and not ok:
        need("the original databook bytes", "this workbook is not the file the run was made from (it is a rebuild, or another save of it); a run proved "
             "to describe it only by account fingerprints is a baseline for the data, not for the file",
             "carry the original with text_transfer.py (the exact bytes), and run this tool on that")
    if plan["databook"]["name_suggests_a_copy"]:
        need("the databook as the client sent it", "the file's name says it is a rebuild, a demo or a copy. A run made from it describes that file; "
             "whether its bytes are the client's original cannot be told from the bytes",
             "use the original .xlsx. Do not rebuild a workbook from the selected-tabs text export: it holds values only, so it is another file")
    plan["status"] = "usable_with_gaps" if missing else "complete"
    return plan


# ---------------------------------------------------------------------------------------------------------------------------
# writing and verifying
# ---------------------------------------------------------------------------------------------------------------------------

def collect_files(plan: Dict[str, Any], leave_out_databook: bool = False) -> List[Tuple[Path, str]]:
    """[(source, relative destination)] for everything the plan says is present."""
    out: List[Tuple[Path, str]] = [] if leave_out_databook else [(Path(plan["databook"]["path"]), f"databook/{Path(plan['databook']['path']).name}")]
    run = Path(plan["run"]["folder"])
    for piece in plan["pieces"]:
        name = piece["piece"]
        if name.startswith("run/") and piece.get("present") and name != "run/evidence":
            out.append((Path(piece["path"]), f"run/{run.name}/{name[4:]}"))
        elif name == "run/evidence":
            for p in sorted(Path(piece["path"]).rglob("*")):
                if p.is_file():
                    out.append((p, f"run/{run.name}/evidence/{p.relative_to(piece['path']).as_posix()}"))
        elif name.startswith("deck") and piece.get("present"):
            out.append((Path(piece["path"]), f"deck/{Path(piece['path']).name}"))
    return out


def write_pack(plan: Dict[str, Any], out: Path, *, config: Optional[Dict[str, Any]] = None, leave_out_databook: bool = False) -> Dict[str, Any]:
    files = collect_files(plan, leave_out_databook)
    values = config_secret_values(config if config is not None else load_config())
    # search the text BEFORE anything is written
    hits = scan_files([src for src, _ in files], values)
    ledger_text = "\n".join(json.dumps(r, ensure_ascii=False) for r in (plan["ledger"] or {}).get("rows") or [])
    for kind in scan_text_for_secrets(ledger_text, values):
        hits.append(("usage/usage_slice.jsonl", kind))
    if hits:
        return {"ok": False, "refused": True, "hits": sorted(set(hits))}
    if out.exists() and any(out.iterdir()):
        return {"ok": False, "refused": True, "reason": "the output folder exists and is not empty"}
    out.mkdir(parents=True, exist_ok=True)
    entries = []
    for src, rel in files:
        dest = out / rel
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dest)
        if sha256_file(dest) != sha256_file(src):
            return {"ok": False, "refused": True, "reason": f"a copied file differs from its source ({rel})"}
        entries.append({"path": rel, "sha256": sha256_file(dest), "bytes": dest.stat().st_size})
    rows = (plan["ledger"] or {}).get("rows") or []
    if rows:
        dest = out / "usage" / "usage_slice.jsonl"
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_text("\n".join(json.dumps(r, ensure_ascii=False) for r in rows) + "\n", encoding="utf-8")
        entries.append({"path": "usage/usage_slice.jsonl", "sha256": sha256_file(dest), "bytes": dest.stat().st_size})
    identity = read_json(Path(plan["run"]["folder"]) / "run_identity.json") or {}
    summary = (plan["ledger"] or {}).get("summary")
    manifest = {
        "pack_version": PACK_VERSION, "written_at": time.strftime("%Y-%m-%dT%H:%M:%S"), "status": plan["status"],
        "databook": {"file": Path(plan["databook"]["path"]).name, "sha256": plan["databook"]["sha256"], "bytes": plan["databook"]["bytes"],
                     "packed": not leave_out_databook,
                     "nature": ("the bytes the run's identity names (sha256 equal)" if plan["proof"].get("kind") == "identity"
                                else "matched to the run by per-account data fingerprints (the run has no identity file)")
                               + "; whether these bytes are the client's original cannot be told from the file"
                               + ("; REBUILT from a text dump of a file with sha256 " + str(plan["databook"]["rebuilt"].get("source_sha256") or "?")[:12]
                                  + ": not the original" if plan["databook"].get("rebuilt") else "")
                               + ("; its name suggests a rebuild, demo or copy" if plan["databook"].get("name_suggests_a_copy") else "")},
        "run": {"id": plan["run"]["id"], "kind": plan["run"].get("kind"), "entity": identity.get("entity"), "sheet": identity.get("sheet"),
                "language": identity.get("language"), "model_type": identity.get("model_type"), "model_name": identity.get("model_name"),
                "code_revision": identity.get("code_revision"), "accounts": len(identity.get("accounts") or [])},
        "proof": {k: v for k, v in plan["proof"].items() if k in ("status", "kind", "reasons", "accounts", "identity_file", "checked", "accounts_in_workbook")},
        "deck": plan["deck"], "usage": summary,
        "tokens_provider_reported_total": total_provider_tokens(summary) if summary else None,
        "missing": plan["missing"], "alternatives": plan["alternatives"],
        "not_included": ["config.yml or any copy of it", "processing.log", "stage prompts and replies (data.yml, audit.jsonl) unless --include-prompts",
                         "ledger error text", "any other run, file or entity's ledger line",
                         "a selected-tabs text export: it holds values only, so a workbook rebuilt from it is a different file and cannot stand in for the original"],
        "files": entries,
    }
    (out / MANIFEST).write_text(json.dumps(manifest, ensure_ascii=False, indent=1), encoding="utf-8")
    (out / "README.txt").write_text(
        "A baseline pack: a databook, the archived run that describes it, the accepted deck (if any) and the run's usage-ledger lines.\n"
        "MANIFEST.json lists every file with its sha256, what is missing, and how the run was shown to describe this databook.\n"
        "Check it again at any time:   python ad-hoc/workbench/baseline_pack.py --verify <this folder>\n"
        "It holds a client's figures and commentary. Keep it under a git-ignored path (for_test/) and do not send it on.\n", encoding="utf-8")
    return {"ok": True, "manifest": manifest, "files": len(entries) + 2}


def verify_pack(pack: Path, *, config: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """Hashes, no stray files, no secrets, and the databook still the one the run's identity names. No extraction, no model."""
    problems: List[str] = []
    manifest = read_json(pack / MANIFEST)
    if not manifest:
        return {"ok": False, "problems": ["no MANIFEST.json"]}
    listed = {e["path"]: e for e in manifest.get("files") or []}
    for rel, entry in listed.items():
        path = pack / rel
        if not path.exists():
            problems.append(f"listed but missing: {rel}")
        elif sha256_file(path) != entry["sha256"]:
            problems.append(f"changed since it was packed: {rel}")
    on_disk = {p.relative_to(pack).as_posix() for p in pack.rglob("*") if p.is_file()} - {MANIFEST, "README.txt"}   # posix: a pack written on Windows is verified on a Mac
    for extra in sorted(on_disk - set(listed)):
        problems.append(f"not in the manifest: {extra}")
    for name in on_disk:
        if Path(name).name in ("config.yml", "config.yaml", ".env") or Path(name).name == "processing.log":
            problems.append(f"a file that must never be packed is here: {name}")
    values = config_secret_values(config if config is not None else load_config())
    for rel, kind in scan_files([pack / r for r in listed], values, root=pack):
        problems.append(f"{kind} in {rel}")
    sha = (manifest.get("databook") or {}).get("sha256")
    db_files = [r for r in listed if r.startswith("databook/")]
    if db_files and sha and sha256_file(pack / db_files[0]) != sha:
        problems.append("the packed databook is not the one the manifest names")
    ident = read_json(pack / "run" / (manifest.get("run") or {}).get("id", "") / "run_identity.json")
    if ident and ident.get("workbook_sha256") != sha and (manifest.get("proof") or {}).get("kind") == "identity":
        problems.append("the run's identity names another databook than the packed one")
    return {"ok": not problems, "problems": problems, "files": len(listed), "status": manifest.get("status")}


# ---------------------------------------------------------------------------------------------------------------------------
# printing (no file, entity or account names unless --names)
# ---------------------------------------------------------------------------------------------------------------------------

def explain_runs(databook: Path, logs_dir: Path, limit: int = 60) -> Tuple[List[str], Dict[str, int]]:
    """Why no run was paired: one line per archived run (newest first) saying what it can prove about this databook. Reads metadata only:
    no extraction, no model, and no file, entity or account name is printed. It pairs nothing; a run it lists as provable still has to
    pass the per-account proof when named with --run."""
    db_sha = sha256_file(databook)
    marker = rebuilt_marker(databook)
    search_sha = (marker or {}).get("source_sha256") or db_sha
    counts = {"these_bytes": 0, "other_bytes_same_name": 0, "other_bytes": 0, "no_identity_fingerprinted": 0, "no_identity_no_fingerprints": 0}
    runs = sorted(list_runs(logs_dir), reverse=True)
    lines: List[str] = []
    for folder in runs[:limit]:
        ident = read_json(folder / "run_identity.json")
        try:
            accounts = [k for k in load_results(folder) if not str(k).startswith("__")]
        except Exception:                                       # noqa: BLE001
            accounts = []
        try:
            from fdd_utils.run_identity import load_run
            prints = {k: v for k, v in ((load_run(folder) or {}).get("fingerprints") or {}).items() if v}
        except Exception:                                       # noqa: BLE001
            prints = {}
        have = len([k for k in accounts if k in prints])
        if ident and ident.get("workbook_sha256") == search_sha:
            tag, key = "identity names THESE bytes", "these_bytes"
        elif ident:
            same_name = ident.get("workbook") == databook.name
            tag = f"identity names other bytes ({str(ident.get('workbook_sha256') or '?')[:8]}), " + ("same file name" if same_name else "another file name")
            key = "other_bytes_same_name" if same_name else "other_bytes"
        elif accounts and have == len(accounts):
            tag, key = "no identity file, every account fingerprinted", "no_identity_fingerprinted"
        else:
            tag, key = "no identity file, accounts without a fingerprint", "no_identity_no_fingerprints"
        counts[key] += 1
        lines.append(f"  {folder.name}  accounts {len(accounts):>3}  fingerprints {have:>3}/{len(accounts):<3}  {tag}")
    head = [f"This databook: sha256 {db_sha[:8]}" + (f" (a rebuild; runs are looked up under {search_sha[:8]})" if marker else "") + f"; {len(runs)} archived run(s), "
            f"{len(lines)} shown, newest first"]
    tail = ["",
            f"  identity names these bytes ........................ {counts['these_bytes']}",
            f"  identity names other bytes, same file name ........ {counts['other_bytes_same_name']}   (this file changed after those runs; the file they were made from is the one needed)",
            f"  identity names other bytes, another file name ..... {counts['other_bytes']}",
            f"  no identity file, every account fingerprinted ..... {counts['no_identity_fingerprinted']}   (can be tried: --run <RUN_ID>; the per-account proof decides)",
            f"  no identity file, fingerprints missing ............ {counts['no_identity_no_fingerprints']}   (nothing ties these to any workbook)"]
    return head + lines + tail, counts


def render(plan: Dict[str, Any], *, names: bool = False) -> List[str]:
    lines = [f"baseline: {plan['status'].upper().replace('_', ' ')}"]
    db = plan.get("databook")
    if db:
        lines.append(f"databook: sha256 {db['sha256'][:12]}, {db['bytes'] // 1024} KB" + (f"  ({Path(db['path']).name})" if names else "")
                     + ("   WARNING: its name suggests a rebuild, demo or copy" if db.get("name_suggests_a_copy") else "")
                     + ("   REBUILT from a text dump: not the original" if db.get("rebuilt") else ""))
    run = plan.get("run")
    if run:
        proof = plan.get("proof") or {}
        accounts = proof.get("accounts") or {}
        lines.append(f"run: {run['id']}" + (f"   kind: {run['kind']}" if run.get("kind") else "") + f"   proof: {proof.get('status')} by {proof.get('kind')}"
                     + (f" ({proof.get('checked')})" if proof.get("checked") else "")
                     + (f"   accounts matched {accounts.get('matched')}/{accounts.get('in_run')}" if accounts.get("in_run") is not None else ""))
    for alt in (plan.get("alternatives") or [])[:5]:
        lines.append(f"  other run that names this databook: {alt['run']} ({alt['kind']}, {alt['accounts']} accounts)")
    if len(plan.get("alternatives") or []) > 5:
        lines.append(f"  ... and {len(plan['alternatives']) - 5} more")
    if plan.get("deck"):
        lines.append(f"deck: found by its lineage sidecar ({plan['deck']['others_paired']} other deck(s) also name this databook)")
    ledger = plan.get("ledger")
    if ledger and ledger["rows"]:
        s = ledger["summary"]
        pr = s["provider_reported"]
        lines.append(f"ledger: {s['rows']} lines for this run; {s['calls_sent']} sent ({s['calls_failed']} failed), {s['response_cache_hits']} answered from the response cache")
        lines.append(f"  provider-reported tokens: {pr['prompt']:,} in + {pr['completion']:,} out = {total_provider_tokens(s):,} "
                     f"over {pr['calls']} call(s); cached prompt {pr['cached_prompt']:,}, reasoning {pr['reasoning']:,}")
        lines.append(f"  calls with no provider usage (estimated only): {s['calls_without_provider_usage']}   (estimates are not added to the figure above)")
        lines.append("  by stage: " + ", ".join(f"{k} {v['calls']}" for k, v in sorted(s["by_stage"].items())))
    for problem in plan.get("problems") or []:
        lines.append(f"PROBLEM: {problem}")
    if plan.get("missing"):
        lines.append("MISSING:")
        for m in plan["missing"]:
            lines.append(f"  - {m['piece']}")
            lines.append(f"      why it matters: {m['why_it_matters']}")
            lines.append(f"      how to get it: {m['how_to_get']}")
    return lines


def make_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--databook", help="the original databook (the file the AI pass was run on)")
    ap.add_argument("--out", help="write the pack here (a git-ignored path); without it nothing is written")
    ap.add_argument("--run", default=None, help="a run id or folder (default: the run whose identity names this databook)")
    ap.add_argument("--deck", default=None, help="the accepted PPTX (default: found by its lineage sidecar)")
    ap.add_argument("--deck-dir", action="append", default=None, help="folder(s) to look for the deck in, instead of pptx_previews/ and fdd_utils/output/ (repeatable)")
    ap.add_argument("--logs-dir", default=str(ROOT / "fdd_utils" / "logs"))
    ap.add_argument("--ledger", default=str(ROOT / "fdd_utils" / "logs" / "usage.jsonl"))
    ap.add_argument("--entity", default=None)
    ap.add_argument("--sheet", default=None)
    ap.add_argument("--financials-from", default=None)
    ap.add_argument("--financials-sheet", default=None)
    ap.add_argument("--no-deep", dest="deep", action="store_false", help="check the identity hash only; skip reading the workbook for the per-account proof")
    ap.add_argument("--leave-out-databook", action="store_true",
                    help="do not copy the databook into the pack (it travels another way, e.g. as a databook_dump.py text); its sha256 is still recorded")
    ap.add_argument("--include-prompts", action="store_true", help="also pack data.yml / audit.jsonl (the stage prompts and replies)")
    ap.add_argument("--allow-urls", action="store_true", help="do not refuse on an https address (after looking at what the hits are); keys and config values are still refused")
    ap.add_argument("--names", action="store_true", help="print file names (the default prints none)")
    ap.add_argument("--verify", default=None, metavar="PACK", help="check a pack that was written earlier")
    ap.add_argument("--explain-runs", action="store_true", help="list every archived run and what it can prove about --databook (metadata only; writes nothing)")
    ap.add_argument("--json", default=None, help="write the plan (with real names) to this local path")
    return ap


def main(argv: Optional[List[str]] = None) -> int:
    ap = make_parser()
    args = ap.parse_args(argv)
    global ALLOW_ADDRESSES
    ALLOW_ADDRESSES = bool(args.allow_urls)
    if args.verify:
        result = verify_pack(Path(args.verify))
        print(("PASS" if result["ok"] else "FAIL") + f": {result.get('files', 0)} file(s), pack status {result.get('status')}")
        for p in result["problems"]:
            print("  -", p)
        return 0 if result["ok"] else 2
    if not args.databook:
        ap.error("--databook is required (or --verify PACK)")
    if args.explain_runs:
        if not Path(args.databook).is_file():
            ap.error("the databook file was not found")
        for line in explain_runs(Path(args.databook), Path(args.logs_dir))[0]:
            print(line)
        return 0
    plan = build_plan(args)
    for line in render(plan, names=args.names):
        print(line)
    if args.json:
        Path(args.json).write_text(json.dumps({k: v for k, v in plan.items() if k != "ledger"}, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    if plan["status"] == "not_a_baseline":
        print("\nNothing was written: this is not a baseline.")
        return 2
    if not args.out:
        print("\nDry run: nothing was written. Add --out <folder under for_test/> to write the pack.")
        return 0 if plan["status"] == "complete" else 1
    out = Path(args.out)
    check_out_path(out)
    result = write_pack(plan, out, leave_out_databook=args.leave_out_databook)
    if not result["ok"]:
        print("\nREFUSED: " + (result.get("reason") or "the files to be packed hold what looks like a secret:"))
        for rel, kind in result.get("hits") or []:
            print(f"  - {kind} in {rel}")
        return 2
    check = verify_pack(out)
    print(f"\nwritten: {result['files']} file(s) under {out if args.names else 'the output folder'};  verify: {'PASS' if check['ok'] else 'FAIL'}")
    for p in check["problems"]:
        print("  -", p)
    return (0 if plan["status"] == "complete" else 1) if check["ok"] else 2


if __name__ == "__main__":
    sys.exit(main())
