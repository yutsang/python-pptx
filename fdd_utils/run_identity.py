"""Which inputs a run's commentary was written from, so it can be reused safely.

Every layout or packing change used to need a fresh ``--run-ai`` pass, because
``--export-pptx`` only exported commentary generated in the same process: ~12
minutes and real tokens to look at a deck whose text had not changed. The text
a run produced is already on disk (results.yml); what was missing is proof
that it belongs to the workbook it would be laid out against, the executive
summary the deck shipped with, and a guarantee that a re-export calls no model.

* ``run_identity.json`` -- written next to results.yml by ``inspect_databook.py
  --run-ai``: the workbook and Financials source (name + sha256), entity,
  sheet, language, model, the accounts the run covered, and the code, config
  and template it ran under.
* ``section_summaries.json`` -- written by ``build_section_summaries``
  (fdd_utils/ui/pptx_export.py) on the CLI and in the app alike.
* ``checkpoint.jsonl`` -- already carries each account's frame fingerprint
  (fdd_utils/ai/evidence.py), which is what lets a run with no identity file,
  e.g. one started from the app, still be checked account by account.

``inspect_databook.py --from-run`` reads all three; ``offline_provider_block``
is how it guarantees zero provider calls.
"""
from __future__ import annotations

import contextlib
import hashlib
import json
import os
import subprocess
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterator, List, Optional

IDENTITY_FILE = "run_identity.json"
REPO_ROOT = Path(__file__).resolve().parent.parent


def file_sha256(path: Optional[str]) -> Optional[str]:
    if not path:
        return None
    try:
        digest = hashlib.sha256()
        with open(path, "rb") as fh:
            for chunk in iter(lambda: fh.read(1 << 20), b""):
                digest.update(chunk)
        return digest.hexdigest()
    except OSError:
        return None


def code_revision() -> str:
    """The checked-out commit, or "" when git is unavailable (the run machine
    has reported an empty commit before; .git/HEAD is read directly then)."""
    try:
        out = subprocess.run(["git", "rev-parse", "--short", "HEAD"], capture_output=True, text=True,
                             timeout=5, cwd=str(REPO_ROOT)).stdout.strip()
        if out:
            return out
    except Exception:
        pass
    try:
        head = (REPO_ROOT / ".git" / "HEAD").read_text(encoding="utf-8").strip()
        if head.startswith("ref:"):
            ref = REPO_ROOT / ".git" / head.split(" ", 1)[1].strip()
            head = ref.read_text(encoding="utf-8").strip() if ref.exists() else ""
        return head[:7]
    except Exception:
        return ""


def environment_fingerprint() -> Dict[str, Any]:
    """Code, config and template a deck was laid out under. config.yml holds the
    API key, so only its hash is ever recorded."""
    return {
        "code_revision": code_revision(),
        "config_sha256": file_sha256(str(REPO_ROOT / "fdd_utils" / "config.yml")),
        "template_sha256": file_sha256(str(REPO_ROOT / "fdd_utils" / "template.pptx")),
    }


def build_identity(*, workbook: str, entity: str, sheet: Optional[str], financials_from: Optional[str],
                   financials_sheet: Optional[str], language: str, model_type: str,
                   model_name: Optional[str], accounts: List[str]) -> Dict[str, Any]:
    return {
        "written_at": datetime.now().isoformat(timespec="seconds"),
        "workbook": Path(workbook).name,
        "workbook_sha256": file_sha256(workbook),
        "entity": entity or "",
        "sheet": sheet or "",
        "financials_from": Path(financials_from).name if financials_from else "",
        "financials_sha256": file_sha256(financials_from) if financials_from else None,
        "financials_sheet": financials_sheet or "",
        "language": language,
        "model_type": model_type,
        "model_name": model_name or "",
        "accounts": sorted(str(a) for a in accounts),
        **environment_fingerprint(),
    }


def write_identity(run_folder: Optional[str], identity: Dict[str, Any]) -> Optional[str]:
    if not run_folder or not os.path.isdir(run_folder):
        return None
    path = os.path.join(run_folder, IDENTITY_FILE)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(identity, fh, ensure_ascii=False, indent=1)
    return path


#: Fields that must agree for a run's commentary to describe this workbook.
#: Model, code, config and template are recorded and reported, not compared:
#: re-laying-out old text under new code is the point of a re-export.
IDENTITY_FIELDS = ("workbook_sha256", "entity", "sheet", "financials_sha256", "financials_sheet", "language")


def compare_identity(saved: Dict[str, Any], current: Dict[str, Any]) -> List[str]:
    return [f"{field}: run={saved.get(field)!r} now={current.get(field)!r}"
            for field in IDENTITY_FIELDS
            if (saved.get(field) or None) != (current.get(field) or None)]


def resolve_run_folder(run: str, log_dir: str = "fdd_utils/logs") -> Optional[str]:
    """A run folder from a path, a folder name or a bare run id."""
    for candidate in (run, os.path.join(log_dir, run), os.path.join(log_dir, "run_" + str(run).replace("run_", ""))):
        if candidate and os.path.isdir(candidate) and os.path.exists(os.path.join(candidate, "results.yml")):
            return candidate
    return None


def load_run(run_folder: str) -> Dict[str, Any]:
    """results.yml, run_identity.json, section_summaries.json and each account's
    last checkpointed frame fingerprint. Absent files come back as None / {}."""
    import yaml

    with open(os.path.join(run_folder, "results.yml"), encoding="utf-8") as fh:
        results = yaml.safe_load(fh) or {}

    def _json(name: str) -> Optional[Dict[str, Any]]:
        path = os.path.join(run_folder, name)
        if not os.path.exists(path):
            return None
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)

    from .ai.logging import PipelineRunLogger
    from .ui.pptx_export import SECTION_SUMMARIES_FILE

    fingerprints: Dict[str, Optional[str]] = {}
    for line in PipelineRunLogger.read_checkpoints(run_folder):
        key = str(line.get("mapping_key") or "")
        if key:
            fingerprints[key] = line.get("fingerprint")
    summaries = _json(SECTION_SUMMARIES_FILE)
    return {
        "results": results,
        "identity": _json(IDENTITY_FILE),
        "summaries": (summaries or {}).get("summaries") if summaries else None,
        "fingerprints": fingerprints,
    }


class OfflineProviderCall(RuntimeError):
    """Raised instead of a model call while ``offline_provider_block`` is active."""


@contextlib.contextmanager
def offline_provider_block() -> Iterator[List[str]]:
    """Every AIClient.get_response raises for the duration, and each attempt is
    recorded (the caller frames) in the yielded list. get_response is the one
    door every model call goes through -- the response cache sits behind it --
    so nothing reaches a provider and nothing is served as if it had."""
    from .ai.client import AIClient

    attempts: List[str] = []
    original = AIClient.get_response

    def _blocked(self, *_args, **_kwargs):
        frames = [f"{Path(f.filename).name}:{f.lineno} {f.name}" for f in traceback.extract_stack()[-6:-1]]
        attempts.append(f"{getattr(self, 'agent_name', '?')} <- " + " <- ".join(reversed(frames)))
        raise OfflineProviderCall("offline re-export: model calls are blocked")

    AIClient.get_response = _blocked
    try:
        yield attempts
    finally:
        AIClient.get_response = original
