#!/usr/bin/env python3
"""Vision smoke test against Workbench GPT-5.5 (default, no model flag).

Confirms two things before building the real contract-extraction pipeline:
  1. PDF pages can be rasterized locally (pypdfium2)
  2. The configured GPT-5.5 deployment actually accepts image input

Uses the same AIClient / config.yml path as the FDD pipeline.

Usage:
    python vision_smoke_test.py contracts
    python vision_smoke_test.py contracts/成都
    python vision_smoke_test.py contracts/成都/some.pdf
"""
from __future__ import annotations

# moved into ad-hoc/ -- put the repo root back on sys.path so
# `import fdd_utils...` still resolves when run from anywhere.
import sys as _sys
from pathlib import Path as _Path
_sys.path.insert(0, str(_Path(__file__).resolve().parents[2]))

import argparse
import sys
from pathlib import Path
from typing import List, Optional, Tuple

from contract_vision import (
    SAFE_SINGLE_IMAGE_BYTES,
    image_file_to_jpeg_bytes,
    is_image_file,
    rasterize_page_jpeg,
    to_data_url,
)
from fdd_utils.ai import AIClient

_DEFAULT_MODEL = "workbench"  # config.yml -> GPT-5.5

_VISION_PROMPT = (
    "Describe in one sentence what kind of document this image shows "
    "(e.g. a lease contract's cover page, a signature page, a rent "
    "schedule table, etc.) and then quote ONE specific piece of text "
    "you can actually read on it, to prove you are reading the real "
    "image content and not guessing from context."
)


def _collect_targets(path: Path) -> List[Path]:
    if path.is_file():
        return [path]
    files: List[Path] = []
    for p in sorted(path.rglob("*")):
        if p.is_file() and (p.suffix.lower() == ".pdf" or is_image_file(p)):
            files.append(p)
    return files


def _load_image(path: Path, page: int) -> bytes:
    if path.suffix.lower() == ".pdf":
        return rasterize_page_jpeg(path, page, max_bytes=SAFE_SINGLE_IMAGE_BYTES)
    return image_file_to_jpeg_bytes(path, max_bytes=SAFE_SINGLE_IMAGE_BYTES)


def _run_one(client: AIClient, path: Path, page: int, root: Optional[Path]) -> Tuple[bool, str, float]:
    try:
        image_bytes = _load_image(path, page)
    except Exception as exc:
        return False, f"rasterize failed: {exc}", 0.0

    vision_content = [
        {"type": "text", "text": _VISION_PROMPT},
        {"type": "image_url", "image_url": {"url": to_data_url(image_bytes)}},
    ]
    try:
        result = client.get_response(
            user_prompt=vision_content,
            system_prompt="You are a helpful assistant.",
        )
    except Exception as exc:
        return False, f"API failed: {exc}", 0.0

    content = str(result.get("content") or "").strip()
    duration = float(result.get("duration") or 0)
    if not content:
        return False, "(empty response)", duration
    return True, content, duration


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Vision smoke test via Workbench GPT-5.5 (no model flag).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python vision_smoke_test.py contracts\n"
            "  python vision_smoke_test.py contracts/成都\n"
            "  python vision_smoke_test.py path/to/file.pdf\n"
        ),
    )
    ap.add_argument("path", help="contracts folder, one project folder, or a single PDF/image")
    ap.add_argument("--page", type=int, default=1, help="PDF page to send (default: 1)")
    args = ap.parse_args()

    path = Path(args.path)
    if not path.exists():
        print(f"❌ Path not found: {path}")
        return 1

    targets = _collect_targets(path)
    if not targets:
        print(f"❌ No PDF/image files found under {path}")
        return 1

    root = path if path.is_dir() else path.parent
    print(f"Model: GPT-5.5 (workbench)  |  page: {args.page}")
    print(f"Files: {len(targets)} under {path}")
    print(f"Image budget: <= {SAFE_SINGLE_IMAGE_BYTES // 1024} KB JPEG (gateway ~4MB body cap)\n")

    try:
        client = AIClient(model_type=_DEFAULT_MODEL, agent_name="subagent_1", language="Eng")
    except Exception as exc:
        print(f"❌ Could not initialize AIClient: {exc}")
        return 1

    ok_n = 0
    fail_n = 0
    for i, target in enumerate(targets, 1):
        label = str(target.relative_to(root)) if target.is_relative_to(root) else target.name
        print("=" * 78)
        print(f"[{i}/{len(targets)}] {label}")
        print("=" * 78)
        ok, detail, duration = _run_one(client, target, args.page, root)
        if ok:
            ok_n += 1
            print(detail)
            print(f"\n✅ ok ({duration:.1f}s)")
        else:
            fail_n += 1
            print(f"❌ {detail}")
        print()

    print("=" * 78)
    print(f"SUMMARY: {ok_n} ok / {fail_n} failed / {len(targets)} total")
    print("=" * 78)
    if fail_n and ok_n == 0:
        print(
            "All calls failed — this deployment may be text-only, or auth/config is wrong. "
            "Paste the first error back before building the extraction pipeline."
        )
        return 1
    if ok_n:
        print(
            "If responses correctly describe lease-contract pages and quote readable text, "
            "vision works — safe to build the full folder extraction next."
        )
    return 0 if fail_n == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
