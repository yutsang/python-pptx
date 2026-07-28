#!/usr/bin/env python3
"""Extract lease-contract fields from a folder of PDFs into the summary template.

Default model: Workbench GPT-5.5 (no --model flag).
Nearly all source PDFs are image-only scans → vision path; rare digital PDFs
use text extraction instead.

Usage:
    python extract_contracts.py contracts
    python extract_contracts.py contracts/成都
    python extract_contracts.py contracts/成都 --validate
        # also compare against any already-filled gold rows in the local template
          (gold values are read at runtime; never stored in this repo)
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import Workbook, load_workbook

from contract_template_schema import CORE_VALUE_COLUMNS, TEMPLATE_COLUMNS, column_map
from contract_vision import (
    SAFE_MULTI_IMAGE_BYTES,
    build_page_data_urls,
    extract_pdf_text,
    is_image_file,
    pdf_is_digital,
    pdf_page_count,
    select_pages,
)
from fdd_utils.ai import AIClient
from inspect_contracts import find_template

_DEFAULT_MODEL = "workbench"
_MISSING = "未提及"
_COL_LETTERS = [letter for letter, _ in TEMPLATE_COLUMNS]
# GPT-5.5 spends max_completion_tokens on hidden reasoning first. Contract
# vision + full A-AF JSON is much heavier than the FDD one-liner smoke test,
# so Generator's 1400 (floored to workbench min_max_tokens=3000) is often
# entirely consumed → empty content with HTTP 200. Override per call.
_EXTRACT_MAX_TOKENS = 16000
_EXTRACT_REASONING = "low"
_RETRY_MAX_TOKENS = 32000


def _collect_pdfs(path: Path) -> List[Path]:
    if path.is_file():
        return [path] if path.suffix.lower() == ".pdf" or is_image_file(path) else []
    return sorted(
        p for p in path.rglob("*")
        if p.is_file() and (p.suffix.lower() == ".pdf" or is_image_file(p))
    )


def _group_by_project(files: List[Path], root: Path) -> List[Tuple[str, List[Path]]]:
    """Group files by first relative folder under root; lone files → root name."""
    groups: Dict[str, List[Path]] = {}
    for f in files:
        try:
            rel = f.relative_to(root)
        except ValueError:
            rel = Path(f.name)
        project = rel.parts[0] if len(rel.parts) > 1 else root.name
        groups.setdefault(project, []).append(f)
    return sorted(groups.items(), key=lambda kv: kv[0])


def _empty_row(filename: str = "") -> Dict[str, str]:
    row = {letter: _MISSING for letter in _COL_LETTERS}
    row["A"] = filename
    row["AC"] = "无"
    return row


def _parse_json_object(text: str) -> Dict[str, Any]:
    raw = (text or "").strip()
    if not raw:
        raise ValueError("empty model response")
    fence = re.search(r"```(?:json)?\s*([\s\S]*?)```", raw, re.IGNORECASE)
    if fence:
        raw = fence.group(1).strip()
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        start = raw.find("{")
        end = raw.rfind("}")
        if start < 0 or end <= start:
            raise
        data = json.loads(raw[start : end + 1])
    if not isinstance(data, dict):
        raise ValueError("model JSON is not an object")
    return data


def _normalize_row(data: Dict[str, Any], filename: str) -> Dict[str, str]:
    headers = column_map()
    header_to_letter = {h: letter for letter, h in headers.items()}
    # also accept header without newlines / extra spaces
    loose = {re.sub(r"\s+", "", h): letter for letter, h in headers.items()}

    row = _empty_row(filename)
    for key, value in data.items():
        key_s = str(key).strip()
        letter = None
        if key_s.upper() in headers:
            letter = key_s.upper()
            if letter == "A" and key_s == "a":
                letter = "A"
        elif key_s in header_to_letter:
            letter = header_to_letter[key_s]
        else:
            letter = loose.get(re.sub(r"\s+", "", key_s))
        if not letter or letter not in row:
            continue
        if value is None:
            continue
        text = str(value).strip()
        if not text:
            continue
        row[letter] = text
    row["A"] = filename
    return row


def _schema_prompt_block() -> str:
    lines = []
    for letter, header in TEMPLATE_COLUMNS:
        if letter == "A":
            continue
        lines.append(f'  "{letter}": "{header}"')
    return "{\n" + ",\n".join(lines) + "\n}"


def _build_extraction_prompt(filename: str, page_note: str) -> str:
    return (
        "你是租赁合同信息抽取助手。根据提供的合同页面内容，填写租赁台账字段。\n"
        f"源文件名: {filename}\n"
        f"{page_note}\n\n"
        "规则:\n"
        f"1. 只输出一个 JSON 对象，key 必须是下列 Excel 列字母；不要 markdown，不要解释。\n"
        f"2. 页面上看不到的字段填 \"{_MISSING}\"；备注没有内容时填 \"无\"。\n"
        "3. 日期尽量用 YYYY-MM-DD；数字不要加千分位。\n"
        "4. 长文本字段 V/W/Z/AA/AB：最多各摘录 120 字关键句；不要整段照抄；"
        f"页面未覆盖则填 \"{_MISSING}\"。\n"
        "5. 不要输出 JSON 以外的任何文字。\n\n"
        "目标字段 (列字母: 含义):\n"
        f"{_schema_prompt_block()}\n"
    )


def _call_model(client: AIClient, user_prompt: Any, *, max_tokens: int) -> Dict[str, Any]:
    return client.get_response(
        user_prompt=user_prompt,
        system_prompt="Return only a single JSON object for the lease ledger fields. No markdown.",
        max_tokens=max_tokens,
        reasoning_effort=_EXTRACT_REASONING,
    )


def _extract_payload(client: AIClient, user_prompt: Any, filename: str) -> Tuple[Dict[str, str], float]:
    result = _call_model(client, user_prompt, max_tokens=_EXTRACT_MAX_TOKENS)
    content = str(result.get("content") or "").strip()
    duration = float(result.get("duration") or 0)
    if not content:
        print(
            f"  ⟳ empty response "
            f"(completion_tokens={result.get('completion_tokens')}, "
            f"{duration:.1f}s) — retry with max_tokens={_RETRY_MAX_TOKENS}"
        )
        result = _call_model(client, user_prompt, max_tokens=_RETRY_MAX_TOKENS)
        content = str(result.get("content") or "").strip()
        duration = float(result.get("duration") or 0)
    if not content:
        raise ValueError(
            f"empty model response after retry "
            f"(completion_tokens={result.get('completion_tokens')}, duration={duration}s)"
        )
    return _normalize_row(_parse_json_object(content), filename), duration


def _extract_from_digital(client: AIClient, path: Path) -> Tuple[Dict[str, str], float]:
    text = extract_pdf_text(path)
    prompt = _build_extraction_prompt(
        path.name,
        "以下是可直接抽取的电子版 PDF 全文（或截断后的首尾）。",
    ) + f"\n\n===== PDF TEXT =====\n{text}\n"
    return _extract_payload(client, prompt, path.name)


def _extract_from_vision(client: AIClient, path: Path, max_pages: int) -> Tuple[Dict[str, str], float]:
    if path.suffix.lower() != ".pdf":
        from contract_vision import image_file_to_jpeg_bytes, to_data_url

        raw = image_file_to_jpeg_bytes(path, max_bytes=SAFE_MULTI_IMAGE_BYTES)
        content: List[Any] = [
            {"type": "text", "text": _build_extraction_prompt(path.name, "单页图片合同。")},
            {"type": "image_url", "image_url": {"url": to_data_url(raw)}},
        ]
        return _extract_payload(client, content, path.name)

    n_pages = pdf_page_count(path)
    pages = select_pages(n_pages, max_pages=max_pages)
    urls = build_page_data_urls(path, pages, max_bytes_each=SAFE_MULTI_IMAGE_BYTES)
    page_note = (
        f"合同共 {n_pages} 页；以下仅提供第 {', '.join(str(p) for p, _ in urls)} 页的扫描图。"
        "只根据这些页面填写；未出现的字段填 "
        f"\"{_MISSING}\"。"
    )
    content = [{"type": "text", "text": _build_extraction_prompt(path.name, page_note)}]
    for page_num, url in urls:
        content.append({"type": "text", "text": f"[page {page_num}]"})
        content.append({"type": "image_url", "image_url": {"url": url}})

    return _extract_payload(client, content, path.name)


def extract_one(client: AIClient, path: Path, max_pages: int) -> Tuple[Dict[str, str], float, str]:
    if path.suffix.lower() == ".pdf" and pdf_is_digital(path):
        row, dur = _extract_from_digital(client, path)
        return row, dur, "digital-text"
    row, dur = _extract_from_vision(client, path, max_pages=max_pages)
    return row, dur, "vision"


def _load_gold_rows(template_path: Path) -> Dict[str, Dict[str, str]]:
    """filename -> row dict, from already-filled data rows in the local template."""
    wb = load_workbook(template_path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    gold: Dict[str, Dict[str, str]] = {}
    for row_idx in range(1, (ws.max_row or 0) + 1):
        filename = ws.cell(row=row_idx, column=1).value
        if not filename or not str(filename).lower().endswith((".pdf", ".jpg", ".png")):
            continue
        party_a = ws.cell(row=row_idx, column=2).value
        if party_a is None or str(party_a).strip() == "":
            continue
        row = _empty_row(str(filename).strip())
        for col_idx, letter in enumerate(_COL_LETTERS, start=1):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val is None:
                continue
            row[letter] = str(val).strip()
        gold[row["A"]] = row
    return gold


def _norm_cmp(value: str) -> str:
    s = str(value or "").strip()
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"\s+", "", s)
    # date noise from Excel
    if s.endswith("00:00:00"):
        s = s[: -len("00:00:00")]
    return s.lower()


def _validate_against_gold(
    extracted: Dict[str, Dict[str, str]],
    gold: Dict[str, Dict[str, str]],
) -> int:
    """Print field-level diffs for overlapping filenames. Returns mismatch count."""
    overlap = sorted(set(extracted) & set(gold))
    if not overlap:
        print("\n(validate) No overlapping filenames between extraction and template gold rows.")
        return 0
    print(f"\n(validate) Comparing {len(overlap)} file(s) against local template gold rows "
          f"on columns: {', '.join(CORE_VALUE_COLUMNS)}")
    mismatches = 0
    for name in overlap:
        print(f"\n--- {name} ---")
        bad = []
        for letter in CORE_VALUE_COLUMNS:
            got = _norm_cmp(extracted[name].get(letter, ""))
            exp = _norm_cmp(gold[name].get(letter, ""))
            if got == exp:
                continue
            # numeric tolerance for floats
            try:
                if abs(float(got) - float(exp)) <= 1e-6 * max(1.0, abs(float(exp))):
                    continue
            except Exception:
                pass
            bad.append(letter)
            mismatches += 1
            print(f"  {letter} mismatch")
            print(f"    gold: {gold[name].get(letter, '')[:120]}")
            print(f"    got:  {extracted[name].get(letter, '')[:120]}")
        if not bad:
            print("  ✅ all core columns match")
    print(f"\n(validate) core-field mismatches: {mismatches}")
    return mismatches


def _write_workbook(
    groups: List[Tuple[str, List[Dict[str, str]]]],
    output_path: Path,
    title: str = "租赁台账",
) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.cell(row=2, column=2, value=title)
    for col_idx, (letter, header) in enumerate(TEMPLATE_COLUMNS, start=1):
        if letter == "A":
            continue
        ws.cell(row=3, column=col_idx, value=header)

    r = 4
    for project, rows in groups:
        ws.cell(row=r, column=2, value=f"{project}")
        r += 1
        for row in rows:
            for col_idx, letter in enumerate(_COL_LETTERS, start=1):
                ws.cell(row=r, column=col_idx, value=row.get(letter, _MISSING))
            r += 1
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Extract lease contracts via Workbench GPT-5.5 into the summary template.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python extract_contracts.py contracts\n"
            "  python extract_contracts.py contracts/成都\n"
            "  python extract_contracts.py contracts/成都 --validate\n"
        ),
    )
    ap.add_argument("path", help="contracts root, one project folder, or a single PDF")
    ap.add_argument("--template", default=None, help="path to 合同汇总模板.xlsx (auto-detected if omitted)")
    ap.add_argument("--out", default=None, help="output xlsx path (default: <path>/合同汇总_extracted.xlsx)")
    ap.add_argument("--max-pages", type=int, default=4, help="max vision pages per PDF (default: 4)")
    ap.add_argument(
        "--validate",
        action="store_true",
        help="compare results to filled gold rows already present in the local template",
    )
    args = ap.parse_args()

    path = Path(args.path)
    if not path.exists():
        print(f"❌ Path not found: {path}")
        return 1

    files = _collect_pdfs(path)
    if not files:
        print(f"❌ No PDF/image files found under {path}")
        return 1

    root = path if path.is_dir() else path.parent
    # If user pointed at a project subfolder, keep that folder as the project name.
    group_root = root

    template_path = Path(args.template) if args.template else find_template(root)
    if template_path is None and path.is_dir():
        template_path = find_template(path)
    if template_path is None and root.parent.exists():
        template_path = find_template(root.parent)

    out_path = Path(args.out) if args.out else (root / "合同汇总_extracted.xlsx")

    print(f"Model: GPT-5.5 (workbench)  |  max_tokens={_EXTRACT_MAX_TOKENS}  |  reasoning={_EXTRACT_REASONING}")
    print(f"Files: {len(files)} under {path}")
    print(f"Vision pages/PDF: up to {args.max_pages}")
    print(f"Output: {out_path}")
    if template_path:
        print(f"Template: {template_path}")
        if template_path.name.startswith("~$"):
            print("❌ Template looks like an Excel lock file (~$...). Close the workbook and retry.")
            return 1
    elif args.validate:
        print("❌ --validate needs a local template with gold rows")
        print("   Tip: close Excel if 合同汇总模板.xlsx is open (avoids ~$ lock files).")
        return 1
    print()

    try:
        # Auditor defaults (low reasoning) — still overridden per-call below.
        client = AIClient(model_type=_DEFAULT_MODEL, agent_name="subagent_2", language="Chi")
    except Exception as exc:
        print(f"❌ Could not initialize AIClient: {exc}")
        return 1

    extracted_by_name: Dict[str, Dict[str, str]] = {}
    grouped_files = _group_by_project(files, group_root)
    grouped_rows: List[Tuple[str, List[Dict[str, str]]]] = []
    fail_n = 0

    file_i = 0
    for project, project_files in grouped_files:
        rows: List[Dict[str, str]] = []
        for f in project_files:
            file_i += 1
            print("=" * 78)
            print(f"[{file_i}/{len(files)}] {f.name}")
            print("=" * 78)
            try:
                row, dur, mode = extract_one(client, f, max_pages=args.max_pages)
                extracted_by_name[f.name] = row
                rows.append(row)
                print(f"✅ {mode} ({dur:.1f}s)  甲方={row.get('B', '')[:40]}  乙方={row.get('C', '')[:40]}")
            except Exception as exc:
                fail_n += 1
                err_row = _empty_row(f.name)
                err_row["AC"] = f"抽取失败: {exc}"
                rows.append(err_row)
                extracted_by_name[f.name] = err_row
                print(f"❌ {exc}")
            print()
        grouped_rows.append((project, rows))

    _write_workbook(grouped_rows, out_path)
    print("=" * 78)
    print(f"Wrote {out_path}")
    print(f"SUMMARY: {len(files) - fail_n} ok / {fail_n} failed / {len(files)} total")
    print("=" * 78)

    mismatch_n = 0
    if args.validate and template_path:
        gold = _load_gold_rows(template_path)
        mismatch_n = _validate_against_gold(extracted_by_name, gold)

    if fail_n:
        return 2
    if args.validate and mismatch_n:
        return 3
    return 0


if __name__ == "__main__":
    sys.exit(main())
