#!/usr/bin/env python3
"""Inspect lease contracts for utility-fee clauses and per-period payment tables.

Deterministic pre-step before any extraction is built:
  - digital PDFs: scan the real text layer for keywords (no AI)
  - image-only PDFs: optionally send the first pages to Workbench GPT-5.5 and
    ask whether the doc contains 水電費 arrangements / 每期支付明細 tables,
    quoting one line as evidence (LLM, only when not --dry)

Usage:
    python inspect_contracts_clauses.py contracts --dry
        # text-layer only scan, no LLM, no cost
    python inspect_contracts_clauses.py contracts/成都 --max-files 3
        # GPT-5.5 vision on up to 3 PDFs (first pages)
    python inspect_contracts_clauses.py contracts --max-files 0
        # no vision, equivalent to --dry
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

from contract_vision import (
    build_page_data_urls,
    multi_image_byte_budget,
    pdf_is_digital,
    pdf_page_count,
    select_pages,
    extract_pdf_text,
)
from fdd_utils.ai import AIClient

_DEFAULT_MODEL = "workbench"

_UTILITY_KEYWORDS = [
    "水电费", "水電費", "水费", "电费", "水費", "電費",
    "公用事业", "公用事業", "抄表", "结算水电", "代扣水电",
]
_PAYMENT_TABLE_KEYWORDS = [
    "每期应付", "每期應付", "支付计划", "支付計畫", "支付計划", "租金表",
    "租金明細", "租金明细", "物业服务费明细", "物業服務費明細",
    "应付租赁费", "應付租賃費", "支付明细", "支付明細", "结算周期",
]
_GROUP_HINTS = ["补充协议", "主合同", "租赁合同", "承諾函", "承诺函"]


def _collect_pdfs(path: Path, max_files: int) -> List[Path]:
    if path.is_file():
        return [path] if path.suffix.lower() == ".pdf" else []
    files = sorted(p for p in path.rglob("*.pdf") if p.is_file())
    if max_files and max_files > 0:
        return files[:max_files]
    return files


def _scan_text(text: str) -> Dict[str, List[str]]:
    hits: Dict[str, List[str]] = {"utility": [], "payment_table": []}
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    for ln in lines:
        for kw in _UTILITY_KEYWORDS:
            if kw in ln and len(hits["utility"]) < 5:
                hits["utility"].append(ln[:160])
                break
        for kw in _PAYMENT_TABLE_KEYWORDS:
            if kw in ln and len(hits["payment_table"]) < 5:
                hits["payment_table"].append(ln[:160])
                break
        if len(hits["utility"]) >= 5 and len(hits["payment_table"]) >= 5:
            break
    return hits


def _vision_prompt(filename: str) -> str:
    return (
        "你是租赁合同结构侦察助手。请只看提供的页面，判断该合同里是否有：\n"
        "1) 水电费相关约定（费率/抄表/结算/代缴等）\n"
        "2) 每期支付金额大表或支付计划（每期应付租赁费与物业服务费明细、租金表等）\n"
        "输出一个 JSON 对象，格式固定为：\n"
        '{\n'
        '  "utility": {"present": true/false, "quote": "一句原文", "note": ""},\n'
        '  "payment_table": {"present": true/false, "quote": "一句原文", "note": ""}\n'
        '}\n'
        "规则:\n"
        "- present 必须为布尔；找不到就 false。\n"
        "- quote 只在 present 为 true 时给，引用页面上的一句原文；否则空字符串。\n"
        "- 不要输出 JSON 以外的文字。\n"
        "- 只看提供的页面，不要臆测未给页。\n"
    )


def _vision_detect(client: AIClient, pdf_path: Path, max_pages: int) -> Dict[str, Dict[str, Any]]:
    n_pages = pdf_page_count(pdf_path)
    pages = select_pages(n_pages, max_pages=min(5, max_pages))
    urls = build_page_data_urls(pdf_path, pages, max_bytes_each=multi_image_byte_budget(len(pages)))
    content: List[Any] = [{"type": "text", "text": _vision_prompt(pdf_path.name) + f"\n提供页面: {pages}\n"}]
    for page_num, url in urls:
        content.append({"type": "text", "text": f"[page {page_num}]"})
        content.append({"type": "image_url", "image_url": {"url": url}})

    result = client.get_response(
        user_prompt=content,
        system_prompt="Return only the JSON detection object. No markdown.",
        max_tokens=8000,
        reasoning_effort="low",
    )
    raw = str(result.get("content") or "").strip()
    fence = re.search(r"```(?:json)?\s*([\s\S]*?)```", raw, re.IGNORECASE)
    if fence:
        raw = fence.group(1).strip()
    data = json.loads(raw)
    if not isinstance(data, dict):
        raise ValueError("not a JSON object")
    out: Dict[str, Dict[str, Any]] = {}
    for key in ("utility", "payment_table"):
        item = data.get(key)
        if isinstance(item, dict):
            out[key] = {
                "present": bool(item.get("present")),
                "quote": str(item.get("quote") or ""),
                "note": str(item.get("note") or ""),
            }
        else:
            out[key] = {"present": False, "quote": "", "note": "bad model output"}
    return out


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Detect utility-fee clauses and per-period payment tables in lease contracts.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument("path", help="contracts root or one project folder or a single PDF")
    ap.add_argument("--max-files", type=int, default=5,
                    help="max PDFs to vision-scan (0 = text only, no LLM)")
    ap.add_argument("--dry", action="store_true",
                    help="text-layer scan only; never call the LLM")
    ap.add_argument("--max-pages", type=int, default=5, help="max pages per PDF for vision")
    args = ap.parse_args()

    path = Path(args.path)
    if not path.exists():
        print(f"❌ Path not found: {path}")
        return 1

    files = _collect_pdfs(path, 0 if args.dry else args.max_files)
    if not files:
        print(f"❌ No PDF files found under {path}")
        return 1

    print(f"Scope: {path}")
    print(f"Files: {len(files)}  |  mode: {'text-only (--dry)' if args.dry else f'vision <= {args.max_files}'}")
    if not args.dry and args.max_files <= 0:
        print("max-files=0 → text-only (no LLM).")
        args.dry = True
    print()

    client: Optional[AIClient] = None
    if not args.dry:
        try:
            client = AIClient(model_type=_DEFAULT_MODEL, agent_name="subagent_2", language="Chi")
        except Exception as exc:
            print(f"❌ Could not initialize AIClient: {exc}")
            return 1

    summary = {
        "digital": 0,
        "image_only": 0,
        "utility_text_hits": 0,
        "payment_text_hits": 0,
        "utility_vision_present": 0,
        "payment_vision_present": 0,
        "errors": 0,
    }

    for i, pdf in enumerate(files, 1):
        print("=" * 78)
        rel = pdf
        if path.is_dir():
            try:
                rel = pdf.relative_to(path)
            except ValueError:
                rel = pdf
        print(f"[{i}/{len(files)}] {rel}")
        print("=" * 78)

        digital = False
        if pdf.suffix.lower() == ".pdf":
            try:
                digital = pdf_is_digital(pdf)
            except Exception:
                digital = False

        if digital:
            summary["digital"] += 1
            try:
                text = extract_pdf_text(pdf)
                hits = _scan_text(text)
                if hits["utility"]:
                    summary["utility_text_hits"] += 1
                    print("  [text] 水电约定 keywords:")
                    for q in hits["utility"]:
                        print(f"    - {q}")
                else:
                    print("  [text] no 水电 keywords found in text layer")
                if hits["payment_table"]:
                    summary["payment_text_hits"] += 1
                    print("  [text] 支付大表/明细 keywords:")
                    for q in hits["payment_table"]:
                        print(f"    - {q}")
                else:
                    print("  [text] no 支付表 keywords found in text layer")
            except Exception as exc:
                summary["errors"] += 1
                print(f"  ❌ text scan failed: {exc}")
            continue

        summary["image_only"] += 1
        if args.dry:
            print("  [dry] image-only PDF — needs GPT-5.5 vision to check clauses; skipped (no LLM).")
            continue

        assert client is not None
        try:
            det = _vision_detect(client, pdf, max_pages=args.max_pages)
            if det["utility"]["present"]:
                summary["utility_vision_present"] += 1
                print("  [vision] 水电约定: YES")
                if det["utility"]["quote"]:
                    print(f"    quote: {det['utility']['quote'][:160]}")
            else:
                print("  [vision] 水电约定: not seen on provided pages")
            if det["payment_table"]["present"]:
                summary["payment_vision_present"] += 1
                print("  [vision] 支付大表/明细: YES")
                if det["payment_table"]["quote"]:
                    print(f"    quote: {det['payment_table']['quote'][:160]}")
            else:
                print("  [vision] 支付大表/明细: not seen on provided pages")
        except Exception as exc:
            summary["errors"] += 1
            print(f"  ❌ vision scan failed: {exc}")
        print()

    print("=" * 78)
    print("SUMMARY")
    print("=" * 78)
    for k, v in summary.items():
        print(f"  {k}: {v}")
    print(
        "\nNext: paste this back to lock scope before building the per-folder "
        "utility-clause sheet and per-period payment-detail workbook."
    )
    return 0 if summary["errors"] == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
