#!/usr/bin/env python3
"""Extract utility-fee clauses and per-period payment details from lease contracts.

Outputs per project folder (contracts/<project>/):
  合同_水电约定.xlsx
      One sheet per PDF in the folder. One row: source filename + the
      utility-fee arrangement text found in the contract.
  合同_每期支付明细.xlsx
      One sheet per PDF in the folder. Rows = each payment period with
      期间/日期/金额/备注 columns as stated in the contract.

Uses Workbench GPT-5.5 vision (image-only scans) or text layer (rare digital PDFs).
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

from openpyxl import Workbook

from contract_vision import (
    build_page_data_urls,
    extract_pdf_text,
    multi_image_byte_budget,
    pdf_is_digital,
    pdf_page_count,
    select_pages,
)
from fdd_utils.ai import AIClient

_DEFAULT_MODEL = "workbench"
_MAX_TOKENS = 16000
_REASONING = "low"

_UTILITY_HEADERS = [
    "源文件名", "水费约定", "电费约定", "抄表/确认方式",
    "公用事业分摊规则", "结算周期", "税费/税率", "其他约定", "原文摘录",
]
_PAYMENT_HEADERS = [
    "期次", "期间开始", "期间结束", "租赁费（含税）", "物业管理服务费（含税）",
    "其他费用", "合计（含税）", "支付截止日/条件", "备注", "原文摘录",
]

_INVALID_SHEET_CHARS = re.compile(r"[:\\/?*\[\]]")


def _safe_sheet_name(name: str, used: set) -> str:
    base = _INVALID_SHEET_CHARS.sub("_", name).strip() or "sheet"
    base = base[:28]  # leave room for suffix
    out, i = base, 1
    while out.lower() in used or len(out) > 31:
        suffix = f"_{i}"
        out = (base[: 31 - len(suffix)] + suffix)
        i += 1
    used.add(out.lower())
    return out


def _collect_pdfs(folder: Path) -> List[Path]:
    """PDFs in this folder only — project folders should not inherit code/
    tooling folders that happen to sit beside the contracts."""
    if folder.is_file():
        return [folder] if folder.suffix.lower() == ".pdf" else []
    return sorted(p for p in folder.glob("*.pdf") if p.is_file())


def _project_folders(root: Path) -> List[Path]:
    """Folders that actually contain PDFs at their top level.

    If `root` itself holds PDFs (e.g. contracts/成都/*.pdf), treat root as the
    project — not any code subfolders like fdd_utils/ or font_metrics/ that may
    live alongside the contracts. Otherwise (contracts/ with no loose PDFs),
    use each direct subfolder.
    """
    if root.is_file():
        return [root.parent]
    if any(p.is_file() for p in root.glob("*.pdf")):
        return [root]
    subs = [p for p in sorted(root.iterdir()) if p.is_dir() and any(p.glob("*.pdf"))]
    return subs if subs else [root]


def _parse_json_object(text: str) -> Dict[str, Any]:
    raw = (text or "").strip()
    fence = re.search(r"```(?:json)?\s*([\s\S]*?)```", raw, re.IGNORECASE)
    if fence:
        raw = fence.group(1).strip()
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        start, end = raw.find("{"), raw.rfind("}")
        if start < 0 or end <= start:
            raise
        data = json.loads(raw[start : end + 1])
    if not isinstance(data, dict):
        raise ValueError("model output is not a JSON object")
    return data


def _vision_pages(client_prompt_text: str, pdf_path: Path, max_pages: int) -> List[Any]:
    n_pages = pdf_page_count(pdf_path)
    pages = select_pages(n_pages, max_pages=max_pages)
    urls = build_page_data_urls(pdf_path, pages, max_bytes_each=multi_image_byte_budget(len(pages)))
    content: List[Any] = [{"type": "text", "text": client_prompt_text + f"\n提供页面: {pages}\n"}]
    for page_num, url in urls:
        content.append({"type": "text", "text": f"[page {page_num}]"})
        content.append({"type": "image_url", "image_url": {"url": url}})
    return content


def _call_json(client: AIClient, user_prompt: Any) -> Dict[str, Any]:
    result = client.get_response(
        user_prompt=user_prompt,
        system_prompt="Return only one JSON object. No markdown.",
        max_tokens=_MAX_TOKENS,
        reasoning_effort=_REASONING,
    )
    raw = str(result.get("content") or "").strip()
    if not raw:
        raise ValueError(f"empty response (completion_tokens={result.get('completion_tokens')})")
    return _parse_json_object(raw)


def _doc_input(client_prompt_text: str, pdf_path: Path, max_pages: int) -> List[Any] | str:
    if pdf_path.suffix.lower() == ".pdf" and pdf_is_digital(pdf_path):
        text = extract_pdf_text(pdf_path)
        return client_prompt_text + f"\n\n===== PDF TEXT =====\n{text}\n"
    return _vision_pages(client_prompt_text, pdf_path, max_pages)


def _utility_prompt(filename: str) -> str:
    return (
        "你是租赁合同条款抽取助手。只根据提供的内容，抽取该合同里关于水电费的约定。\n"
        f"源文件名: {filename}\n\n"
        "输出一个 JSON 对象（没有就对应字段填空字符串）：\n"
        "{\n"
        '  "水费约定": "",\n'
        '  "电费约定": "",\n'
        '  "抄表/确认方式": "",\n'
        '  "公用事业分摊规则": "",\n'
        '  "结算周期": "",\n'
        '  "税费/税率": "",\n'
        '  "其他约定": "",\n'
        '  "原文摘录": ""\n'
        "}\n"
        "规则:\n"
        "- 各字段为简明短句（可合并多句），没有的填空字符串。\n"
        "- 原文摘录: 抄一句最能代表水电约定的原文（最长约200字）。\n"
        "- 只输出 JSON，不要 markdown。\n"
    )


def _payment_prompt(filename: str) -> str:
    return (
        "你是租赁合同支付明细抽取助手。只根据提供的内容，抽取该合同里每一期应付金额明细。\n"
        f"源文件名: {filename}\n\n"
        "输出一个 JSON 对象：\n"
        "{\n"
        '  "periods": [\n'
        '    {\n'
        '      "期次": "",\n'
        '      "期间开始": "",\n'
        '      "期间结束": "",\n'
        '      "租赁费（含税）": "",\n'
        '      "物业管理服务费（含税）": "",\n'
        '      "其他费用": "",\n'
        '      "合计（含税）": "",\n'
        '      "支付截止日/条件": "",\n'
        '      "备注": "",\n'
        '      "原文摘录": ""\n'
        "    }\n"
        "  ]\n"
        "}\n"
        "规则:\n"
        "- 按期间/期次一行一行列；找不到明细表就 periods 为空数组。\n"
        "- 日期尽量 YYYY-MM-DD；金额不要加千分位。\n"
        "- 每行原文摘录一句（可选，最长约160字）。\n"
        "- 只输出 JSON，不要 markdown。\n"
    )


def _get_str(d: Dict[str, Any], *keys: str) -> str:
    for k in keys:
        v = d.get(k)
        if v is not None and str(v).strip():
            return str(v).strip()
    return ""


def _utility_row(filename: str, data: Dict[str, Any]) -> List[str]:
    return [
        filename,
        _get_str(data, "水费约定"),
        _get_str(data, "电费约定"),
        _get_str(data, "抄表/确认方式", "抄表方式", "确认方式"),
        _get_str(data, "公用事业分摊规则", "分摊规则", "公用事业分摊"),
        _get_str(data, "结算周期", "结算方式"),
        _get_str(data, "税费/税率", "税费", "税率"),
        _get_str(data, "其他约定", "其他"),
        _get_str(data, "原文摘录", "quote"),
    ]


def _payment_rows(data: Dict[str, Any]) -> List[List[str]]:
    periods = data.get("periods")
    if not isinstance(periods, list):
        return []
    rows: List[List[str]] = []
    for item in periods:
        if not isinstance(item, dict):
            continue
        rows.append([
            _get_str(item, "期次", "期数"),
            _get_str(item, "期间开始", "开始日期", "开始"),
            _get_str(item, "期间结束", "结束日期", "结束"),
            _get_str(item, "租赁费（含税）", "租赁费", "租赁费含税"),
            _get_str(item, "物业管理服务费（含税）", "物业管理服务费", "物业服务费"),
            _get_str(item, "其他费用", "其他"),
            _get_str(item, "合计（含税）", "合计", "小计"),
            _get_str(item, "支付截止日/条件", "支付条件", "截止日"),
            _get_str(item, "备注"),
            _get_str(item, "原文摘录", "quote"),
        ])
    return rows


def extract_utility(client: AIClient, pdf: Path, max_pages: int) -> List[str]:
    payload = _doc_input(_utility_prompt(pdf.name), pdf, max_pages)
    data = _call_json(client, payload)
    return _utility_row(pdf.name, data)


def extract_payment(client: AIClient, pdf: Path, max_pages: int) -> List[List[str]]:
    payload = _doc_input(_payment_prompt(pdf.name), pdf, max_pages)
    data = _call_json(client, payload)
    return _payment_rows(data)


def _write_utility_wb(rows_by_file: List[tuple], out_path: Path) -> None:
    wb = Workbook()
    wb.remove(wb.active)
    used = set()
    for filename, row in rows_by_file:
        ws = wb.create_sheet(title=_safe_sheet_name(Path(filename).stem, used))
        ws.append(_UTILITY_HEADERS)
        ws.append(row)
        ws.column_dimensions["A"].width = 40
        for col in "BCDEFGHI":
            ws.column_dimensions[col].width = 26
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def _write_payment_wb(rows_by_file: List[tuple], out_path: Path) -> None:
    wb = Workbook()
    wb.remove(wb.active)
    used = set()
    for filename, rows in rows_by_file:
        ws = wb.create_sheet(title=_safe_sheet_name(Path(filename).stem, used))
        ws.append(_PAYMENT_HEADERS)
        for r in rows:
            ws.append(r)
        ws.column_dimensions["A"].width = 12
        for col in "BCDEFGHIJ":
            ws.column_dimensions[col].width = 20
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Extract 水电约定 + 每期支付明细 from lease contracts (GPT-5.5).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python extract_contract_clauses.py contracts/成都\n"
            "      # writes 合同_水电约定.xlsx + 合同_每期支付明细.xlsx into contracts/成都/\n"
            "  python extract_contract_clauses.py contracts --max-files 5\n"
            "      # every project folder, limit 5 PDFs each (trial run)\n"
        ),
    )
    ap.add_argument("path", help="contracts root or one project folder")
    ap.add_argument("--max-files", type=int, default=0, help="max PDFs per folder (0 = all)")
    ap.add_argument("--max-pages", type=int, default=8, help="max vision pages per PDF")
    ap.add_argument("--skip-utility", action="store_true", help="skip 水电约定 workbook")
    ap.add_argument("--skip-payment", action="store_true", help="skip 每期支付明细 workbook")
    args = ap.parse_args()

    root = Path(args.path)
    if not root.exists():
        print(f"❌ Path not found: {root}")
        return 1

    folders = _project_folders(root)
    print(f"Model: GPT-5.5 (workbench)  |  pages/PDF <= {args.max_pages}")
    print(f"Project folder(s) with top-level PDFs: {len(folders)}")
    for f in folders:
        print(f"  - {f}  ({len(_collect_pdfs(f))} pdf(s))")
    print()

    try:
        client = AIClient(model_type=_DEFAULT_MODEL, agent_name="subagent_2", language="Chi")
    except Exception as exc:
        print(f"❌ Could not initialize AIClient: {exc}")
        return 1

    for folder in folders:
        pdfs = _collect_pdfs(folder)
        if args.max_files and args.max_files > 0:
            pdfs = pdfs[: args.max_files]
        if not pdfs:
            print(f"(skip) no PDFs in {folder}")
            continue

        print("=" * 78)
        print(f"FOLDER: {folder}  ({len(pdfs)} pdf(s))")
        print("=" * 78)

        utility_rows: List[tuple] = []
        payment_rows: List[tuple] = []
        fail = 0

        for i, pdf in enumerate(pdfs, 1):
            print(f"\n[{i}/{len(pdfs)}] {pdf.name}")
            if not args.skip_utility:
                try:
                    utility_rows.append((pdf.name, extract_utility(client, pdf, args.max_pages)))
                    print("  ✅ 水电约定")
                except Exception as exc:
                    fail += 1
                    utility_rows.append((pdf.name, [pdf.name] + [f"抽取失败: {exc}"] + [""] * 7))
                    print(f"  ❌ 水电约定: {exc}")
            if not args.skip_payment:
                try:
                    rows = extract_payment(client, pdf, args.max_pages)
                    payment_rows.append((pdf.name, rows))
                    print(f"  ✅ 支付明细 ({len(rows)} 行)")
                except Exception as exc:
                    fail += 1
                    payment_rows.append((pdf.name, [[""] * 9 + [f"抽取失败: {exc}"]]))
                    print(f"  ❌ 支付明细: {exc}")

        if not args.skip_utility:
            out_u = folder / "合同_水电约定.xlsx"
            _write_utility_wb(utility_rows, out_u)
            print(f"\nWrote {out_u}")
        if not args.skip_payment:
            out_p = folder / "合同_每期支付明细.xlsx"
            _write_payment_wb(payment_rows, out_p)
            print(f"Wrote {out_p}")
        print(f"FOLDER SUMMARY: {len(pdfs) - min(fail, len(pdfs))} ok / {fail} failed")

    print("\nDone. Open the xlsx files above and paste back any rows that look wrong.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
