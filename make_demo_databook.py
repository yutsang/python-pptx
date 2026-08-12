#!/usr/bin/env python3
"""
Turn a real client databook into a demo one: scale every figure by a single
factor, and replace every real company / entity name with a stable fake.

Why one global scale factor rather than per-row noise: the workbook carries
ZERO formulas -- every Subtotal, Total and `Check` row is a hardcoded value.
Anything other than a linear transform means recomputing the whole tie-out
structure by hand and hoping the inferred structure was right.  Scaling is
linear, so every detail sum, subtotal, BS balance, Financials<->schedule
tie-out and `Check = 0` row survives untouched -- and the amounts quoted
inside the Chinese remark text ("应收租金 29.1w") can be scaled by the same
factor and stay consistent with the table beside them.

Usage:
    python make_demo_databook.py "Project Gold Kunshan.databook.xlsx"
    python make_demo_databook.py <src> --scale 0.8734 --rename 昆山润泽置业有限公司=昆山瑞茂置业有限公司
    python make_demo_databook.py <src> --dry-run
"""

from __future__ import annotations

import argparse
import hashlib
import re
import shutil
import zipfile
from collections import defaultdict
from pathlib import Path

import openpyxl

DEFAULT_SCALE = 0.8734

# ---------------------------------------------------------------- name parsing

# Geography kept as-is: the user wants the demo to still read as a real
# Chinese logistics deal, and a scrambled province makes the commentary absurd.
PLACE_PREFIXES = [
    "内蒙古", "黑龙江", "石家庄", "哈尔滨", "乌鲁木齐",
    "上海", "北京", "深圳", "广州", "天津", "重庆", "成都", "武汉", "南京",
    "杭州", "苏州", "无锡", "常州", "昆山", "宁波", "青岛", "济南", "西安",
    "郑州", "长沙", "合肥", "福州", "厦门", "大连", "沈阳", "佛山", "东莞",
    "珠海", "中山", "惠州", "嘉兴", "绍兴", "temp",
    "浙江", "江苏", "山东", "广东", "河北", "河南", "湖北", "湖南", "四川",
    "福建", "安徽", "辽宁", "陕西", "山西", "江西", "云南", "贵州", "广西",
    "海南", "甘肃", "青海", "宁夏", "新疆", "西藏", "吉林",
    "中国", "华东", "华南", "华北", "华中", "西南", "东北",
]
PLACE_PREFIXES = [p for p in PLACE_PREFIXES if p != "temp"]
PLACE_PREFIXES.sort(key=len, reverse=True)

# Industry tail, kept as-is so the fake name still describes the same business.
INDUSTRY_TAILS = [
    "投资管理咨询", "企业管理咨询", "供应链管理", "设备安装工程", "房地产经纪",
    "信息科技", "货物运输", "仓储服务", "物业服务", "物业管理", "投资管理",
    "造价咨询", "管理咨询", "安装工程", "财产保险", "供应链", "电动车",
    "新能源", "房地产", "商贸", "贸易", "物流", "电子", "实业", "科技",
    "仓储", "建设", "工程", "咨询", "保险", "运输", "置业", "装饰", "机械",
    "材料", "食品", "医药", "生物", "环保", "能源", "租赁", "销售", "服务",
    "发展", "实体", "工贸", "农业", "化工", "纺织", "家具", "包装", "印刷",
]
INDUSTRY_TAILS.sort(key=len, reverse=True)

COMPANY_SUFFIXES = [
    "股份有限公司", "有限责任公司", "有限公司", "合伙企业", "分公司",
    "集团", "公司", "中心", "工厂", "厂",
]
COMPANY_SUFFIXES.sort(key=len, reverse=True)

# Generic references that look like company names but name nobody.
GENERIC_NAMES = {
    "目标公司", "本公司", "母公司", "子公司", "关联公司", "总公司", "分公司",
    "公司", "第三方", "该公司", "集团", "关联方", "供应商", "承租人", "出租人",
    "物业公司", "施工单位", "建设单位", "保险公司", "银行", "税务局",
}

# A whole cell that is nothing but a company name.  Deliberately anchored:
# a long remark sentence that merely *mentions* a company must NOT match here,
# or the extracted "real name" would be a sentence fragment.  Remarks are
# handled by full-text substitution of the names found in the clean cells.
COMPANY_CELL_RE = re.compile(
    r"^[一-鿿（）()·\s]{2,30}(?:"
    + "|".join(COMPANY_SUFFIXES)
    + r")$"
)

BRAND_CHARS = (
    "瑞宏恒嘉启泰盛鑫达丰隆通和兴利源正德信顺安佳元茂康联创立明辉阳"
    "卓越弘毅锦程远航睿臻拓新润泽晟昊博轩宁禾岳川岭峰晖朗骏毅睿祺"
)


def split_company(name: str) -> tuple[str, str, str]:
    """Split into (place prefix, brand core, industry+suffix tail)."""
    rest = name
    prefix = ""
    for p in PLACE_PREFIXES:
        if rest.startswith(p):
            prefix, rest = p, rest[len(p):]
            break

    suffix = ""
    for s in COMPANY_SUFFIXES:
        if rest.endswith(s):
            suffix, rest = s, rest[: -len(s)]
            break

    # A second place name can sit before the suffix: "...有限公司上海市分公司".
    industry = ""
    for t in INDUSTRY_TAILS:
        if rest.endswith(t):
            industry, rest = t, rest[: -len(t)]
            break

    return prefix, rest, industry + suffix


# "百盟（上海）供应链有限公司" carries a place name INSIDE the brand.  Blindly
# scrambling those six characters turns the bracket pair into gibberish, so the
# bracketed place is carved out and kept.
BRACKET_PLACE_RE = re.compile(
    r"[（(](?:" + "|".join(PLACE_PREFIXES) + r")[）)]"
)

# A brand this long is not a brand: it is a compound name the splitter failed to
# parse (e.g. "...股份有限公司上海市分公司").  Flagged for an explicit --rename
# rather than silently scrambled.
LONG_BRAND_WARN = 8


def _scramble(seed: str, length: int, salt: int) -> str:
    digest = hashlib.md5(f"{seed}|{salt}".encode()).digest()
    while len(digest) < length:
        digest += hashlib.md5(digest).digest()
    return "".join(BRAND_CHARS[b % len(BRAND_CHARS)] for b in digest[:length])


def fake_brand(real_brand: str, salt: int = 0) -> str:
    """Deterministic same-length replacement brand.

    Same length matters more than it looks: PPTX packing and line-wrap are
    measured per character, so a demo file whose tenant names are a different
    width stops being a valid layout test.
    """
    m = BRACKET_PLACE_RE.search(real_brand)
    if m:
        head, kept, tail = real_brand[: m.start()], m.group(0), real_brand[m.end():]
        return _scramble(head, len(head), salt) + kept + _scramble(tail, len(tail), salt + 1)
    return _scramble(real_brand, len(real_brand), salt)


def brand_warning(name: str) -> str:
    """Why this name needs a human decision, or '' if it parses cleanly."""
    _, brand, _ = split_company(name)
    if not brand:
        return "商號為空，不會被替換"
    if any(suf in brand for suf in COMPANY_SUFFIXES):
        return "商號內還有第二個公司後綴 -> 複合名稱，請用 --rename 指定"
    if len(brand) > LONG_BRAND_WARN:
        return f"商號長達 {len(brand)} 字 -> 可能夾雜其他文字，請用 --rename 指定"
    return ""


def build_name_map(real_names: set[str]) -> dict[str, str]:
    """Map each real company name to a stable fake of identical length."""
    mapping: dict[str, str] = {}
    taken: set[str] = set()
    for name in sorted(real_names):
        prefix, brand, tail = split_company(name)
        if not brand:  # e.g. bare "上海分公司" -- nothing to anonymise
            continue
        salt = 0
        while True:
            candidate = prefix + fake_brand(brand, salt) + tail
            if candidate not in taken and candidate != name:
                break
            salt += 1
        mapping[name] = candidate
        taken.add(candidate)
    return mapping


def build_abbrev_map(name_map: dict[str, str]) -> dict[str, str]:
    """Short forms used in the remark text ("尚泽" for 上海尚泽铭电子有限公司).

    Only the brand core and its 2-char head are mapped, and only when they are
    long enough not to be ordinary Chinese.  Every hit is reported so the
    substitutions can be eyeballed rather than trusted.
    """
    abbrev: dict[str, str] = {}
    for real, fake in name_map.items():
        _, real_brand, _ = split_company(real)
        _, fake_brand_core, _ = split_company(fake)
        if len(real_brand) >= 2 and real_brand not in name_map:
            abbrev[real_brand] = fake_brand_core
        if len(real_brand) >= 3:
            abbrev[real_brand[:2]] = fake_brand_core[:2]
    return abbrev


# ------------------------------------------------------------- text scaling

# Only numbers carrying an explicit money unit get scaled.  "25.1" in
# "应收25.1一个月的含税租金" is a year-month, and "26年2月3日" is a date --
# scaling either would corrupt the narrative rather than anonymise it.
MONEY_RE = re.compile(
    r"(?<![\d.年月])(\d{1,3}(?:,\d{3})+|\d+)(\.\d+)?\s*(w元|W元|w|W|万元|万|元)"
)


def scale_money_in_text(text: str, factor: float, hits: list[tuple[str, str]]) -> str:
    def repl(m: re.Match) -> str:
        int_part, dec_part, unit = m.group(1), m.group(2) or "", m.group(3)
        try:
            value = float(int_part.replace(",", "") + dec_part)
        except ValueError:
            return m.group(0)
        scaled = value * factor
        decimals = len(dec_part) - 1 if dec_part else 0
        new_text = f"{scaled:.{decimals}f}{unit}"
        hits.append((m.group(0), new_text))
        return new_text

    return MONEY_RE.sub(repl, text)


# ------------------------------------------------------------------- scanning


def scan_workbook(wb) -> tuple[set[str], dict[str, set[str]]]:
    """Collect clean company-name cells and the numeric columns per sheet."""
    names: set[str] = set()
    numeric_cols: dict[str, dict[str, list[float]]] = defaultdict(lambda: defaultdict(list))
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str):
                    s = v.strip()
                    if s not in GENERIC_NAMES and COMPANY_CELL_RE.match(s):
                        names.add(s)
                elif isinstance(v, (int, float)) and not isinstance(v, bool):
                    numeric_cols[ws.title][cell.column_letter].append(float(v))
    return names, numeric_cols


def detect_ratio_columns(numeric_cols) -> dict[str, set[str]]:
    """Columns that are ratios, not amounts -- scaling them breaks the ratio.

    A ratio column keeps its value under a global scale (numerator and
    denominator both move), so it must be left alone.  Heuristic: every
    non-zero value in the column is within +/-1, and there are at least two
    of them.  A real amount column always carries something bigger.
    """
    ratio: dict[str, set[str]] = defaultdict(set)
    for sheet, cols in numeric_cols.items():
        for col, values in cols.items():
            nz = [v for v in values if v != 0]
            if len(nz) >= 2 and all(abs(v) <= 1.0 for v in nz):
                ratio[sheet].add(col)
    return ratio


# ------------------------------------------------------------------ transform


def apply_replacements(text: str, ordered_pairs: list[tuple[str, str]]) -> tuple[str, list[str]]:
    used: list[str] = []
    for real, fake in ordered_pairs:
        if real in text:
            text = text.replace(real, fake)
            used.append(real)
    return text, used


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("source", help="Path to the real databook .xlsx")
    ap.add_argument("--output", help="Output path (default: '[DEMO]<stem>.xlsx' beside this script)")
    ap.add_argument("--scale", type=float, default=DEFAULT_SCALE, help=f"Global factor (default {DEFAULT_SCALE})")
    ap.add_argument("--rename", action="append", default=[],
                    help="Extra literal replacement 'old=new'; repeatable. Use for project/entity names.")
    ap.add_argument("--no-abbrev", action="store_true", help="Do not substitute brand short forms in remark text")
    ap.add_argument("--no-ratio-detect", action="store_true", help="Scale every numeric cell, including ratio columns")
    ap.add_argument("--inspect", action="store_true",
                    help="Diagnose the source only. Prints a REDACTED report (safe to share) "
                         "and writes the full one to --inspect-out.")
    ap.add_argument("--inspect-out", default="demo_inspect_FULL.txt",
                    help="Where the unredacted inspect report goes (default: demo_inspect_FULL.txt)")
    ap.add_argument("--dry-run", action="store_true", help="Report only; write nothing")
    ap.add_argument("--emit-mapping", help="Write the real->fake table to this file (SENSITIVE: never commit it)")
    args = ap.parse_args()

    src = Path(args.source)
    if not src.exists():
        print(f"ERROR: source not found: {src}")
        return 1
    out = Path(args.output) if args.output else src.parent / f"[DEMO]{src.stem}.xlsx"

    with zipfile.ZipFile(src) as z:
        embedded = [n for n in z.namelist() if n.startswith(("xl/media/", "xl/charts/"))]
    if embedded:
        print(f"WARNING: source contains {len(embedded)} embedded image/chart part(s); "
              f"openpyxl will drop them from the demo file.")

    wb = openpyxl.load_workbook(src, data_only=False)
    real_names, numeric_cols = scan_workbook(wb)
    ratio_cols = {} if args.no_ratio_detect else detect_ratio_columns(numeric_cols)

    name_map = build_name_map(real_names)

    if args.inspect:
        return inspect_only(src, wb, real_names, numeric_cols, ratio_cols,
                            name_map, Path(args.inspect_out))

    abbrev_map = {} if args.no_abbrev else build_abbrev_map(name_map)

    extra_map: dict[str, str] = {}
    for spec in args.rename:
        if "=" not in spec:
            print(f"ERROR: --rename needs 'old=new', got {spec!r}")
            return 1
        old, new = spec.split("=", 1)
        extra_map[old] = new

    # Longest first, so "上海邵奥贸易有限公司" is consumed before the "邵奥"
    # abbreviation rule can chew a hole in the middle of it.
    ordered_pairs = sorted({**name_map, **abbrev_map, **extra_map}.items(),
                           key=lambda kv: len(kv[0]), reverse=True)

    print(f"\n=== SOURCE: {src.name} ===")
    print(f"  sheets: {len(wb.sheetnames)}   scale factor: {args.scale}")
    print(f"\n--- company names found ({len(name_map)}) ---")
    warned = 0
    for real, fake in sorted(name_map.items()):
        warn = brand_warning(real)
        print(f"  {real}  ->  {fake}" + (f"   ⚠ {warn}" if warn else ""))
        warned += bool(warn)
    if warned:
        print(f"  ⚠ {warned} name(s) above need an explicit --rename; the auto-generated")
        print(f"    replacement for them will look like gibberish.")
    if extra_map:
        print(f"\n--- extra --rename rules ({len(extra_map)}) ---")
        for old, new in extra_map.items():
            print(f"  {old}  ->  {new}")
    if abbrev_map:
        print(f"\n--- brand short forms substituted in remark text ({len(abbrev_map)}) ---")
        for old, new in sorted(abbrev_map.items()):
            print(f"  {old}  ->  {new}")
    if ratio_cols:
        print("\n--- columns treated as RATIOS and left unscaled ---")
        for sheet, cols in sorted(ratio_cols.items()):
            print(f"  {sheet}: {', '.join(sorted(cols))}")

    scaled_cells = 0
    skipped_ratio = 0
    text_cells_changed = 0
    money_hits: list[tuple[str, str]] = []
    names_hit: set[str] = set()

    for ws in wb.worksheets:
        sheet_ratio = ratio_cols.get(ws.title, set())
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, bool) or v is None:
                    continue
                if isinstance(v, (int, float)):
                    if cell.column_letter in sheet_ratio:
                        skipped_ratio += 1
                        continue
                    cell.value = v * args.scale
                    scaled_cells += 1
                elif isinstance(v, str):
                    new, used = apply_replacements(v, ordered_pairs)
                    names_hit.update(used)
                    new = scale_money_in_text(new, args.scale, money_hits)
                    if new != v:
                        cell.value = new
                        text_cells_changed += 1

    renamed_sheets = []
    for ws in wb.worksheets:
        new_title, _ = apply_replacements(ws.title, ordered_pairs)
        if new_title != ws.title:
            renamed_sheets.append((ws.title, new_title))
            ws.title = new_title

    print("\n--- transform ---")
    print(f"  numeric cells scaled : {scaled_cells}")
    print(f"  numeric cells skipped (ratio columns) : {skipped_ratio}")
    print(f"  text cells changed   : {text_cells_changed}")
    print(f"  sheets renamed       : {len(renamed_sheets)}")
    for old, new in renamed_sheets:
        print(f"      {old}  ->  {new}")
    if money_hits:
        print(f"  amounts rewritten inside remark text ({len(money_hits)}):")
        for old, new in money_hits[:40]:
            print(f"      {old}  ->  {new}")
        if len(money_hits) > 40:
            print(f"      ... and {len(money_hits) - 40} more")

    unhit = sorted(set(name_map) - names_hit)
    if unhit:
        print(f"\n  NOTE: {len(unhit)} mapped name(s) never appeared during substitution "
              f"(they were found in a cell that got replaced by a longer rule first):")
        for n in unhit:
            print(f"      {n}")

    if args.dry_run:
        print("\n[dry run] nothing written.")
        return 0

    if out.exists():
        backup = out.with_suffix(out.suffix + ".bak")
        shutil.copy2(out, backup)
        print(f"\n  existing output backed up to {backup.name}")
    wb.save(out)
    print(f"\n  written: {out}")

    if args.emit_mapping:
        lines = ["# SENSITIVE - real -> fake mapping. Do NOT commit.", ""]
        lines += [f"{r}\t{f}" for r, f in sorted({**name_map, **extra_map}.items())]
        Path(args.emit_mapping).write_text("\n".join(lines) + "\n", encoding="utf-8")
        print(f"  mapping written: {args.emit_mapping}  (SENSITIVE - do not commit)")

    return verify(src, out, args.scale, name_map, extra_map, ratio_cols)


# ------------------------------------------------------------------- inspect

CJK_RE = re.compile(r"[一-鿿]")
PURE_CN_RE = re.compile(r"^[一-鿿]{2,8}$")
SUBTOTAL_WORDS = ("subtotal", "total", "check", "小计", "小計", "合计", "合計", "总计", "總計")
# Deliberately loose: the point is to discover units the money regex does NOT
# yet cover (百万, 亿, k, K...), not to confirm the ones it does.
NUM_UNIT_PROBE = re.compile(r"\d(?:[\d,]*\.?\d*)\s*([一-鿿A-Za-z%]{1,3})")


def redact(s: str) -> str:
    """Structure-preserving redaction: shape survives, content does not."""
    s = CJK_RE.sub("〇", s)
    s = re.sub(r"\d", "N", s)
    return re.sub(r"[A-Za-z]", "A", s)


def redact_company(name: str) -> str:
    prefix, brand, tail = split_company(name)
    m = BRACKET_PLACE_RE.search(brand)
    if m:
        shown = "〇" * m.start() + m.group(0) + "〇" * (len(brand) - m.end())
    else:
        shown = "〇" * len(brand)
    warn = brand_warning(name)
    return (f"[地名:{prefix or '—'}]"
            f"[商號:{shown or '!!空!!'}({len(brand)}字)]"
            f"[尾:{tail or '!!無!!'}]"
            + (f"   ⚠ {warn}" if warn else ""))


def inspect_only(src: Path, wb, real_names, numeric_cols, ratio_cols, name_map, full_path: Path) -> int:
    """Report what the transform WOULD see, twice: redacted here, full to disk.

    Everything printed to the terminal is safe to paste to someone who must
    not see the client's data; the file written beside it is not.
    """
    safe: list[str] = []
    full: list[str] = []

    def emit(line: str = "", safe_line: str | None = None):
        full.append(line)
        safe.append(line if safe_line is None else safe_line)

    emit(f"=== INSPECT: {src.name} ===")
    emit(f"sheets: {len(wb.sheetnames)}")
    emit("sheet names:", "sheet names (redacted):")
    for name in wb.sheetnames:
        emit(f"   {name}", f"   {redact(name)}")

    with zipfile.ZipFile(src) as z:
        media = [n for n in z.namelist() if n.startswith(("xl/media/", "xl/charts/"))]
    emit(f"\nembedded media/charts: {len(media)}  "
         f"{'(WILL BE LOST by openpyxl)' if media else ''}")

    formulas, merged, precise, rounded = [], 0, 0, 0
    money_units: dict[str, int] = defaultdict(int)
    money_samples: list[str] = []
    near_miss: list[tuple[str, str, str]] = []
    short_cn: dict[str, int] = defaultdict(int)
    subtotal_rows: dict[str, list[str]] = defaultdict(list)

    for ws in wb.worksheets:
        merged += len(ws.merged_cells.ranges)
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str):
                    s = v.strip()
                    if s.startswith("="):
                        formulas.append(f"{ws.title}!{cell.coordinate}")
                        continue
                    if any(w in s.lower() for w in SUBTOTAL_WORDS) and len(s) <= 20:
                        subtotal_rows[ws.title].append(f"r{cell.row}:{s}")
                    if PURE_CN_RE.match(s):
                        short_cn[s] += 1
                    if s not in real_names and s not in GENERIC_NAMES and \
                            any(suf in s for suf in COMPANY_SUFFIXES):
                        reason = ("長度>30" if len(s) > 30 else
                                  "含數字" if any(c.isdigit() for c in s) else
                                  "含非中文字元")
                        near_miss.append((ws.title + "!" + cell.coordinate, s, reason))
                    for m in NUM_UNIT_PROBE.finditer(s):
                        money_units[m.group(1)] += 1
                    for m in MONEY_RE.finditer(s):
                        money_samples.append(m.group(0))
                elif isinstance(v, float) and not isinstance(v, bool):
                    if v == round(v, 2):
                        rounded += 1
                    else:
                        precise += 1

    emit(f"merged cell ranges: {merged}")
    emit(f"formula cells: {len(formulas)}"
         + ("   <-- CRITICAL: writing values would DESTROY these" if formulas else ""))
    for f in formulas[:10]:
        emit(f"   {f}")

    total_float = precise + rounded
    emit(f"\nfloat precision: {precise} full-precision / {total_float} floats "
         f"({100 * precise / total_float:.0f}% carry >2dp)")
    emit("  -> full precision means detail sums tie to subtotals EXACTLY;"
         if precise > rounded else
         "  -> mostly 2dp: subtotals may carry rounding drift (same as the original).")

    emit(f"\n--- company names cleanly extracted: {len(real_names)} ---")
    for n in sorted(real_names):
        emit(f"   {n}   ->   {name_map.get(n, '(不變)')}",
             f"   {redact_company(n)}")

    emit(f"\n--- NEAR MISSES: contain a company suffix but were NOT extracted: {len(near_miss)} ---")
    emit("    (long remark sentences here are FINE - full-text substitution covers them.")
    emit("     A short one with no punctuation = a name format I am failing to catch.)")
    for loc, s, reason in near_miss[:40]:
        emit(f"   [{reason}] {loc}: {s[:60]}",
             f"   [{reason}] {loc}: len={len(s)} {redact(s[:40])}")
    if len(near_miss) > 40:
        emit(f"   ... and {len(near_miss) - 40} more")

    emit(f"\n--- number+unit tokens seen (to check the money regex covers them) ---")
    for unit, count in sorted(money_units.items(), key=lambda kv: -kv[1])[:30]:
        covered = "SCALED" if unit in ("w", "W", "万", "元") or unit.startswith(("w元", "万元")) else "not scaled"
        emit(f"   '{unit}' x{count}   [{covered}]")
    emit(f"   money regex currently matches {len(money_samples)} token(s); samples:")
    emit(f"      {', '.join(money_samples[:15])}",
         f"      {', '.join(redact(m) for m in money_samples[:15])}")

    emit(f"\n--- ratio columns (left unscaled) ---")
    for sheet, cols in sorted(ratio_cols.items()):
        emit(f"   {sheet}: {', '.join(sorted(cols))}",
             f"   {redact(sheet)}: {', '.join(sorted(cols))}")

    emit(f"\n--- subtotal/total/check rows per sheet (the subtable structure) ---")
    for sheet, rows in sorted(subtotal_rows.items()):
        emit(f"   {sheet}: {'; '.join(rows[:12])}",
             f"   {redact(sheet)}: {'; '.join(redact(r) for r in rows[:12])}")

    emit(f"\n--- short pure-Chinese strings (2-8 chars), {len(short_cn)} distinct ---")
    emit("    These are where tenant SHORT FORMS and PERSON NAMES hide.")
    emit("    Redacted terminal output cannot show them - check the full file yourself.")
    for s, c in sorted(short_cn.items(), key=lambda kv: -kv[1])[:60]:
        full.append(f"   x{c:<4} {s}")

    full_path.write_text("\n".join(full) + "\n", encoding="utf-8")
    print("\n".join(safe))
    print(f"\n\n[full, UNREDACTED report written to: {full_path}]")
    print("[the terminal output above is redacted and safe to share; the file is NOT]")
    return 0


# -------------------------------------------------------------- verification


def verify(src: Path, out: Path, factor: float, name_map, extra_map, ratio_cols) -> int:
    """Prove linearity cell-by-cell, then hunt for leaked real names.

    Linearity is the whole argument for the totals still tying out, so it is
    checked directly rather than by re-adding the schedules: if every numeric
    cell equals source*factor, then every sum of them does too.
    """
    print("\n=== VERIFY ===")
    wb_src = openpyxl.load_workbook(src, data_only=False)
    wb_out = openpyxl.load_workbook(out, data_only=False)

    src_sheets = wb_src.sheetnames
    out_sheets = wb_out.sheetnames
    if len(src_sheets) != len(out_sheets):
        print(f"  FAIL: sheet count {len(src_sheets)} -> {len(out_sheets)}")
        return 1

    checked = mismatches = 0
    worst = 0.0
    for s_name, o_name in zip(src_sheets, out_sheets):
        ws_s, ws_o = wb_src[s_name], wb_out[o_name]
        sheet_ratio = ratio_cols.get(s_name, set())
        for row in ws_s.iter_rows():
            for cell in row:
                v = cell.value
                if not isinstance(v, (int, float)) or isinstance(v, bool):
                    continue
                got = ws_o[cell.coordinate].value
                expect = v if cell.column_letter in sheet_ratio else v * factor
                checked += 1
                if got is None:
                    mismatches += 1
                    continue
                denom = max(abs(expect), 1e-9)
                rel = abs(got - expect) / denom
                worst = max(worst, rel)
                if rel > 1e-9:
                    mismatches += 1
                    if mismatches <= 5:
                        print(f"  MISMATCH {o_name}!{cell.coordinate}: got {got!r}, expected {expect!r}")

    print(f"  linearity: {checked} numeric cells checked, {mismatches} mismatch(es), "
          f"worst relative error {worst:.2e}")
    if mismatches == 0:
        print("  PASS -- every figure is exactly source x factor, so every subtotal,")
        print("         total, Check row and Financials tie-out is preserved by construction.")

    leaked = defaultdict(list)
    real_tokens = list(name_map) + list(extra_map)
    for ws in wb_out.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str):
                    continue
                for tok in real_tokens:
                    if tok in cell.value:
                        leaked[tok].append(f"{ws.title}!{cell.coordinate}")

    if leaked:
        print(f"\n  FAIL: {len(leaked)} real name(s) still present in the output:")
        for tok, locs in leaked.items():
            print(f"      {tok}: {', '.join(locs[:5])}{' ...' if len(locs) > 5 else ''}")
    else:
        print(f"  name leak: none of the {len(real_tokens)} mapped names survive in the output.")

    # Character-level residue: catches typo variants of a real brand
    # ("上海卲奥" for 上海邵奥) that literal substitution cannot reach.
    brand_chars: set[str] = set()
    for real in name_map:
        _, brand, _ = split_company(real)
        brand_chars.update(brand)
    residue = defaultdict(list)
    for ws in wb_out.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str) or len(cell.value) < 2:
                    continue
                for ch in brand_chars:
                    if ch in cell.value:
                        residue[ch].append(f"{ws.title}!{cell.coordinate}")

    if residue:
        print(f"\n  REVIEW BY EYE: {len(residue)} character(s) from real brand names still occur.")
        print("  Most are ordinary Chinese and harmless; look for a misspelt company name.")
        for ch, locs in sorted(residue.items(), key=lambda kv: -len(kv[1]))[:15]:
            print(f"      '{ch}' x{len(locs)}: {', '.join(locs[:3])}{' ...' if len(locs) > 3 else ''}")

    return 1 if (mismatches or leaked) else 0


if __name__ == "__main__":
    raise SystemExit(main())
