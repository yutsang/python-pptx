"""Free; no model call. Sentences in the shapes a writer uses, built from a contract's own figures, and the verifier's verdict on each.

    python ad-hoc/workbench/claim_shapes.py <databook.xlsx> [--sheet NAME] [--show 12]
    python ad-hoc/workbench/claim_eval.py <databook.xlsx>          runs these too (CORRECT:shape_* and BROKEN:shape_* rows)

Why it exists. The verifier reads which period and which line each figure belongs to out of Chinese prose, with rules. Every new way
of writing a list, a bracket, a range or a make-up is a new chance for a rule to misread it, and a misread sentence is refused
three times and ends up marked. Three paid runs on six books each found about eight such misreadings, none of them the same as the
run before. This module finds them without a model: for every account of a book it writes correct sentences in each shape
(series with 分别, a series and a bracket after it, two lines on one list of periods, a period written straight before its amount,
a range of years, movements, a make-up and the lines put under one, the rest, a total and its bracket, a nil balance, a difference)
under each way of labelling a period, and the same sentences with the figures put under the wrong period. A correct sentence the
verifier refuses is a FALSE REJECT (a misreading to fix before anything is paid for); a wrong one it passes is a FALSE ACCEPT.
Run it on every new book first. It names accounts only as counts unless --show is given; the texts hold client figures.
"""
from __future__ import annotations

import argparse
import re
import sys
from collections import Counter, defaultdict
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parents[1]))

LABEL_STYLES = ("end", "date", "year")


def money(value: float, yi: bool = False) -> str:
    """A figure the way the report writes it: 万元 to one decimal (half up), 元 under 1万元; with yi, 亿元 to two decimals from 1亿."""
    value = abs(float(value))
    if yi and value >= 1e8:
        return f"{Decimal(str(value / 1e8)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP):,}亿元"
    if value < 10000:
        return f"{value:,.0f}元"
    return f"{Decimal(str(value / 1e4)).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP):,}万元"


def label(contract: Any, period: str, style: str) -> str:
    y, m, d = (int(x) for x in period.split("-"))
    if str(getattr(contract, "statement_type", "") or "").upper() == "IS":
        months = ((contract.period_meta or {}).get(period) or {}).get("months")
        months = int(months) if months else months
        return f"{y}年1-{months}月" if months and months < 12 else f"{y}年度"
    if style == "date":
        return f"{y}年{m}月{d}日"
    if style == "year":                 # "2024年312.2万元": the year written straight before its amount, no 末
        return f"{y}年" if (m, d) == (12, 31) else f"{y}年{m}月"
    return f"{y}年末" if (m, d) == (12, 31) else f"{y}年{m}月末"


def named(name: str, figure: str) -> str:
    """A line and its figure. A name that ends in a digit or a letter ("…户38900") cannot be followed by the figure with nothing between
    them: no reader, and no program, could tell where the name ends."""
    return f"{name}为{figure}" if re.search(r"[0-9A-Za-z]$", name) else f"{name}{figure}"


def joined(items: List[str], last: str = "和") -> str:
    return items[0] if len(items) == 1 else "、".join(items[:-1]) + last + items[-1]


class Shapes:
    def __init__(self, contract: Any, cl: Any):
        from fdd_utils.ai.evidence_contract import BASIS_ADJUSTED, BASIS_DERIVED, ROLE_COMPONENT, ROLE_TOTAL
        self.c, self.cl = contract, cl
        perm = contract.permissions
        self.ps = [p for p in contract.periods if (perm.get(p) or {}).get("describe_total", True)]
        self.total = {}
        for p in self.ps:
            f = next((x for x in contract.facts if x["period"] == p and x["role"] == ROLE_TOTAL and x["basis"] == BASIS_ADJUSTED), None)
            if f is not None:
                self.total[p] = f
        self.ps = [p for p in self.ps if p in self.total]
        by_line: Dict[str, Dict[str, List[Dict[str, Any]]]] = defaultdict(lambda: defaultdict(list))
        for f in contract.facts:
            if f["role"] == ROLE_COMPONENT and f.get("line") and f.get("basis") != BASIS_DERIVED and f["period"] in self.ps and not f.get("nested_in"):
                by_line[cl._bare(f["line"])][f["period"]].append(f)
        # a line that is one fact in every period it appears in, and positive
        self.lines = {name: {p: fs[0] for p, fs in per.items() if len(fs) == 1 and fs[0]["value"] > 0}
                      for name, per in by_line.items()}
        # lines that are gross of something (原值, 累计摊销, 减值准备) are no part of a balance a sentence says is made of them
        self.lines = {k: v for k, v in self.lines.items() if len(v) >= 2 and len(v) == len(by_line[k]) and not re.search(r"原值|累计|减值|折旧|摊销|准备", k)}
        self.detail_ok = lambda p: (perm.get(p) or {}).get("detail_basis") in ("tied", "ambiguous")

    def out(self, shape: str, text: str, periods: List[str], cites: List[str], kind: str = "amount",
            wrong: Optional[str] = None) -> Dict[str, Any]:
        return {"shape": shape, "claim": {"account": self.c.account, "kind": kind, "text": text, "periods": list(periods), "cites": list(cites)},
                "wrong": wrong}

    def lines_over(self, ps: List[str], limit: int = 2) -> List[str]:
        ok = [n for n, per in self.lines.items() if all(p in per for p in ps) and self.detail_ok_all(ps)]
        ok.sort(key=lambda n: -max(abs(per["value"]) for per in self.lines[n].values()))
        picked: List[str] = []
        for n in ok:
            if all(n not in q and q not in n for q in picked):
                picked.append(n)
            if len(picked) == limit:
                break
        return picked

    def detail_ok_all(self, ps: List[str]) -> bool:
        return all(self.detail_ok(p) for p in ps)

    @staticmethod
    def short(name: str) -> str:
        """A counterparty as a sentence shortens it: no code in front, nothing from the company form on."""
        core = re.sub(r"^[A-Za-z0-9_\-－]+", "", name)
        return re.split(r"实业|有限|股份|公司|集团|（", core)[0]

    def distinct(self, values: List[float]) -> bool:
        """The figures differ as WRITTEN (to one decimal in 万元): two figures that print alike cannot show a swap."""
        return len({money(v) for v in values}) == len(values) and len(values) >= 2

    def generate(self) -> List[Dict[str, Any]]:
        c, ps, out = self.c, self.ps, []
        if len(ps) < 2:
            return out
        use = ps[-3:]
        for style in LABEL_STYLES:
            if style != "end" and str(getattr(c, "statement_type", "") or "").upper() == "IS":
                continue
            lab = [label(c, p, style) for p in use]
            amts = [money(self.total[p]["value"]) for p in use]
            cites = [self.total[p]["ref"] for p in use]
            rev = amts[::-1]
            for last in ("和", "及"):
                head = f"目标公司{joined(lab, last)}{c.account}余额分别为"
                out.append(self.out(f"series_{last}", head + joined(amts, last), use, cites,
                                    wrong=head + joined(rev, last) if self.distinct([self.total[p]["value"] for p in use]) else None))
            if style == "end" and any(abs(self.total[p]["value"]) >= 1e8 for p in use):
                big = [money(self.total[p]["value"], yi=True) for p in use]
                head = f"{joined(lab)}{c.account}余额分别为"
                out.append(self.out("series_in_yi", head + joined(big), use, cites,
                                    wrong=head + joined(big[::-1]) if self.distinct([self.total[p]["value"] for p in use]) else None))
            names = self.lines_over(use, 2)
            if names:
                a = names[0]
                la = [money(self.lines[a][p]["value"]) for p in use]
                lc = [self.lines[a][p]["ref"] for p in use]
                head = f"{joined(lab)}{c.account}余额分别为{joined(amts)}"
                bad = self.distinct([self.lines[a][p]["value"] for p in use])
                out.append(self.out("series_then_bracket", f"{head}，其中{a}为主要项目（{joined(la)}）。", use, cites + lc,
                                    wrong=f"{head}，其中{a}为主要项目（{joined(la[::-1])}）。" if bad else None))
                out.append(self.out("labelled_series", f"{a}" + "、".join(f"{l}{x}" for l, x in zip(lab, la)) + "。", use, lc,
                                    wrong=(f"{a}" + "、".join(f"{l}{x}" for l, x in zip(lab, la[::-1])) + "。") if bad else None))
                if len(names) == 2:
                    b = names[1]
                    lb = [money(self.lines[b][p]["value"]) for p in use]
                    bc = [self.lines[b][p]["ref"] for p in use]
                    badb = self.distinct([self.lines[b][p]["value"] for p in use])
                    out.append(self.out("two_lines_one_list", f"{joined(lab)}{a}分别为{joined(la)}；{b}分别为{joined(lb)}。", use, lc + bc,
                                        wrong=f"{joined(lab)}{a}分别为{joined(la)}；{b}分别为{joined(lb[::-1])}。" if badb else None))
                    out.append(self.out("labelled_then_unlabelled_bracket",
                                        f"{c.account}中{a}为主要项目（" + "、".join(f"{l}{x}" for l, x in zip(lab, la)) + f"），{b}（{joined(lb)}）。",
                                        use, lc + bc,
                                        wrong=(f"{c.account}中{a}为主要项目（" + "、".join(f"{l}{x}" for l, x in zip(lab, la)) + f"），{b}（{joined(lb[::-1])}）。")
                                        if badb else None))
            # a range of year ends
            if len(use) >= 3 and style == "end" and all(p.endswith("12-31") for p in use):
                years = [int(p[:4]) for p in use]
                if years == list(range(years[0], years[0] + len(years))):
                    head = f"{years[0]}-{years[-1]}年各年末{c.account}余额分别为"
                    out.append(self.out("range_of_years", head + joined(amts), use, cites,
                                        wrong=head + joined(rev) if self.distinct([self.total[p]["value"] for p in use]) else None))
            if len(use) >= 3 and style == "end" and names and all(p.endswith("12-31") for p in use):
                years = [int(p[:4]) for p in use]
                if years == list(range(years[0], years[0] + len(years))):
                    a = names[0]
                    la = [money(self.lines[a][p]["value"]) for p in use]
                    head = f"{years[0]}-{years[-1]}年各年{named(a, '')}为"
                    out.append(self.out("range_then_run_of_amounts", head + joined(la, "、"), use, [self.lines[a][p]["ref"] for p in use],
                                        wrong=head + joined(la[::-1], "、") if self.distinct([self.lines[a][p]["value"] for p in use]) else None))
            # movements
            moves = []
            for p0, p1 in zip(ps, ps[1:]):
                d = self.cl.derive_movement(c, p0, p1)
                if d.get("valid") and d.get("value") is not None and abs(d["value"]) >= 10000 and p0 in self.total and p1 in self.total:
                    moves.append((p0, p1, d))
            for p0, p1, d in moves[-2:]:
                word = "增加" if d["value"] > 0 else "减少"
                flip = "减少" if d["value"] > 0 else "增加"
                out.append(self.out("movement", f"{label(c, p1, style)}较{label(c, p0, style)}{word}{money(d['value'])}", [p0, p1], [d["ref"]],
                                    kind="movement", wrong=f"{label(c, p1, style)}较{label(c, p0, style)}{flip}{money(d['value'])}"))
            if len(moves) >= 2 and moves[-2][1] == moves[-1][0]:
                (a0, a1, da), (b0, b1, db) = moves[-2], moves[-1]
                text = "，".join(f"{label(c, y1, style)}较{label(c, y0, style)}{'增加' if d['value'] > 0 else '减少'}{money(d['value'])}"
                                 for y0, y1, d in ((a0, a1, da), (b0, b1, db)))
                out.append(self.out("two_movements", text, [a0, a1, b1], [da["ref"], db["ref"]], kind="movement"))
        # a make-up, the lines put under one, the rest, a total with a bracket
        p = ps[-1]
        names = self.lines_over([p], 3)
        lab = label(c, p, "end")
        if len(names) >= 2:
            a, b = names[0], names[1]
            va, vb = self.lines[a][p]["value"], self.lines[b][p]["value"]
            if va + vb <= self.total[p]["value"] * 1.1:       # two lines that add up to more than the balance are no make-up of it
                out.append(self.out("make_up", f"{lab}{c.account}余额主要为{named(a, money(va))}、{named(b, money(vb))}", [p],
                                    [self.lines[a][p]["ref"], self.lines[b][p]["ref"]], kind="composition"))
            if (self.c.permissions.get(p) or {}).get("detail_basis") == "tied":
                rest = self.total[p]["value"] - va - vb
                if rest > 10000:
                    out.append(self.out("make_up_and_rest", f"{lab}{c.account}余额主要为{named(a, money(va))}、{named(b, money(vb))}，其余项目合计{money(rest)}", [p],
                                        [self.lines[a][p]["ref"], self.lines[b][p]["ref"]], kind="composition"))
            out.append(self.out("total_with_bracket", f"{lab}{a}、{b}合计{money(va + vb)}（{named(a, money(va))}、{named(b, money(vb))}）", [p],
                                [self.lines[a][p]["ref"], self.lines[b][p]["ref"]], kind="composition"))
        orgs = [n for n in self.lines if p in self.lines[n] and re.search(r"有限公司|公司", n) and len(self.short(n)) >= 3]
        if names and len(orgs) >= 2 and self.detail_ok(p):
            top, b, d = names[0], orgs[0], orgs[1]
            if top not in (b, d):
                out.append(self.out("parent_then_short_names",
                                    f"{lab}主要为{named(top, '')}，其中{self.short(b)}{money(self.lines[b][p]['value'])}、{self.short(d)}{money(self.lines[d][p]['value'])}",
                                    [p], [self.lines[top][p]["ref"], self.lines[b][p]["ref"], self.lines[d][p]["ref"]], kind="composition"))
        for f in c.facts:
            if f["period"] == p and f["role"] == "component" and f.get("value", 0) < -10000 and f.get("line") and not f.get("nested_in") and self.detail_ok(p):
                out.append(self.out("negative_line", f"{lab}{self.cl._bare(f['line'])}余额为负（-{money(f['value'])}）", [p], [f["ref"]], kind="composition"))
                break
        kids = defaultdict(list)
        for f in c.facts:
            if f.get("nested_in") and f["period"] == p and f["value"] > 0:
                kids[f["nested_in"]].append(f)
        for parent_ref, children in kids.items():
            parent = next((f for f in c.facts if f["ref"] == parent_ref), None)
            children = sorted(children, key=lambda f: -f["value"])[:2]
            if parent and len(children) == 2 and sum(f["value"] for f in children) <= parent["value"] and self.detail_ok(p):
                out.append(self.out("lines_under_a_line", f"{lab}其中{named(self.cl._bare(parent['line']), money(parent['value']))}，"
                                    f"主要为{named(self.cl._bare(children[0]['line']), money(children[0]['value']))}、"
                                    f"{named(self.cl._bare(children[1]['line']), money(children[1]['value']))}", [p],
                                    [parent["ref"], children[0]["ref"], children[1]["ref"]], kind="composition"))
                break
        # a balance that was nil, then was not
        for p0, p1 in zip(ps, ps[1:]):
            if abs(self.total[p0]["value"]) < 0.5 and self.total[p1]["value"] > 10000 and (c.permissions.get(p0) or {}).get("state_nil_or_first_arose", True) \
                    and self.cl.derive_movement(c, p0, p1).get("valid"):
                out.append(self.out("nil_then_not", f"{label(c, p0, 'date')}{c.account}余额为0元，{label(c, p1, 'date')}增至{money(self.total[p1]['value'])}",
                                    [p0, p1], [self.total[p0]["ref"], self.total[p1]["ref"]]))
                break
        # where tab and Financials disagree: the Financials figure and the difference
        for p in c.periods:
            d = self.cl.derive_difference(c, p)
            fin = next((f for f in c.facts if f["period"] == p and f.get("basis") == "financials" and f["role"] == "total"), None)
            if d.get("valid") and fin is not None and abs(d["value"]) >= 10000:
                shown = ("-" if fin["value"] < 0 else "") + money(fin["value"])
                out.append(self.out("difference_and_its_inputs", f"{label(c, p, 'date')}{c.account}按财务口径为{shown}，与底稿差异{money(d['value'])}",
                                    [p], [d["ref"]], kind="difference"))
                break
        return out


def evaluate(contract: Any, cl: Any, contracts: Optional[Dict[str, Any]] = None) -> List[Dict[str, Any]]:
    """Each shape's correct sentence and its wrong twin through the verifier: [{shape, kind: 'CORRECT'|'BROKEN', text, ok, codes}]."""
    rows = []
    book = contracts or {contract.account: contract}
    for item in Shapes(contract, cl).generate():
        v = cl.verify_claim(item["claim"], book)
        rows.append({"shape": item["shape"], "kind": "CORRECT", "text": item["claim"]["text"], "ok": v["verdict"] in ("verified", "verified_inference"),
                     "codes": sorted({f["code"] for f in v["findings"] if f["severity"] in ("failed", "review")})})
        if item.get("wrong"):
            w = cl.verify_claim(dict(item["claim"], text=item["wrong"]), book)
            rows.append({"shape": item["shape"], "kind": "BROKEN", "text": item["wrong"], "ok": w["verdict"] in ("verified", "verified_inference"),
                         "codes": sorted({f["code"] for f in w["findings"] if f["severity"] in ("failed", "review")})})
    return rows


def summarise(rows: List[Dict[str, Any]], show: int = 0) -> Tuple[List[str], int]:
    """(lines, number of faults): a correct sentence refused, a wrong one passed."""
    stats: Dict[str, Counter] = defaultdict(Counter)
    for r in rows:
        stats[f"{r['kind']}:{r['shape']}"]["n"] += 1
        stats[f"{r['kind']}:{r['shape']}"]["accepted" if r["ok"] else "rejected"] += 1
    lines, faults = [], 0
    for key in sorted(stats):
        c = stats[key]
        bad = c["rejected"] if key.startswith("CORRECT") else c["accepted"]
        faults += bad
        lines.append(f"  {key:46s} n={c['n']:3d}  accepted={c['accepted']:3d}  rejected={c['rejected']:3d}" + ("   <- FAULT" if bad else ""))
    wrong = [r for r in rows if (r["kind"] == "CORRECT") == (not r["ok"])]
    lines.append(f"SHAPES {len(rows)} sentence(s): {sum(1 for r in wrong if r['kind'] == 'CORRECT')} FALSE REJECT of "
                 f"{sum(1 for r in rows if r['kind'] == 'CORRECT')}, {sum(1 for r in wrong if r['kind'] == 'BROKEN')} FALSE ACCEPT of "
                 f"{sum(1 for r in rows if r['kind'] == 'BROKEN')}")
    for r in wrong[:show]:
        lines.append(f"   ! {r['kind']}:{r['shape']} | {r['text'][:110]} {r['codes']}")
    return lines, faults


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("workbook")
    ap.add_argument("--sheet", default=None)
    ap.add_argument("--show", type=int, default=0, help="print this many of the faulty sentences (they hold client figures)")
    args = ap.parse_args()
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    import claim_eval
    contracts = claim_eval.load_contracts(args.workbook, args.sheet)
    from fdd_utils.ai import claims as cl
    rows = [dict(r, account=key) for key, contract in contracts.items() for r in evaluate(contract, cl, {key: contract})]
    lines, faults = summarise(rows, args.show)
    print(f"{Path(args.workbook).name}: {len(contracts)} account(s)")
    print("\n".join(lines))
    return 1 if faults else 0


if __name__ == "__main__":
    sys.exit(main())
