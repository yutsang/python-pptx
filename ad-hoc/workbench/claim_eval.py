"""False accept / false reject / caveated share of the claim verifier, measured on real contracts.

    python ad-hoc/workbench/claim_eval.py <databook.xlsx> [--sheet NAME] [--json OUT]

Not a count of "unsupported clauses": a ground truth. For every account it writes
CORRECT claims from the contract (a balance, a movement, a composition) and then
breaks each one in one named way -- right amount from the wrong period, the
neighbour line's figure, a unit slip, a flipped direction, a percentage nudged,
the citation removed, a cause nothing supports. The verifier should accept the
first group and reject the second:

    false reject   a CORRECT claim (or a legitimate re-rounding of it) that fails
    false accept   a BROKEN claim that is verified
    caveated       verified, but on a basis it could not fully test (nested detail)

Zero tokens: the claims are written by code, so this measures the verifier, not a
model. What a model writes is measured the same way once a run produces claims.
Output names accounts only as counts; --json adds the failing texts, keep it local.
"""
from __future__ import annotations

import argparse
import collections
import json
import re
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))


def date_text(period: str) -> str:
    y, m, d = (int(x) for x in period.split("-"))
    return f"{y}年{m}月{d}日"


def wan(value: float, decimals: int = 1) -> str:
    return f"{abs(value) / 1e4:,.{decimals}f}万元"


def correct_claims(contract, cl):
    """[(name, claim, should_pass)] -- claims the contract supports."""
    from fdd_utils.ai.evidence_contract import BASIS_ADJUSTED, ROLE_COMPONENT, ROLE_TOTAL
    out = []
    periods = [p for p in contract.periods if contract.permissions.get(p, {}).get("describe_total")]
    if not periods:
        return out
    last = periods[-1]
    total = next((f for f in contract.facts if f["period"] == last and f["role"] == ROLE_TOTAL and f["basis"] == BASIS_ADJUSTED), None)
    if total is not None and abs(total["value"]) >= 1000:
        out.append(("balance", {"account": contract.account, "kind": "amount", "periods": [last], "cites": [total["ref"]],
                                "text": f"截至{date_text(last)}，{contract.account}余额{wan(total['value'])}"}, True))
        out.append(("balance_rounded", {"account": contract.account, "kind": "amount", "periods": [last], "cites": [total["ref"]],
                                        "text": f"截至{date_text(last)}，{contract.account}余额约{wan(total['value'], 0)}"}, True))
    if len(periods) >= 2:
        prev = periods[-2]
        t_prev = next((f for f in contract.facts if f["period"] == prev and f["role"] == ROLE_TOTAL and f["basis"] == BASIS_ADJUSTED), None)
        if total is not None and t_prev is not None and abs(total["value"] - t_prev["value"]) >= 1000:
            out.append(("two_periods", {"account": contract.account, "kind": "amount", "periods": [prev, last],
                                        "cites": [t_prev["ref"], total["ref"]],
                                        "text": f"{date_text(prev)}余额{wan(t_prev['value'])}，{date_text(last)}余额{wan(total['value'])}"}, True))
        d = cl.derive_movement(contract, prev, last)
        if d["valid"] and d["pct"] is not None and abs(d["value"]) >= 1000:
            word = "增加" if d["value"] > 0 else "减少"
            out.append(("movement", {"account": contract.account, "kind": "movement", "periods": [prev, last], "cites": [d["ref"]],
                                     "text": f"较{date_text(prev)}{word}{wan(d['value'])}，变动幅度{abs(d['pct']) * 100:.1f}%"}, True))
            # the same movement with the percentage signed, and with the direction as a rate word
            out.append(("movement_signed", {"account": contract.account, "kind": "movement", "periods": [prev, last], "cites": [d["ref"]],
                                            "text": f"较{date_text(prev)}{word}{wan(d['value'])}，变动幅度{'+' if d['value'] > 0 else '-'}{abs(d['pct']) * 100:.1f}%"}, True))
            out.append(("movement_rate", {"account": contract.account, "kind": "movement", "periods": [prev, last], "cites": [d["ref"]],
                                          "text": f"较{date_text(prev)}{'增长' if d['value'] > 0 else '下降'}{abs(d['pct']) * 100:.1f}%"}, True))
    perm = contract.permissions.get(last, {})
    if perm.get("detail_basis") in ("tied", "ambiguous"):
        comps = sorted((f for f in contract.facts if f["period"] == last and f["role"] == ROLE_COMPONENT and abs(f["value"]) >= 1000),
                       key=lambda f: -abs(f["value"]))[:2]
        if len(comps) == 2 and comps[0]["line"] != comps[1]["line"]:
            out.append(("composition", {"account": contract.account, "kind": "composition", "periods": [last],
                                        "cites": [c["ref"] for c in comps],
                                        "text": f"截至{date_text(last)}，" + "、".join(f"{c['line']}为{wan(c['value'])}" for c in comps)}, True))
    if len(periods) >= 2:
        prev = periods[-2]
        comps = sorted((f for f in contract.facts if f["period"] == last and f["role"] == ROLE_COMPONENT and f["line"] and abs(f["value"]) >= 1000),
                       key=lambda f: -abs(f["value"]))
        for f in comps[:3]:
            d1, d2 = cl.derive_share(contract, f["line"], prev), cl.derive_share(contract, f["line"], last)
            if d1.get("valid") and d2.get("valid") and abs(d1["pct"] - d2["pct"]) >= 0.01 and f["line"] != contract.account:
                others = [c["line"] for c in comps if c["line"] != f["line"] and c["line"] not in f["line"] and f["line"] not in c["line"]]
                out.append(("share_two_periods", {
                    "account": contract.account, "kind": "share", "periods": [prev, last], "cites": [d1["ref"], d2["ref"]],
                    "text": f"{date_text(prev)}{f['line']}占比{abs(d1['pct']) * 100:.1f}%，{date_text(last)}{f['line']}占比{abs(d2['pct']) * 100:.1f}%",
                    "_meta": {"line": f["line"], "p1": abs(d1["pct"]) * 100, "p2": abs(d2["pct"]) * 100, "prev": prev, "last": last,
                              "other": others[0] if others else None}}, True))
                word = "升至" if abs(d2["pct"]) > abs(d1["pct"]) else "降至"
                out.append(("share_direction", {
                    "account": contract.account, "kind": "share", "periods": [prev, last], "cites": [d1["ref"], d2["ref"]],
                    "text": f"{f['line']}占比由{date_text(prev)}的{abs(d1['pct']) * 100:.1f}%{word}{date_text(last)}的{abs(d2['pct']) * 100:.1f}%",
                    "_meta": {"p2": abs(d2["pct"]) * 100, "last": last, "line": f["line"]}}, True))
                break
    if len(periods) >= 2:
        # a line whose AMOUNT and SHARE move in opposite ways: each direction word belongs to its own quantity
        prev = periods[-2]
        for f in sorted((x for x in contract.facts if x["period"] == last and x["role"] == ROLE_COMPONENT and x["line"]
                         and abs(x["value"]) >= 1000 and x["line"] != contract.account), key=lambda x: -abs(x["value"])):
            f0 = next((x for x in contract.facts if x["period"] == prev and x["role"] == ROLE_COMPONENT and x["line"] == f["line"]), None)
            d1, d2 = cl.derive_share(contract, f["line"], prev), cl.derive_share(contract, f["line"], last)
            if f0 is None or not (d1.get("valid") and d2.get("valid")) or abs(f0["value"]) < 1000:
                continue
            amount_up, share_up = f["value"] > f0["value"], abs(d2["pct"]) > abs(d1["pct"])
            if amount_up != share_up and abs(f["value"] - f0["value"]) >= 1000 and abs(abs(d2["pct"]) - abs(d1["pct"])) >= 0.01:
                out.append(("amount_vs_share", {
                    "account": contract.account, "kind": "composition", "periods": [prev, last],
                    "cites": [f0["ref"], f["ref"], d1["ref"], d2["ref"]],
                    "text": (f"{f['line']}余额由{date_text(prev)}的{wan(f0['value'])}{'增至' if amount_up else '减至'}{date_text(last)}的{wan(f['value'])}，"
                             f"{f['line']}占比由{date_text(prev)}的{abs(d1['pct']) * 100:.1f}%{'升至' if share_up else '降至'}{date_text(last)}的{abs(d2['pct']) * 100:.1f}%")}, True))
                break
    return out


def flip(text, pair):
    """Swap the two words of `pair` (each occurrence of one becomes the other)."""
    a, b = pair
    return text.replace(a, "@@").replace(b, a).replace("@@", b)


def mutations(name, claim, contract, cl):
    """[(mutation name, mutated claim)] -- each one wrong in exactly one way."""
    out = []
    text, periods = claim["text"], claim["periods"]
    amount = re.search(r"([\d,]+(?:\.\d+)?)万元", text)
    other = [p for p in contract.periods if p not in periods]
    if other:
        o = other[0]
        out.append(("period_in_text", dict(claim, text=re.sub(r"\d{4}年\d{1,2}月\d{1,2}日", date_text(o), text))))
        out.append(("period_declared", dict(claim, periods=[o])))
    if amount and name in ("balance", "composition"):
        v = float(amount.group(1).replace(",", ""))
        out.append(("unit_slip_x100", dict(claim, text=text.replace(amount.group(0), f"{v * 100:,.1f}万元"))))
        out.append(("unit_slip_x0.01", dict(claim, text=text.replace(amount.group(0), f"{v / 100:,.2f}万元"))))
        out.append(("amount_off_5pct", dict(claim, text=text.replace(amount.group(0), f"{v * 1.05:,.1f}万元"))))
    if name == "share_direction":
        m = claim["_meta"]
        out.append(("share_word_flipped", dict(claim, text=flip(claim["text"], ("升至", "降至")))))
        out.append(("share_given_a_minus_sign", dict(claim, text=claim["text"].replace(f"{date_text(m['last'])}的{m['p2']:.1f}%",
                                                                                      f"{date_text(m['last'])}的-{m['p2']:.1f}%"))))
        return out
    if name == "amount_vs_share":
        out.append(("amount_word_flipped", dict(claim, text=flip(claim["text"], ("增至", "减至")))))
        out.append(("share_word_flipped_beside_amount", dict(claim, text=flip(claim["text"], ("升至", "降至")))))
        out.append(("both_words_flipped", dict(claim, text=flip(flip(claim["text"], ("增至", "减至")), ("升至", "降至")))))
        return out
    if name == "movement_signed":
        out.append(("percent_sign_flipped", dict(claim, text=re.sub(r"([+-])(\d+\.\d%)$", lambda m: ("-" if m.group(1) == "+" else "+") + m.group(2), claim["text"]))))
        return out
    if name == "movement_rate":
        out.append(("rate_word_flipped", dict(claim, text=flip(claim["text"], ("增长", "下降")))))
        return out
    if name == "share_two_periods":
        m = claim["_meta"]
        line, other = m["line"], m["other"]
        out.append(("shares_swapped", dict(claim, text=f"{date_text(m['prev'])}{line}占比{m['p2']:.1f}%，{date_text(m['last'])}{line}占比{m['p1']:.1f}%")))
        out.append(("share_called_growth", dict(claim, text=claim["text"].replace("占比", "增长"))))
        if other:
            out.append(("share_wrong_line", dict(claim, text=f"{date_text(m['prev'])}{line}占比{m['p1']:.1f}%，{date_text(m['last'])}{other}占比{m['p2']:.1f}%")))
        out.append(("share_period_dropped", dict(claim, text=f"{line}占比{m['p1']:.1f}%，{m['p2']:.1f}%")))
        return out
    if name == "two_periods":
        m = re.match(r"(.+?余额)([\d,]+(?:\.\d+)?万元)，(.+?余额)([\d,]+(?:\.\d+)?万元)$", text)
        if m:
            out.append(("periods_swapped", dict(claim, text=f"{m.group(1)}{m.group(4)}，{m.group(3)}{m.group(2)}")))
        out.append(("period_dropped_for_one", dict(claim, text=re.sub(r"^.+?余额", "余额", text, count=1))))
    if name == "movement":
        out.append(("direction_flipped", dict(claim, text=text.replace("增加", "@@").replace("减少", "增加").replace("@@", "减少"))))
        pct = re.search(r"([\d.]+)%", text)
        if pct:
            out.append(("percent_off", dict(claim, text=text.replace(pct.group(0), f"{float(pct.group(1)) + 5:.1f}%"))))
    if name == "composition":
        parts = claim["cites"]
        out.append(("line_uncited", dict(claim, cites=parts[:1])))
    if name == "balance":
        # the neighbouring period's figure, written as this period's
        fact_other = next((f for f in contract.facts if f["role"] == "total" and f["basis"] == "authority_adjusted"
                           and f["period"] in other and abs(f["value"] - next(x for x in contract.facts if x["ref"] == claim["cites"][0])["value"]) > 1e4), None)
        if fact_other is not None:
            out.append(("other_period_value", dict(claim, text=re.sub(r"[\d,]+(?:\.\d+)?万元", wan(fact_other["value"]), text, count=1))))
    out.append(("cites_removed", dict(claim, cites=[])))
    out.append(("unsupported_cause", dict(claim, text=text + "，主要系市场需求变化所致")))
    return out


def kind_sweep(claim, key, contract, cl, inference):
    """Verdicts of the same sentence under every legal kind. `kind` and `inference` are declared by whoever wrote the
    sentence; neither may change what is accepted."""
    return {k: cl.verify_claim(dict(claim, kind=k, inference=inference), {key: contract})["verdict"] for k in cl.CLAIM_KINDS}


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook")
    ap.add_argument("--sheet", default=None)
    ap.add_argument("--json", default=None)
    args = ap.parse_args()

    from fdd_utils.ai import claims as cl
    from fdd_utils.ai import evidence_contract as ec
    from fdd_utils.ai.facts import build_financials_by_key
    from fdd_utils.workbook import get_effective_mappings, load_mappings, process_workbook_data

    from fdd_utils.workbook.visibility import find_visible_financials_sheet
    sheet = args.sheet or find_visible_financials_sheet(args.workbook)   # a hidden sheet is never chosen
    state = process_workbook_data(temp_path=args.workbook, entity_name="", selected_sheet=sheet)
    dfs, bs_is = state["dfs"], state.get("bs_is_results") or {}
    mappings = get_effective_mappings(load_mappings(), state.get("resolution"))
    fin = {}
    for statement in ("balance_sheet", "income_statement"):
        for key, values in build_financials_by_key({statement: bs_is.get(statement)}, dfs, mappings).items():
            fin.setdefault(key, {}).update(values)
    digest = (state.get("resolution") or {}).get("workbook_digest")

    contracts = {}
    for key, df in dfs.items():
        contracts[key] = ec.build_contract(key, df, statement_type=str(((df.attrs.get("integrity") or {}).get("statement_type")) or ""),
                                           financials=fin.get(key), digest=digest)
    stats = collections.defaultdict(lambda: collections.Counter())
    sweep = collections.Counter()
    failures = []
    for key, contract in contracts.items():
        for name, claim, _ in correct_claims(contract, cl):
            v = cl.verify_claim(claim, {key: contract})
            ok = v["verdict"] in ("verified", "verified_inference")
            kind = "CORRECT:" + name
            stats[kind]["n"] += 1
            stats[kind]["accepted" if ok else "rejected"] += 1
            stats[kind]["caveated"] += bool(v.get("caveats"))
            if not ok:
                failures.append({"class": kind, "account": key, "text": claim["text"], "findings": [f["code"] for f in v["findings"] if f["severity"] == "failed"]})
            base = v["verdict"]
            sweep["correct_n"] += 1
            if any(x != base for x in kind_sweep(claim, key, contract, cl, False).values()):
                sweep["correct_verdict_changed_by_kind"] += 1
                failures.append({"class": "KIND:" + name, "account": key, "text": claim["text"], "findings": ["changes with kind"]})
            for mname, broken in mutations(name, claim, contract, cl):
                bv = cl.verify_claim(broken, {key: contract})
                accepted = bv["verdict"] in ("verified", "verified_inference")
                if mname != "unsupported_cause":        # inference=true legitimately turns an unsupported cause into a marked one
                    for inference in (False, True):
                        sweep["broken_n"] += 1
                        got = kind_sweep(broken, key, contract, cl, inference)
                        if any(x in ("verified", "verified_inference") for x in got.values()):
                            sweep["broken_accepted_under_some_kind"] += 1
                            failures.append({"class": f"KIND+INFERENCE={inference}:" + mname, "account": key, "text": broken["text"], "findings": []})
                stats["BROKEN:" + mname]["n"] += 1
                stats["BROKEN:" + mname]["accepted" if accepted else "rejected"] += 1
                if accepted:
                    failures.append({"class": "BROKEN:" + mname, "account": key, "text": broken["text"], "findings": []})
    print(f"{Path(args.workbook).name}: {len(contracts)} account(s)")
    fr_n = fr = fa_n = fa = cav = ok_n = 0
    for kind in sorted(stats):
        c = stats[kind]
        print(f"  {kind:34s} n={c['n']:3d}  accepted={c['accepted']:3d}  rejected={c['rejected']:3d}"
              + (f"  caveated={c['caveated']}" if c["caveated"] else ""))
        if kind.startswith("CORRECT:"):
            fr_n += c["n"]; fr += c["rejected"]; cav += c["caveated"]; ok_n += c["accepted"]
        else:
            fa_n += c["n"]; fa += c["accepted"]
    print(f"FALSE REJECT {fr}/{fr_n} ({100 * fr / max(fr_n, 1):.1f}%)   FALSE ACCEPT {fa}/{fa_n} ({100 * fa / max(fa_n, 1):.1f}%)   "
          f"CAVEATED {cav}/{ok_n} of the accepted correct claims ({100 * cav / max(ok_n, 1):.1f}%)")
    print(f"KIND SWEEP (8 legal kinds, inference off and on): correct claims whose verdict changes with kind "
          f"{sweep['correct_verdict_changed_by_kind']}/{sweep['correct_n']}   broken claims accepted under some kind or inference "
          f"{sweep['broken_accepted_under_some_kind']}/{sweep['broken_n']}")
    for item in failures[:12]:
        print("   !", item["class"], item["account"], "|", item["text"][:60], item["findings"])
    if args.json:
        Path(args.json).write_text(json.dumps({"stats": {k: dict(v) for k, v in stats.items()}, "failures": failures},
                                              ensure_ascii=False, indent=1), encoding="utf-8")
    return 1 if (fa or sweep["correct_verdict_changed_by_kind"] or sweep["broken_accepted_under_some_kind"]) else 0


if __name__ == "__main__":
    sys.exit(main())
