"""Free; no model call. A review file a person or another model can read WITHOUT the verifier's verdicts.

    python ad-hoc/workbench/blind_review.py --databook <xlsx> --session <trial_<arm>_<stamp>.agent-session.json> [--out FILE]
    python ad-hoc/workbench/portfolio_trial.py --dir databooks --entities A,B --blind        the newest paid trial of each book

Per account: its sentences, numbered as the deck's marks number them (0, 1, 2 ... in the order written), and for each the workbook
facts and notes the sentences cite. Nothing else: no colour, no refusal, no flag, no count of what the verifier passed. The trial's
own side-by-side file (trial_quality_<stamp>.md) carries all of those, so it is not a blind file; a reviewer who sees which
sentences are marked finds the marked ones. The numbering lets the reviewer's list of problem sentences be laid against the marks
afterwards (recheck_session.py / trial_digest.py print the marks by account and number). The file holds client figures and sentences:
local only.
"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
import tempfile
from pathlib import Path
from types import SimpleNamespace
from typing import List

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ROOT))

INSTRUCTIONS = """# 逐句來源審閱（盲評）

你是財務盡職調查報告的審閱人。下面每個科目有一段中文評論，已逐句編號（從 0 開始），和這段評論引用的工作簿原文。
原文的編號：F 是工作簿裡讀到的數字（行名、期間、金額，金額單位是元），D 是由 F 算出來的（差額、變動、占比，附公式），N 是附註或說明的原文。
方括號裡是口徑：raw 是該分頁自己的數字，financials 是 Financials 報表的數字，adjusted 是報告採用的科目合計（有取代時以 Financials 為準）。
評論裡的金額單位是萬元。

逐句檢查每一句，只依據下面給的原文判斷，不要補充你自己知道的背景。只列有問題的句子，每個問題寫：
1. 科目和句子編號，句子原文。
2. 類別：數字、期間、正負或方向、口徑（調整前後、明細與淨額、重分類）、沒有依據的原因、讀錯附註、主體（寫成了別的公司或項目）、其他。
3. 嚴重程度：重大（讀者會因此得到錯的數字或錯的結論）或輕微（措辭、精確度或不影響結論）。
4. 依據：引用的編號和它實際寫的內容，一句話。
5. 把握：確定，或爭議（要看工作簿原檔才能判斷）。
沒有問題的科目寫「無」。不要寫改寫建議。這份材料沒有任何機器檢查結果，不要假設有。
"""


def _num(value: object) -> str:
    return f"{value:,.2f}" if isinstance(value, (int, float)) and not isinstance(value, bool) else str(value or "")


def blind(databook: str, session: str) -> str:
    from entity_trial import sources_for_review
    from fdd_utils.agent.workspace import WorkSession
    saved = json.loads(Path(session).read_text(encoding="utf-8"))
    out: List[str] = [INSTRUCTIONS]
    with tempfile.TemporaryDirectory(prefix="blind_") as tmp:
        copy = Path(tmp) / Path(session).name                  # open() would save into the file it resumed from
        shutil.copyfile(session, copy)
        work = WorkSession.open(databook, session_path=str(copy), resume=True)
        work.path = None
        case = 0
        for key, cand in (saved.get("candidates") or {}).items():
            claims = list(cand.get("claims") or [])
            if not claims:
                continue
            case += 1
            src = sources_for_review(work.source, key, SimpleNamespace(claims=claims))
            out.append(f"## 科目 {case}：{key}")
            out.append("")
            out.extend(f"[{n}] {str(c.get('text') or '').strip()}" for n, c in enumerate(claims))
            out.append("")
            out.append("引用的原文：")
            for x in src.get("cited") or []:
                if x.get("basis") == "note":
                    continue                                    # a note is listed once below
                period = x.get("period")
                out.append(f"- {x['ref']}：{', '.join(period) if isinstance(period, list) else period or ''} "
                           f"{x.get('line') or '（科目合計）'}［{x.get('basis') or ''}］ "
                           f"{_num(x.get('value')) if isinstance(x.get('value'), (int, float)) else ''}"
                           f"{(' = ' + x['formula']) if x.get('formula') else ''}"
                           f"{(' （' + x['where'] + '）') if x.get('where') else ''} {x.get('text') or ''}".rstrip())
            notes = src.get("all_notes") or []
            if notes:
                out.append("該科目的附註與說明：")
                out.extend(f"- {n['ref']}：{n.get('text')}" for n in notes)
            out.append("")
    return "\n".join(out)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--databook", required=True, help="the workbook the session was written on")
    ap.add_argument("--session", required=True, help="trial_<arm>_<stamp>.agent-session.json")
    ap.add_argument("--out", default="", help="write here instead of printing")
    args = ap.parse_args()
    text = blind(args.databook, args.session)
    if args.out:
        Path(args.out).write_text(text, encoding="utf-8")
        print(f"wrote {args.out}")
    else:
        print(text)
    return 0


if __name__ == "__main__":
    sys.exit(main())
