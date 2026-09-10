"""Smoke-test ask_run.py's loop with a SCRIPTED model -- zero tokens.

The agent loop has parts a real model cannot be trusted to exercise on demand:
JSON parsing around prose and code fences, a tool error and the retry after
it, the consulted-accounts bookkeeping, the grounding gate over the final
answer, the ⚠ marking and the exit status. A scripted model replies with a
fixed sequence per question, so each of those is hit deterministically.

    python ad-hoc/workbench/ask_run_smoke.py --run <replay or real run folder>

Five questions, mirroring the plan's M1 gate. The fifth deliberately puts an
invented figure in the answer to prove the gate refuses it.
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import argparse
import json
from typing import Any, Dict, List

import ask_run
from fdd_utils.ai.run_memory import RunMemory


class ScriptedClient:
    """Replies with the next scripted turn regardless of the prompt."""

    def __init__(self, turns: List[str]) -> None:
        self.turns = list(turns)
        self.calls = 0

    def get_response(self, user_prompt, system_prompt=None, **_kw) -> Dict[str, Any]:
        self.calls += 1
        if not self.turns:
            return {"content": '{"answer": "(script exhausted)"}'}
        return {"content": self.turns.pop(0)}


def run_case(memory: RunMemory, question: str, turns: List[str], expect_ok: bool,
             expect_tools: List[str]) -> bool:
    client = ScriptedClient(turns)
    import fdd_utils.ai.client as client_mod
    real = client_mod.AIClient
    client_mod.AIClient = lambda **kw: client  # type: ignore[assignment]
    try:
        result = ask_run.ask(memory, question, model_type="local", model_name=None,
                             max_steps=8, verbose=False)
    finally:
        client_mod.AIClient = real
    used = [t["tool"] for t in result["trace"] if t.get("tool")]
    ok_tools = all(t in used for t in expect_tools)
    ok = (result["ok"] == expect_ok) and ok_tools
    print("%s  %s" % ("PASS" if ok else "FAIL", question))
    print("      tools used: %s   consulted: %s   ok=%s (expected %s)"
          % (used, result["consulted"], result["ok"], expect_ok))
    if result["grounding"]["ungrounded"]:
        print("      ungrounded: %s" % [r["amount"] for r in result["grounding"]["ungrounded"]])
    if not ok:
        print("      answer: %r" % result["answer"][:200])
    return ok


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--run", required=True, help="a run folder with results.yml, audit.jsonl and evidence/")
    args = ap.parse_args()
    memory = RunMemory(ask_run._resolve_run(args.run))
    accounts = memory.accounts()
    first = next((k for k in accounts if memory.evidence.get(k) and memory.evidence[k].facts), accounts[0])
    real_cell = next(f["value"] for f in memory.evidence[first].facts
                     if f["kind"] in ("cell", "analysis_cell") and abs(float(f["value"])) >= 1000)
    tool = lambda name, **a: json.dumps({"tool": name, "args": a}, ensure_ascii=False)
    answer = lambda text: json.dumps({"answer": text}, ensure_ascii=False)

    cases = [
        ("哪些科目被標記為幻覺，原因是什麼？",
         [tool("run_health"), tool("flagged_clauses", key=first),
          answer("根據 audit.jsonl，%s 的 clause 檢查已完成；run health 見 results.yml。" % first)],
         True, ["run_health", "flagged_clauses"]),
        ("%s 的 %s 是從哪裡來的？" % (first, f"{real_cell:,.0f}"),
         # prose + code fence around the JSON, and a wrong-account error first
         ["Let me check.\n```json\n" + tool("trace_amount", key="不存在的科目", amount=real_cell) + "\n```",
          tool("trace_amount", key=first, amount=real_cell),
          answer("%s 的 %s元 來自 evidence/%s.json 的來源儲存格。" % (first, f"{real_cell:,.0f}", first))],
         True, ["trace_amount"]),
        ("這個 run 可以出貨嗎？",
         [tool("run_health"), answer("safe_to_export 見 results.yml:__run_health__。")],
         True, ["run_health"]),
        ("%s 的 prompt 給了模型什麼？" % first,
         [tool("what_was_shown", key=first), answer("見 evidence/%s.json 的 shown 欄。" % first)],
         True, ["what_was_shown"]),
        ("%s 的構成項加起來等於總數嗎？" % first,
         # the model INVENTS a figure the tools never returned -> the gate must refuse it
         [tool("series", key=first), tool("flagged_clauses", key=first),
          answer("%s 的合計為 987,654,321元，構成項加總一致。" % first)],
         False, ["flagged_clauses"]),
    ]
    results = [run_case(memory, q, turns, ok, tools) for q, turns, ok, tools in cases]
    print("\n%d/%d cases behaved as expected" % (sum(results), len(results)))
    return 0 if all(results) else 1


if __name__ == "__main__":
    sys.exit(main())
