#!/usr/bin/env python3
"""Ask a finished run a question, and get an answer whose numbers are grounded.

One model, one tool loop, no planner, no second agent. The model sees the
run's account list and nine deterministic tools over what the run left on disk
(fdd_utils/ai/run_memory.py); it calls them by writing one JSON object per
turn and answers when it has enough. Every amount in the final answer is then
run through the SAME grounding pool that graded the deck -- an amount the pool
cannot find is marked ⚠ and the exit status is 2. That rule is what makes a
chat over financial data safe to hand to a reviewer: the answer is held to the
arithmetic the deck was held to.

    python ask_run.py "哪些科目被標記為幻覺，原因是什麼？"
    python ask_run.py --run 20260910_171836 "固定资产 的 0.06亿元 是從哪裡來的？"
    python ask_run.py --run 20260910_171836 --json "這個 run 可以出貨嗎？"

Defaults to the newest run under fdd_utils/logs. --model picks the provider
(local | workbench | deepseek | openai), same as inspect_databook.py. Costs a
handful of LLM calls (one per step, at most --max-steps); the tools are free.

Output names accounts and prints figures; do not put it anywhere public.
"""
from __future__ import annotations

import argparse
import glob
import json
import os
import re
import sys
from typing import Any, Dict, List, Optional, Tuple

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from fdd_utils.ai.run_memory import RunMemory, TOOLS, call_tool  # noqa: E402

LOG_ROOT = os.path.join("fdd_utils", "logs")
RULE = "=" * 78
_TOOL_RESULT_CHARS = 6000   # per tool result, into the transcript; the model asks again if it needs more


def _newest_run() -> Optional[str]:
    runs = sorted(glob.glob(os.path.join(LOG_ROOT, "run_*")))
    return runs[-1] if runs else None


def _resolve_run(arg: Optional[str]) -> str:
    if not arg:
        folder = _newest_run()
        if not folder:
            sys.exit("no runs under %s" % LOG_ROOT)
        return folder
    if os.path.isdir(arg):
        return arg
    folder = os.path.join(LOG_ROOT, "run_%s" % str(arg).replace("run_", ""))
    if not os.path.isdir(folder):
        sys.exit("no such run: %s" % folder)
    return folder


def _system_prompt(memory: RunMemory) -> str:
    lang = memory.language() or "Chi"
    tool_lines = "\n".join(
        "  - %s(%s): %s" % (name, ", ".join(spec["args"]), spec["help"]) for name, spec in TOOLS.items()
    )
    answer_lang = "繁體中文" if lang == "Chi" else "English"
    return (
        "You answer questions about ONE finished commentary run of a financial due diligence tool. "
        "You have no knowledge of the run except what the tools return. Rules:\n"
        "1. Use the tools. Never state a figure, a phase, a flag or a cause the tools did not return.\n"
        "2. Every amount you state must be one a tool returned, written with its unit as the tool "
        "gave it (base units are yuan; 6,000,000 may be written 600万元 or 0.06亿元).\n"
        "3. Name the account each fact belongs to, and mention the provenance the tool gave "
        "(e.g. 'audit.jsonl', 'evidence/固定资产.json').\n"
        "4. If a tool returns an error, correct the arguments once; if it still fails, say what "
        "you could not find. Do not guess.\n"
        "5. Answer in %s. Be brief and concrete.\n\n"
        "Run %s has %d accounts: %s\n\n"
        "Tools:\n%s\n\n"
        "PROTOCOL. Reply with exactly ONE JSON object and nothing else:\n"
        '  {"tool": "<name>", "args": {...}}   to call a tool\n'
        '  {"answer": "<your answer>"}          when you have enough\n'
        "trace_amount takes the amount in base units (yuan): 0.06亿元 -> 6000000, 26.9万元 -> 269000."
        % (answer_lang, memory.run_id, len(memory.accounts()), ", ".join(memory.accounts()), tool_lines)
    )


_JSON_OBJ = re.compile(r"\{.*\}", re.S)


def _parse_turn(text: str) -> Tuple[Optional[Dict[str, Any]], str]:
    """(parsed object, error). Tolerates prose around the JSON and code fences."""
    raw = str(text or "").strip()
    raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw, flags=re.S)
    m = _JSON_OBJ.search(raw)
    if not m:
        return None, "no JSON object in the reply"
    blob = m.group(0)
    # Try the largest balanced object first, then shrink from the right.
    for end in range(len(blob), 0, -1):
        try:
            obj = json.loads(blob[:end])
            if isinstance(obj, dict):
                return obj, ""
        except Exception:
            continue
    return None, "reply was not valid JSON"


def _truncate(obj: Any, limit: int = _TOOL_RESULT_CHARS) -> str:
    text = json.dumps(obj, ensure_ascii=False, default=str)
    if len(text) <= limit:
        return text
    return text[:limit] + " …[truncated %d chars; ask a narrower question]" % (len(text) - limit)


def ask(memory: RunMemory, question: str, *, model_type: str, model_name: Optional[str],
        max_steps: int, verbose: bool) -> Dict[str, Any]:
    from fdd_utils.ai.client import AIClient
    client = AIClient(model_type=model_type, agent_name="subagent_4",
                      language=memory.language() or "Chi", model_name=model_name)
    system = _system_prompt(memory)
    transcript: List[str] = ["QUESTION: %s" % question]
    consulted: List[str] = []
    trace: List[Dict[str, Any]] = []
    answer: Optional[str] = None
    for step in range(1, max_steps + 1):
        user = "\n\n".join(transcript) + "\n\nYour next JSON object:"
        response = client.get_response(user, system, temperature=0.1, max_tokens=1200, allow_thinking=False)
        reply = str((response or {}).get("content") or "")
        obj, err = _parse_turn(reply)
        if verbose:
            print("--- step %d model reply:\n%s\n" % (step, reply[:800]))
        if obj is None:
            transcript.append("SYSTEM: %s. Reply with one JSON object only." % err)
            trace.append({"step": step, "error": err, "reply": reply[:300]})
            continue
        if "answer" in obj:
            answer = str(obj.get("answer") or "").strip()
            trace.append({"step": step, "answer": True})
            break
        name = str(obj.get("tool") or "")
        args = obj.get("args") if isinstance(obj.get("args"), dict) else {}
        result = call_tool(memory, name, args)
        key = args.get("key")
        if key and "error" not in result:
            resolved = memory.resolve(str(key))
            if resolved and resolved not in consulted:
                consulted.append(resolved)
        trace.append({"step": step, "tool": name, "args": args,
                      "error": result.get("error") if isinstance(result, dict) else None})
        transcript.append("TOOL %s(%s) RETURNED:\n%s" % (name, json.dumps(args, ensure_ascii=False), _truncate(result)))
    if answer is None:
        answer = ""
    grounding = memory.ground_text(answer, consulted or None)
    return {
        "run_id": memory.run_id,
        "question": question,
        "answer": answer,
        "consulted": consulted,
        "steps": len(trace),
        "trace": trace,
        "grounding": grounding,
        "ok": bool(answer) and not grounding["ungrounded"],
    }


def _mark(answer: str, ungrounded: List[Dict[str, Any]]) -> str:
    out = answer
    for item in sorted(ungrounded, key=lambda r: r["span"][0], reverse=True):
        s, e = item["span"]
        out = out[:s] + "⚠" + out[s:e] + "⚠" + out[e:]
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("question")
    ap.add_argument("--run", default=None, help="run id or folder (default: newest)")
    ap.add_argument("--model", default="local", help="local | workbench | deepseek | openai")
    ap.add_argument("--model-name", default=None)
    ap.add_argument("--max-steps", type=int, default=8)
    ap.add_argument("--json", action="store_true", help="print the full result as JSON")
    ap.add_argument("--verbose", action="store_true", help="print each model reply")
    args = ap.parse_args()

    memory = RunMemory(_resolve_run(args.run))
    result = ask(memory, args.question, model_type=args.model, model_name=args.model_name,
                 max_steps=args.max_steps, verbose=args.verbose)

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
        return 0 if result["ok"] else 2

    print(RULE)
    print("RUN %s   %s" % (result["run_id"], memory.run_folder))
    print("Q: %s" % args.question)
    print(RULE)
    if not result["answer"]:
        print("(no answer within %d steps)" % args.max_steps)
    else:
        print(_mark(result["answer"], result["grounding"]["ungrounded"]))
    print()
    print("tools: " + " -> ".join(
        "%s(%s)%s" % (t.get("tool"), ", ".join(str(v) for v in (t.get("args") or {}).values()),
                      " ✗" if t.get("error") else "")
        for t in result["trace"] if t.get("tool")
    ) or "(none)")
    g = result["grounding"]
    if g["amounts"]:
        print("grounding: %d amount(s) in the answer, %d grounded against %s"
              % (len(g["amounts"]), len(g["amounts"]) - len(g["ungrounded"]), g["accounts"]))
    if g["ungrounded"]:
        print("⚠ NOT GROUNDED: " + ", ".join("%s" % r["amount"] for r in g["ungrounded"])
              + " -- the tools never returned these figures; do not repeat them")
    print(RULE)
    return 0 if result["ok"] else 2


if __name__ == "__main__":
    sys.exit(main())
