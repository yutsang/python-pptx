#!/usr/bin/env python3
"""The evidence-grounded agent path, from the command line.

    python fdd_agent.py summary <databook>                        numerical summary, every figure traced     (free)
    python fdd_agent.py ask <databook> "question" [--base-run ID] investigate and answer, sources + states   (model)
    python fdd_agent.py edit <databook> --account KEY "instruction" --base-run ID [--export] [--out DECK]    (model)
    python fdd_agent.py diff <databook> --baseline-deck DECK      what the sources changed since that deck   (free)
    python fdd_agent.py deck <databook> --deck RETURNED [--update --out DECK]   read a returned deck         (free)
    python fdd_agent.py strip DECK --out OUTSIDE                  the copy of a deck without its manifest    (free)

One session, several turns (investigate -> follow up -> put it in the report -> accept -> export -> update a returned deck):

    python fdd_agent.py chat <databook> [--session FILE] [--base-run ID]      interactive; /help lists the commands (model on free text)
    python fdd_agent.py ask <databook> "question" --session FILE              one turn, kept in the session file
    python fdd_agent.py report <databook> --session FILE [--request "..."] [--account KEY] [--export] [--out DECK]
                                                                              "put that in the report": no model call unless the wording changes
    python fdd_agent.py candidates <databook> --session FILE                  the drafts, what they replace, what changed, what is open  (free)
    python fdd_agent.py accept <databook> --session FILE --account KEY        the user accepts a draft                               (free)
    python fdd_agent.py export <databook> --session FILE [--out DECK]         write the deck from accepted drafts                    (free)
    python fdd_agent.py session show FILE                                     what a saved session holds, without loading the workbook (free)

A session file holds the workbook's figures and commentary: it is created only when --session is given, only at a path git
would not track, and never leaves the machine. Without --session nothing is kept after the command.

`ask` and `edit` call a model and run only when the agent path is switched on
(FDD_AGENT=1, or agent.enabled: true in config.yml). The default is off, and the
Generator -> Auditor -> Validator pipeline (inspect_databook.py, the app) is untouched
either way; unsetting the flag is the rollback. The rest read workbooks and decks and
call no model.

--scripted FILE replaces the model with a JSON list of replies (one object per turn), for
replaying a conversation offline.

Exit status: 0 the run finished (answered / done / the free command succeeded); 2 it stopped
without an acceptable result (budget, no progress, a question for the user, an answer the
verifier would not accept) or the agent path is off; 1 an error. The report names accounts and
figures: keep it local.
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

RULE = "=" * 78


class ScriptedModel:
    """Replays a list of replies; running out of them is an error the controller reports."""

    def __init__(self, replies):
        self.replies, self.calls = list(replies), 0

    def __call__(self, system, user, max_tokens):
        if self.calls >= len(self.replies):
            raise RuntimeError("the scripted conversation ran out of replies")
        reply = self.replies[self.calls]
        self.calls += 1
        content = reply if isinstance(reply, str) else json.dumps(reply, ensure_ascii=False)
        return {"content": content, "prompt_tokens": None, "completion_tokens": None}


def _session(args):
    from fdd_utils.agent.session import SourceSession
    return SourceSession.load(args.databook, entity=args.entity or "", sheet=args.sheet, financials_from=args.financials_from,
                              financials_sheet=args.financials_sheet, base_run=getattr(args, "base_run", None))


def _model(args, session):
    if args.scripted:
        with open(args.scripted, encoding="utf-8") as fh:
            return ScriptedModel(json.load(fh))
    from fdd_utils.agent.controller import AIClientModel
    return AIClientModel(model_type=args.model or "local", model_name=args.model_name, language=session.language)


def _gate(args) -> bool:
    from fdd_utils.agent.flag import agent_enabled
    if agent_enabled():
        return True
    print("The agent path is off (agent.enabled: false, FDD_AGENT not set). The legacy pipeline is unaffected.\n"
          "To run this one time:  FDD_AGENT=1 python fdd_agent.py " + " ".join(sys.argv[1:]))
    return False


def _print_run(run, as_json: bool) -> int:
    if as_json:
        print(json.dumps(run.to_dict(), ensure_ascii=False, indent=2, default=str))
        return 0 if run.ok else 2
    print(RULE)
    print(f"{run.mode.upper()}  {run.request}")
    print(f"stop: {run.stop_reason}   tools: {' -> '.join(run.tool_path) or '(none)'}")
    print(RULE)
    if run.answer:
        print(run.answer.get("text") or run.answer.get("summary") or "")
        verification = run.answer.get("verification")
        if verification:
            print(f"\nverification: {verification['verified']} verified, {verification['inference']} marked as inference, {verification['failed']} failed")
        for source in run.answer.get("sources") or []:
            where = f"{source.get('sheet')}!row {source.get('row')}" if source.get("row") is not None else (source.get("sheet") or "")
            print(f"  [{source['ref']}] {source['account']} {source.get('period') or ''} {source.get('basis') or ''} {source.get('line') or ''}  {where}")
        for item in run.answer.get("unresolved") or []:
            print(f"  UNRESOLVED: {item}")
        for export in run.answer.get("exports") or []:
            print(f"  deck: {export['path']}  delivery: {export['delivery']['status']}  touched slides: {export['touched_slide_numbers']}")
    if run.question:
        print("QUESTION FOR YOU:", run.question)
    if run.report:
        r = run.report
        print("why it stopped:", r.get("why_stopped"))
        for g in r.get("gaps") or []:
            print("  open:", g)
        for n in r.get("next_steps") or []:
            print("  next:", n)
    u = run.usage
    print(f"\nmodel calls {u.get('model_calls')}  tokens est. {u.get('prompt_tokens_est')}+{u.get('completion_tokens_est')}"
          + (f"  provider reported {u.get('provider_prompt_tokens')}+{u.get('provider_completion_tokens')}" if u.get("provider_reported") else "  (provider reported none)")
          + f"  {u.get('seconds')}s")
    return 0 if run.ok else 2


def _work(args):
    from fdd_utils.agent.flag import agent_budget, agent_limits, agent_policy, parse_role
    from fdd_utils.agent.workspace import WorkSession
    model = None
    if getattr(args, "scripted", None):
        with open(args.scripted, encoding="utf-8") as fh:
            model = ScriptedModel(json.load(fh))
    extra = {}
    if getattr(args, "model", None):
        extra["model_type"] = args.model
    if getattr(args, "model_name", None):
        extra["model_name"] = args.model_name
    if getattr(args, "thinking", None):
        extra["thinking"] = args.thinking
    escalation = parse_role(args.escalation_model) if getattr(args, "escalation_model", None) else None
    overrides = dict(x.split("=", 1) for x in (getattr(args, "mapping_override", None) or []) if "=" in x) or None
    work = WorkSession.open(
        args.databook, session_path=getattr(args, "session", None), entity=args.entity or "", sheet=args.sheet,
        financials_from=args.financials_from, financials_sheet=args.financials_sheet, base_run=getattr(args, "base_run", None),
        mapping_overrides=overrides,
        model=model, budget=agent_budget(), limits=agent_limits(max_model_calls=getattr(args, "max_model_calls", None)),
        policy=agent_policy(mode=getattr(args, "routing", None)), escalation=escalation, **extra)
    if work.opening_notice:
        print("The workbook changed since this session was saved: " + work.opening_notice["summary"])
        for e in work.opening_notice["candidate_effects"]:
            print(f"  {e['account']}: {e['status']}" + (f" -- {'; '.join(e['reasons'])[:200]}" if e["reasons"] else ""))
    return work


def _print_turn(turn, as_json: bool = False) -> int:
    if as_json:
        print(json.dumps(turn.to_dict(), ensure_ascii=False, indent=2, default=str))
        return 0 if turn.ok else 2
    print(RULE)
    print(f"[{turn.n}] {turn.mode.upper()}  {turn.request}")
    print(f"stop: {turn.stop_reason}   tools: {' -> '.join(turn.tool_path) or '(none)'}")
    print(RULE)
    if turn.text:
        print(turn.text)
    if turn.verification:
        v = turn.verification
        print(f"\nverification: {v.get('verified')} verified, {v.get('inference')} marked as inference, {v.get('failed')} failed")
    for src in turn.sources:
        where = f"{src.get('sheet')}!row {src.get('row')}" if src.get("row") is not None else (src.get("sheet") or "")
        print(f"  [{src['ref']}] {src['account']} {src.get('period') or ''} {src.get('basis') or ''} {src.get('line') or ''}  {where}")
    for item in turn.unresolved:
        print(f"  UNRESOLVED: {item}")
    for note in turn.notes:
        print(f"  NOTE: {note}")
    if turn.question:
        print("QUESTION FOR YOU:", turn.question)
    if turn.report.get("why_stopped"):
        print("why it stopped:", turn.report["why_stopped"])
        for g in turn.report.get("gaps") or []:
            print("  open:", g)
    if turn.report.get("model_error"):
        print("model error:", turn.report["model_error"])
    for item in turn.report.get("promoted") or []:
        print(f"  draft for {item['account']}: {'validated' if item['ok'] else 'NOT validated (' + str(item.get('error')) + ')'}")
    export = turn.report.get("export")
    if export:
        _print_export(export)
    u = turn.usage or {}
    if u.get("model_calls") is not None and turn.mode != "promote":
        print(f"\nmodel calls {u.get('model_calls')}  tokens est. {u.get('prompt_tokens_est')}+{u.get('completion_tokens_est')}"
              + (f"  provider reported {u.get('provider_prompt_tokens')}+{u.get('provider_completion_tokens')}" if u.get("provider_reported") else "  (provider reported none)")
              + f"  {u.get('seconds')}s")
        for role, row in (u.get("by_role") or {}).items():
            print(f"  {role:<11} {row.get('model') or '?'}: {row['calls']} call(s), est. {row['prompt_est']}+{row['completion_est']}"
                  + (f", provider {row['provider_prompt']}+{row['provider_completion']}" if row.get("provider_prompt") else ", provider reported none"))
    routing = turn.report.get("routing") or {}
    if routing:
        print(f"routing: {routing.get('mode')}   stop class: {routing.get('stop_class')}   ended on: {routing.get('role_at_end')}")
        for e in routing.get("escalations") or []:
            print(f"  escalated at step {e['step']} to {e['model']}: {e['trigger']} ({', '.join(e['codes'])}); hand-off ~{e['handoff_tokens_est']} tokens")
    return 0 if turn.ok else 2


def _print_export(result: dict) -> None:
    if result.get("ok") is False and result.get("error") and not result.get("data"):
        print(f"  EXPORT REFUSED [{result['error']['code']}]: {result['error']['message']}")
        for step in ((result["error"].get("first_generation") or {}).get("steps") or []):
            print(f"    - {step['what']}\n      {step['command']}")
        return
    data = result.get("data") or {}
    err = result.get("error")
    print(f"  deck: {data.get('path')}   delivery: {(data.get('delivery') or {}).get('status')}   touched slides: {data.get('touched_slide_numbers')}"
          + (f"   [{err['code']}] {err['message'][:160]}" if err else ""))


def _print_candidates(views: list, unresolved: list) -> None:
    if not views:
        print("no drafts yet")
    for v in views:
        print(RULE)
        verif = v["verification"]
        print(f"{v['account']}  [{v['status'].upper()}]  from {v['origin']}{' (carried over a source change)' if v['carried'] else ''}   "
              f"{verif['verified']} verified, {verif['inference']} inference, {verif['failed']} failed, {verif['unverified']} unverified")
        if v["invalid_reason"]:
            print(f"  NO LONGER STANDS: {v['invalid_reason']}")
        print("  draft:   " + v["text"].replace("\n", "\n           "))
        if v["replaces"]:
            print("  replaces: " + v["replaces"][:400].replace("\n", "\n           "))
            print(f"  change:  {len(v['changes']['added'])} sentence(s) added, {len(v['changes']['removed'])} removed, {v['changes']['kept']} kept")
        else:
            print("  replaces: nothing (no commentary for this account in the base run)")
        for u in v["unresolved"]:
            print(f"  UNRESOLVED: {u}")
    shown = {x for v in views for x in v["unresolved"]}
    for u in unresolved:
        if u["message"] not in shown:
            print(f"  open ({u['from']}): {u['message']}")


def cmd_chat(args) -> int:
    if not _gate(args):
        return 2
    from fdd_utils.agent import tools as T
    from fdd_utils.agent.workspace import classify_request
    work = _work(args)
    interactive = sys.stdin.isatty()
    st = work.status()
    print(f"session {st['id']}  {st['workbook']}  source version {st['source_version']}  base run: {st['base_run'] or 'none (export needs one)'}")
    print("free text asks (or, if it says 'put that in the report', makes a draft); /help for commands.")
    help_text = ("/ask Q   /edit [@ACCOUNT] INSTRUCTION   /report [REQUEST]   /candidates   /unresolved   /accept [ACCOUNT]   /unaccept ACCOUNT\n"
                 "/export [PATH]   /refresh   /deck PATH   /update PATH [OUT]   /status   /save [PATH]   /quit\n"
                 "/mapping ACCOUNT   /override ACCOUNT SHEET   /clear-override ACCOUNT   /overrides   (a mapping is yours to change; the model cannot)\n"
                 "/components ACCOUNT [PERIOD] [OFFSET]   /base-run RUN_ID   /first-deck   /books")
    while True:
        try:
            line = input("> " if interactive else "").strip()
        except EOFError:
            break
        if not line:
            continue
        if not interactive:
            print(f"> {line}")
        try:
            if not line.startswith("/"):
                line = ("/report " if classify_request(line) in ("report", "report_edit") else "/ask ") + line
            cmd, _, rest = line[1:].partition(" ")
            rest = rest.strip()
            if cmd in ("quit", "exit", "q"):
                break
            elif cmd == "help":
                print(help_text)
            elif cmd == "ask":
                _print_turn(work.ask(rest))
            elif cmd == "edit":
                account = ""
                if rest.startswith("@"):
                    account, _, rest = rest[1:].partition(" ")
                _print_turn(work.edit(rest, account))
            elif cmd == "report":
                _print_turn(work.report(rest))
                _print_candidates(work.candidates_view(), work.unresolved())
            elif cmd == "candidates":
                _print_candidates(work.candidates_view(), work.unresolved())
            elif cmd == "unresolved":
                for u in work.unresolved():
                    print(f"  ({u['from']}) {u['account'] or 'workbook'}: {u['message']}")
            elif cmd == "accept":
                targets = [rest] if rest else [c["account"] for c in work.candidates_view() if c["status"] == "validated"]
                for t in targets:
                    print(f"  {t}: {'accepted' if work.accept(t) else 'cannot be accepted (no validated draft against the current sources)'}")
            elif cmd == "unaccept":
                print("  done" if work.unaccept(rest) else "  no such draft")
            elif cmd == "export":
                _print_export(work.export(out_path=rest))
            elif cmd == "refresh":
                r = work.refresh()
                print("  refreshed: " + (work._effects_line(r.get("candidate_effects")) if r.get("ok") else str(r.get("error"))))
            elif cmd == "deck":
                print(_deck_text(work.read_deck(rest)["data"], False))
            elif cmd == "update":
                path, _, out = rest.partition(" ")
                r = work.update_deck(path, out.strip())
                print(_deck_text(r["data"], True) if r.get("ok") else json.dumps(r.get("error"), ensure_ascii=False))
            elif cmd == "status":
                print(json.dumps(work.status(), ensure_ascii=False, indent=1, default=str))
            elif cmd == "mapping":
                r = T.call_tool(work.source, "explain_mapping", {"account": rest})
                print(json.dumps(r.data if r.ok else r.error, ensure_ascii=False, indent=1, default=str))
            elif cmd == "override":
                account, _, sheet = rest.partition(" ")
                if not account or not sheet.strip():
                    print("  usage: /override ACCOUNT SHEET")
                else:
                    r = work.set_mapping_override(account, sheet.strip())
                    print("  read again with the override: " + (work._effects_line(r.get("candidate_effects")) if r.get("ok") else str(r.get("error"))))
            elif cmd == "clear-override":
                r = work.clear_mapping_override(rest)
                print("  read again without it: " + (work._effects_line(r.get("candidate_effects")) if r.get("ok") else str(r.get("error"))))
            elif cmd == "overrides":
                print(json.dumps(work.mapping_overrides, ensure_ascii=False))
            elif cmd == "components":
                parts = rest.split()
                args_ = {"account": parts[0] if parts else ""}
                if len(parts) > 1:
                    args_["period"] = parts[1]
                if len(parts) > 2:
                    args_["offset"] = int(parts[2])
                r = T.call_tool(work.source, "get_account_components", args_)
                print(json.dumps(r.data if r.ok else r.error, ensure_ascii=False, indent=1, default=str))
            elif cmd == "base-run":
                r = work.set_base_run(rest)
                check = r.get("base_run_check") or {}
                print(f"  base run bound: check {check.get('status')}" + (f" -- {'; '.join(check.get('reasons') or [])[:300]}" if check.get("reasons") else ""))
            elif cmd == "first-deck":
                plan = work.first_generation_plan()
                print("  " + plan["why"])
                for step in plan["steps"]:
                    print(f"  - {step['what']}\n    {step['command']}")
            elif cmd == "books":
                print(json.dumps(work.shared.summary(), ensure_ascii=False, indent=1, default=str))
            elif cmd == "save":
                print("  saved to", work.save(rest or None))
            else:
                print(f"  unknown command /{cmd}; /help")
        except Exception as exc:                                    # noqa: BLE001
            print(f"  ERROR: {type(exc).__name__}: {exc}")
    st = work.status()
    u = st["usage"]
    print(f"\nmodel calls {u['model_calls']}  tokens est. {u['prompt_tokens_est']}+{u['completion_tokens_est']}"
          + (f"  provider reported {u['provider_prompt_tokens']}+{u['provider_completion_tokens']}" if u["provider_reported"] else "  (provider reported none)")
          + f"  {u['seconds']}s")
    print(f"session kept in {work.path}" if work.path else "session not kept (use --session FILE to keep it)")
    return 0


def cmd_ask_session(args) -> int:
    if not _gate(args):
        return 2
    work = _work(args)
    code = _print_turn(work.ask(args.question), args.json)
    if work.path:
        print(f"(session: {work.path})")
    return code


def cmd_report(args) -> int:
    if not _gate(args):
        return 2
    work = _work(args)
    code = _print_turn(work.report(args.request or "", args.account or "", export=args.export, out_path=args.out or ""), args.json)
    if not args.json:
        _print_candidates(work.candidates_view(), work.unresolved())
    return code


def cmd_candidates(args) -> int:
    work = _work(args)
    views = work.candidates_view()
    if args.json:
        print(json.dumps({"candidates": views, "unresolved": work.unresolved(), "status": work.status()}, ensure_ascii=False, indent=2, default=str))
    else:
        _print_candidates(views, work.unresolved())
    return 0


def cmd_accept(args) -> int:
    work = _work(args)
    ok = work.accept(args.account)
    print(f"{args.account}: {'accepted' if ok else 'cannot be accepted (no validated draft against the current sources)'}")
    return 0 if ok else 2


def cmd_export(args) -> int:
    work = _work(args)
    result = work.export(out_path=args.out or "")
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
    else:
        _print_export(result)
    return 0 if result.get("ok") else 2


def cmd_session(args) -> int:
    with open(args.file, encoding="utf-8") as fh:
        data = json.load(fh)
    if args.json:
        print(json.dumps(data, ensure_ascii=False, indent=2))
        return 0
    print(f"session {data['id']}  saved {data['saved']}  workbook {data['source']['workbook']}  source version {data['source']['source_version']}")
    print(f"model {data['model']}   turns {len(data['turns'])}   model calls {data['model_calls']}   usage {data['usage']}")
    for t in data["turns"]:
        print(f"  [{t['n']}] {t['mode']:7s} ok={t['ok']!s:5s} stale={t['stale']!s:5s} {t['request'][:70]}")
    for account, c in data["candidates"].items():
        state = "invalid" if c["invalid_reason"] else ("accepted" if c["accepted"] else "draft")
        print(f"  draft {account}: {state}  from {c['origin']}" + (f"  -- {c['invalid_reason'][:140]}" if c["invalid_reason"] else ""))
    print("(read from the file only; the workbook was not loaded, so whether it has changed since is not checked)")
    return 0


def cmd_summary(args) -> int:
    from fdd_utils.agent.tools import call_tool
    session = _session(args)
    result = call_tool(session, "numerical_summary")
    if args.json:
        print(json.dumps(result.to_dict(), ensure_ascii=False, indent=2, default=str))
    else:
        print(f"{os.path.basename(args.databook)}  source version {session.source_version}  statements: {session.binding.statement_source or '?'}")
        for line in (result.data.get("lines") or []):
            print(line)
        trace = result.data.get("traceability") or {}
        print(f"\ntraceability: {trace.get('traced')} of {trace.get('numbers_printed')} printed numbers resolve to a fact"
              + ("" if result.ok else f"  -- {result.error}"))
    return 0 if result.ok else 2


def cmd_ask(args) -> int:
    if getattr(args, "session", None):
        return cmd_ask_session(args)
    if not _gate(args):
        return 2
    from fdd_utils.agent.controller import AgentController
    from fdd_utils.agent.flag import agent_budget
    session = _session(args)
    controller = AgentController(session, _model(args, session), budget=agent_budget(), language=session.language)
    return _print_run(controller.run(args.question, mode="ask"), args.json)


def cmd_edit_session(args) -> int:
    if not _gate(args):
        return 2
    work = _work(args)
    code = _print_turn(work.edit(args.instruction, args.account or "", accept=args.export, export=args.export, out_path=args.out or ""), args.json)
    if not args.json:
        _print_candidates(work.candidates_view(), work.unresolved())
    return code


def cmd_edit(args) -> int:
    if getattr(args, "session", None):
        return cmd_edit_session(args)
    if not _gate(args):
        return 2
    from fdd_utils.agent.controller import AgentController
    from fdd_utils.agent.flag import agent_budget
    if not args.base_run:
        print("edit needs --base-run: the earlier run whose commentary the other accounts keep.")
        return 1
    if not args.account:
        print("edit needs --account KEY (or --session FILE, which takes the account of the last answer).")
        return 1
    session = _session(args)
    controller = AgentController(session, _model(args, session), budget=agent_budget(), language=session.language)
    return _print_run(controller.run(args.instruction, mode="edit", account=args.account, export=args.export), args.json)


def cmd_diff(args) -> int:
    from fdd_utils.agent.tools import call_tool
    session = _session(args)
    result = call_tool(session, "diff_sources", {"baseline_deck": args.baseline_deck})
    if args.json or not result.ok:
        print(json.dumps(result.to_dict(), ensure_ascii=False, indent=2, default=str))
        return 0 if result.ok else 2
    data = result.data
    print(f"diff mode: {data['mode']}   changed accounts: {data['changed_accounts'] or 'none'}")
    for account, d in (data.get("accounts") or {}).items():
        print(f"  {account}: {d['status']}")
        for c in d.get("changed", [])[:6]:
            print(f"      {c['id']}: {c['old']:,.2f} -> {c['new']:,.2f}")
    if data.get("note"):
        print(data["note"])
    return 0


def cmd_deck(args) -> int:
    from fdd_utils.agent.tools import call_tool
    session = _session(args)
    if args.update:
        result = call_tool(session, "update_deck", {"path": args.deck, "out_path": args.out or ""})
    else:
        result = call_tool(session, "read_deck", {"path": args.deck})
    print(json.dumps(result.to_dict(), ensure_ascii=False, indent=2, default=str) if args.json or not result.ok else _deck_text(result.data, args.update))
    return 0 if result.ok else 2


def _deck_text(data, updated: bool) -> str:
    if updated:
        lines = [f"updated copy: {data['out_path']}   {'COMPLETE' if data.get('complete') else 'PARTIAL'}   applied {len(data['applied'])}, "
                 f"left for a person {data['waiting_for_a_person']}   (the returned deck itself was not written)"]
        lines += [f"  unfinished {u.get('element') or u.get('account')}: {u.get('action')} {u.get('reason') or ''}" for u in data.get("unfinished") or []]
        lines += [f"  next: {step}" for step in data.get("next_steps") or []]
        return "\n".join(lines)
    lines = [f"mode: {data['mode']}   history: {data['history']}   source moved: {data.get('source_moved')}", data.get("note") or ""]
    for action in data.get("actions") or []:
        lines.append(f"  {action.get('element', '')}: {action['action']}  {action.get('reason') or ''}")
    if data.get("summary"):
        lines.append("summary: " + json.dumps(data["summary"], ensure_ascii=False))
    return "\n".join(x for x in lines if x)


def cmd_strip(args) -> int:
    from fdd_utils.pptx.lineage import strip_manifest
    out = strip_manifest(args.deck, args.out)
    print("written:", out)
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    def common(p, databook=True):
        if databook:
            p.add_argument("databook")
        p.add_argument("--entity", default=None)
        p.add_argument("--sheet", default=None, help="Financials sheet (default: auto-detect)")
        p.add_argument("--financials-from", default=None)
        p.add_argument("--financials-sheet", default=None)
        p.add_argument("--json", action="store_true")

    def model_args(p):
        p.add_argument("--model", default=None, help="local | workbench | dashscope | deepseek | openai (default: local, or the session's)")
        p.add_argument("--model-name", default=None)
        p.add_argument("--scripted", default=None, metavar="FILE", help="replay these replies instead of calling a model")
        p.add_argument("--max-model-calls", type=int, default=None, help="a hard cap on model calls over the whole session")
        p.add_argument("--routing", choices=("single", "escalate"), default=None,
                       help="single: one model answers (default; agent.routing). escalate: a second model takes over after repeated failures it could repair")
        p.add_argument("--escalation-model", default=None, metavar="PROVIDER::MODEL",
                       help="the second model (default: agent.models.escalation in config.yml); used only with --routing escalate")
        p.add_argument("--thinking", default=None, metavar="off|on|low|medium|high",
                       help="thinking for the first model; recorded as 'not sent' on provider branches that send no such parameter")

    def session_arg(p, required=False):
        p.add_argument("--session", default=None, required=required, metavar="FILE",
                       help="keep the thread in this file (figures and commentary: a path git does not track)")
        p.add_argument("--base-run", default=None)
        p.add_argument("--mapping-override", action="append", default=None, metavar="ACCOUNT=SHEET",
                       help="read ACCOUNT from SHEET (yours to set; repeatable; kept in the session file)")

    s = sub.add_parser("summary"); common(s); s.set_defaults(fn=cmd_summary)
    a = sub.add_parser("ask"); common(a); model_args(a); a.add_argument("question"); session_arg(a); a.set_defaults(fn=cmd_ask)
    e = sub.add_parser("edit"); common(e); model_args(e); e.add_argument("instruction")
    e.add_argument("--account", default=None, help="required without --session; with one, the account of the last answer")
    session_arg(e); e.add_argument("--out", default=None)
    e.add_argument("--export", action="store_true"); e.set_defaults(fn=cmd_edit)
    c = sub.add_parser("chat"); common(c); model_args(c); session_arg(c); c.set_defaults(fn=cmd_chat)
    r = sub.add_parser("report"); common(r); model_args(r); session_arg(r, required=True); r.add_argument("--request", default="")
    r.add_argument("--account", default=None); r.add_argument("--export", action="store_true"); r.add_argument("--out", default=None); r.set_defaults(fn=cmd_report)
    cs = sub.add_parser("candidates"); common(cs); session_arg(cs, required=True); cs.set_defaults(fn=cmd_candidates)
    ac = sub.add_parser("accept"); common(ac); session_arg(ac, required=True); ac.add_argument("--account", required=True); ac.set_defaults(fn=cmd_accept)
    ex = sub.add_parser("export"); common(ex); session_arg(ex, required=True); ex.add_argument("--out", default=None); ex.set_defaults(fn=cmd_export)
    ss = sub.add_parser("session"); ss.add_argument("action", choices=["show"]); ss.add_argument("file"); ss.add_argument("--json", action="store_true")
    ss.set_defaults(fn=cmd_session)
    d = sub.add_parser("diff"); common(d); d.add_argument("--baseline-deck", default="", help="a deck exported by this tool"); d.set_defaults(fn=cmd_diff)
    k = sub.add_parser("deck"); common(k); k.add_argument("--deck", required=True)
    k.add_argument("--update", action="store_true"); k.add_argument("--out", default=None); k.set_defaults(fn=cmd_deck)
    t = sub.add_parser("strip"); t.add_argument("deck"); t.add_argument("--out", default=None); t.set_defaults(fn=cmd_strip)
    args = ap.parse_args(argv)
    started = time.perf_counter()
    try:
        code = args.fn(args)
    except Exception as exc:                                        # noqa: BLE001
        print(f"ERROR: {type(exc).__name__}: {exc}")
        return 1
    if args.cmd in ("summary", "diff", "deck", "candidates", "accept", "export", "session"):
        print(f"({time.perf_counter() - started:.1f}s, no model call)")
    return code


if __name__ == "__main__":
    sys.exit(main())
