"""The agent workbench inside the app: ask, follow up, put it in the report, review the draft, accept, export.

Off unless the agent path is switched on (agent.enabled in config.yml, or FDD_AGENT=1): with it off this renders nothing
and the app is exactly what it was. On, it is a thin face over agent/workspace.WorkSession -- every rule (what carries
between turns, what a changed workbook invalidates, what an export refuses) lives there and is tested there; this file
only shows it and passes clicks through.

The session lives in st.session_state for as long as the browser tab does. Nothing is written to disk unless the user
presses "Save session", and then only to fdd_utils/output/agent_sessions/ (git-ignored).
"""
from __future__ import annotations

import os
from typing import Any, Callable, Dict, Optional

import streamlit as st

from ..agent.flag import agent_enabled, agent_limits, agent_policy

KEY = "agent_work"
TURN_KEY = "agent_last_turn"
EXPORT_KEY = "agent_last_export"


def _open_work(session_state: Any):
    """A WorkSession over what the app has loaded: the same workbook, entity, Financials source and base run."""
    from ..agent.workspace import WorkSession
    rollup = session_state.get("rollup_temp_path") if session_state.get("rollup_sheet") else None
    return WorkSession.open(
        session_state.get("temp_path"), entity=session_state.get("entity_name") or "",
        sheet=None if rollup else session_state.get("selected_sheet"),
        financials_from=rollup, financials_sheet=session_state.get("rollup_sheet") if rollup else None,
        base_run=session_state.get("last_run_folder"),
        model_type=session_state.get("model_type", "local"), model_name=session_state.get("model_name"),
        limits=agent_limits(), policy=agent_policy())


def _show_turn(turn: Any) -> None:
    if turn is None:
        return
    head = f"Turn {turn.n} ({turn.mode})"
    if turn.ok:
        st.success(head)
    else:
        st.warning(f"{head}: {turn.stop_reason or 'did not finish'}")
    if turn.text:
        st.markdown(turn.text)
    v = turn.verification or {}
    if v:
        st.caption(f"{v.get('verified', 0)} verified, {v.get('inference', 0)} marked as inference, {v.get('failed', 0)} failed")
    if turn.question:
        st.info(turn.question)
    for note in turn.notes:
        st.caption(note)
    routing = turn.report.get("routing") or {}
    for e in routing.get("escalations") or []:
        st.caption(f"The second model ({e['model']}) took over at step {e['step']}: {e['trigger']}.")
    if routing and not turn.ok and routing.get("stop_class") not in (None, "ok"):
        st.caption("Why it stopped: " + str(routing["stop_class"]).replace("_", " "))
    if turn.report.get("model_error"):
        st.error(f"model: {turn.report['model_error']}")
    elif turn.report.get("why_stopped"):
        st.caption("why it stopped: " + str(turn.report["why_stopped"]))
    if turn.sources:
        st.dataframe([{"ref": s["ref"], "account": s["account"], "period": s.get("period"), "basis": s.get("basis"), "line": s.get("line"),
                       "where": f"{s.get('sheet')}!row {s.get('row')}" if s.get("row") is not None else s.get("sheet")} for s in turn.sources],
                     use_container_width=True, hide_index=True)
    for item in turn.unresolved:
        st.warning("Unresolved: " + item)


def _show_candidates(work: Any, can_export: bool) -> None:
    views = work.candidates_view()
    if not views:
        st.caption("No draft yet. Ask a question, then use 'Put into report'.")
        return
    for view in views:
        status = view["status"]
        label = {"accepted": "ACCEPTED", "validated": "draft, checked", "invalid": "NO LONGER STANDS", "failed": "FAILED CHECKS"}[status]
        if view.get("needs_review") and status == "validated":
            label = "still verifies, but the account moved: look again"
        with st.container(border=True):
            st.markdown(f"**{view['account']}** — {label}")
            if view["invalid_reason"]:
                st.error(view["invalid_reason"])
            for change in view.get("context_changes") or []:
                st.warning("The account moved around this draft: " + change)
            st.markdown(view["text"])
            v = view["verification"]
            st.caption(f"{v['verified']} verified, {v['inference']} inference, {v['failed']} failed, {v['unverified']} unverified"
                       + (" · carried over a change in the workbook" if view["carried"] else ""))
            if view["replaces"]:
                # shown inline: this panel is itself one expander, and Streamlit does not allow an expander inside an expander
                st.markdown(f"**What it replaces** ({len(view['changes']['added'])} added, {len(view['changes']['removed'])} removed, {view['changes']['kept']} kept)")
                st.markdown(view["replaces"])
            else:
                st.caption("Replaces nothing: the base run has no commentary for this account.")
            for item in view["unresolved"]:
                st.warning("Unresolved: " + item)
            cols = st.columns(2)
            accept_label = "Accept again, knowing what moved" if view.get("needs_review") else "Accept this draft"
            if status == "validated" and cols[0].button(accept_label, key=f"agent_accept_{view['account']}"):
                work.accept(view["account"])
                st.rerun()
            if status == "accepted" and cols[0].button("Take back acceptance", key=f"agent_unaccept_{view['account']}"):
                work.unaccept(view["account"])
                st.rerun()


def _show_export(result: Optional[Dict[str, Any]]) -> None:
    if not result:
        return
    data = result.get("data") or {}
    error = result.get("error")
    if result.get("ok") and data.get("path"):
        status = (data.get("delivery") or {}).get("status")
        st.success(f"Deck written ({status}); slides touched: {data.get('touched_slide_numbers')}")
        if os.path.exists(data["path"]):
            with open(data["path"], "rb") as fh:
                st.download_button("Download the deck", fh.read(), file_name=os.path.basename(data["path"]),
                                   mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", key="agent_download")
    elif error:
        st.error(f"[{error['code']}] {error['message']}")
        if data.get("path"):
            st.caption(f"A draft was still written: {os.path.basename(data['path'])}")
    for reason in ((data.get("delivery") or {}).get("findings") or [])[:6]:
        if reason.get("severity") in ("blocked", "review"):
            st.caption(f"[{reason['severity']}] {reason.get('account') or 'deck'}: {reason['message']}")


def render_agent_panel(session_state: Any, factory: Optional[Callable[[Any], Any]] = None) -> None:
    if not agent_enabled():
        return
    if not session_state.get("temp_path") or not os.path.exists(str(session_state.get("temp_path"))):
        return
    with st.expander("🧭 Agent workbench: investigate, put it in the report, export", expanded=bool(session_state.get(KEY))):
        work = session_state.get(KEY)
        if work is None:
            st.caption("An investigation session keeps what you asked, what the answers rested on, and the drafts you accept, until you close the tab. "
                       "Questions are answered by the model selected in the sidebar and are billed by the provider.")
            if not session_state.get("last_run_folder"):
                st.warning("There is no AI run for this workbook yet, so a deck made here would hold only the accounts you edit. "
                           "Generate content first, or accept that an export will be refused.")
            if st.button("Start a session on this workbook", key="agent_start"):
                with st.spinner("Reading the workbook…"):
                    try:
                        session_state[KEY] = (factory or _open_work)(session_state)
                    except Exception as exc:                                # noqa: BLE001
                        st.error(f"{type(exc).__name__}: {exc}")
                        return
                st.rerun()
            return
        status = work.status()
        st.caption(f"Session {status['id']} · source version {status['source_version']} · base run {status['base_run'] or 'none'} · "
                   f"model calls {status['model_calls']}" + (f" of {status['max_model_calls']}" if status["max_model_calls"] else "")
                   + f" · tokens charged {status['budget']['charged_tokens']:,}"
                   + (f" of {status['budget']['limits']['max_tokens']:,}" if status["budget"]["limits"].get("max_tokens") else "")
                   + f" · routing {status['routing']}")
        if status["stale"]:
            st.warning(status["stale"] + ". It is read again before the next step; drafts that cite a figure that moved will stop standing.")
            if st.button("Read the workbook again now", key="agent_refresh"):
                result = work.refresh()
                st.session_state["agent_notice"] = work._effects_line(result.get("candidate_effects")) if result.get("ok") else str(result.get("error"))
                st.rerun()
        if st.session_state.get("agent_notice"):
            st.info("The workbook was read again: " + st.session_state.pop("agent_notice"))
        text = st.text_area("Ask a question, or say what to put in the report", key="agent_request", height=90,
                            placeholder="e.g. Why did this account move? / put that in the report / shorten it and put it in the report")
        a, b, c = st.columns(3)
        if a.button("Ask", key="agent_ask", disabled=not text.strip()):
            with st.spinner("Investigating…"):
                session_state[TURN_KEY] = work.ask(text.strip())
        if b.button("Put into report", key="agent_report"):
            with st.spinner("Making the draft…"):
                session_state[TURN_KEY] = work.report(text.strip())
        if c.button("Save session to a local file", key="agent_save"):
            st.caption("Saved to " + work.save())
        _show_turn(session_state.get(TURN_KEY))
        st.markdown("##### Drafts")
        _show_candidates(work, bool(work.source.base_results))
        accepted = [v for v in work.candidates_view() if v["status"] == "accepted"]
        if st.button(f"Export a deck from {len(accepted)} accepted draft(s)", key="agent_export", disabled=not accepted):
            with st.spinner("Writing the deck and checking the slides…"):
                session_state[EXPORT_KEY] = work.export()
        _show_export(session_state.get(EXPORT_KEY))
        open_items = work.unresolved()
        if open_items:
            st.markdown(f"##### Open items ({len(open_items)})")
            for item in open_items:
                st.write(f"({item['from']}) {item['account'] or 'workbook'}: {item['message']}")
