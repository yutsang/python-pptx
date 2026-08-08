# CLAUDE.md — FDD Tool

## Project Overview
Financial Due Diligence (FDD) tool that processes Excel databooks into AI-generated commentary and PowerPoint slides. The pipeline: Profile sheets -> Resolve mappings -> Normalize schedules -> Reconcile -> AI subagent pipeline -> PPTX export.

## Handing work back to the user (ALWAYS do this)

End any response that needs something from the user with an explicit
**「你要做的」** section. Never leave them to infer it. It must state:

1. **The exact command**, copy-pasteable, nothing to fill in.
2. **How long it costs.** Label it — a full `--run-ai` pass is ~12 minutes
   and real tokens; `inspect_pptx.py` / `inspect_table_bands.py` against an
   already-exported file is seconds and free. Never make them spend the
   expensive one when the cheap one answers the question.
3. **What to look for**, as named lines or numbers in the output, and
   **what to paste back**.
4. **What I could NOT verify here, and why.** This machine differs from
   theirs in ways that have repeatedly hidden real bugs:
   - `fdd_utils/template.pptx` is gitignored and per-machine (this copy has
     no placeholder text and different insets);
   - font metrics are `system-font` here vs `client-metrics` there, so the
     same sentence wraps to a different number of lines;
   - `fdd_utils/config.yml` is gitignored, so a value measured here is NOT
     theirs — check it on their side before quoting it.
   A local PASS on anything geometry-, template- or font-dependent proves
   nothing. Say so rather than implying it is verified.
5. **If nothing is needed from them, say that explicitly too.**

When a change can only be judged by eye (colours, punctuation, spacing),
say so and ask them to open the deck — don't present a geometry check as
if it settled the question.

## Testing Requirements (Mandatory Before Completing Any Task)
1. **Databook test**: verify functionality changes against the real
   databook: `python inspect_databook.py for_test/<databook>.xlsx --run-ai --export-pptx`
   (the user runs this; it is ~12 min and costs real tokens). The file is
   gitignored — the whole `for_test/` folder is gitignored.
2. **PPTX overflow check**: after any change affecting commentary text or
   PPTX generation, run `inspect_pptx.py` on the exported deck and confirm
   no `OVERFLOW RISK`, no `TABLE OVERLAPS REAL TEXT`, and no
   `TEMPLATE PLACEHOLDER LEAKED`. Cheap and free against an existing file.
3. **Import check**: `python -c "from fdd_utils.workbook import process_workbook_data; from fdd_utils.ui import render_sidebar_upload; from fdd_utils.ai import FDDConfig; print('OK')"`

## Before changing layout, packing or table styling

Read **`docs/failed-attempts.md`** first. It indexes the things that were tried
and reverted, with the function name to grep for. It is an index — the full
reasoning is still in the code comment at each site.

**Never delete a comment that records a failed attempt.** The grid-table header
colour was re-added twice in one day, each time producing genuinely blank pages
in real PowerPoint, because the record of the first attempt was not where the
second attempt was made. Verbose is not the same as worthless: cut restatement,
keep history.

## Key Architecture

The four big modules are **packages**, not files. Each was split along the
section markers it already carried; every `__init__.py` re-exports exactly what
the flat module exposed, so `from fdd_utils.pptx import X` is unchanged.

- `fdd_utils/workbook/` — Workbook processing engine, 12 modules.
  `inspector.py` is the shared utility hub (called from everywhere, calls almost
  nothing); `schedules.py` and `statements.py` are the heavy consumers;
  `flow.py` holds `process_workbook_data`, the top-level entry.
- `fdd_utils/ai/` — AI subagent pipeline, 7 modules. `config`, `english` and
  `logging` are leaves; `pipeline.py` is the only orchestrator.
  `SUBAGENT_SEQUENCE` is the truth: Generator -> Auditor -> Validator.
  subagent_3 (Refiner) is DORMANT — it has prompts but is not in the sequence.
  The Validator is also SELECTIVE: it runs only on accounts whose text asserts
  a causal claim (`commentary_asserts_inference`); every other account is
  grounded by `verify_commentary` with no LLM call. So most accounts cost 2 LLM
  calls, not 3. A retry loop (`_run_feedback_loop_for_key`) re-runs all three
  when a *hallucination* is found — never for a `reasoning` flag, which the
  deliverable wants.
- `fdd_utils/ui/` — Streamlit UI, 6 modules. Note `pptx_export.py` and the
  `batch_*` functions in `ai_panel.py` contain no `st.` calls — they are
  headless orchestration filed under the UI package, which is why
  `inspect_databook.py` imports from here.
- `fdd_utils/pptx/` — PPTX generation. `PowerPointGenerator` is assembled from
  five mixins, one per concern: `gen_packing` (which account lands in which
  slot: DP + rebalance passes), `gen_tables` (drawing tables), `gen_splitting`
  (cutting commentary at an acceptable boundary), `gen_measurement` (how tall
  text comes out), `gen_summary` (the executive summary band). `generation.py`
  keeps `__init__`, the template, slide orchestration and `save`.
  `helpers.py` holds functions lifted out of the class because they needed no
  instance state — four of them also have delegating methods on the generator
  because tooling outside the package calls them as attributes.
- `fdd_utils/mappings.yml` — Account definitions, aliases, subagent_1 prompts
- `fdd_utils/prompts.yml` — Subagent 2-4 prompts, plus `style_pack` (the
  language/formatting/judgement text `PromptStylePack` reads)
- `fdd_utils/config.yml` — Runtime config (AI providers, debug mode, processing
  settings). Gitignored and per-machine — never quote a value measured on one
  machine as if it were the user's.

One-off diagnostic scripts live in `ad-hoc/{contracts,bridge,fonts,
pptx-probes,databook-probes,workbench}/`. The repo root keeps only the live
tools: `inspect_databook.py`, `inspect_pptx.py`, `inspect_table_bands.py`,
`diagnose_summary.py` and the capacity toolkit.

## Naming Conventions
- AI agents are called "subagents" (subagent_1 through subagent_4)
- Pipeline stages: Generator -> Auditor -> Refiner -> Validator
- Mapping keys match sheet names or canonical account names from mappings.yml
