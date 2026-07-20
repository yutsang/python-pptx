# Financial Due Diligence (FDD) Tool

Automated financial commentary generation from Excel databooks, powered by a multi-stage AI pipeline with reconciliation and PowerPoint export.

---

## Quick Start

```bash
pip install -r requirements.txt
streamlit run fdd_app.py
```

---

## Pipeline Overview

```
Excel Databook
      |
      v
+---------------------+
| 1. Profile & Resolve|  Detect sheet types, match tabs to account mappings
+---------------------+
      |
      v
+---------------------+
| 2. Normalize        |  Extract indicative-adjusted periods, build DataFrames
+---------------------+
      |
      v
+---------------------+
| 3. Reconcile        |  Cross-verify tab totals against Financials sheet (BS/IS)
+---------------------+
      |
      v
+---------------------------+
| 4. AI Subagent Pipeline   |
|                           |
|  Generator ──> Auditor ──> Validator   (Refiner: wired up, dormant)
|  (create)     (verify)    (evidence check)
|      ^                          |
|      |____ feedback loop (if needed) ____|
+---------------------------+
      |
      v
+---------------------+
| 5. PPTX Export      |  Slot-based text distribution across template slides
+---------------------+
      |
      v
  Final Report (.pptx)
```

---

## Architecture

| Module | Responsibility |
|--------|---------------|
| `fdd_utils/workbook.py` | Workbook profiling, sheet resolution, normalization, reconciliation |
| `fdd_utils/ai.py` | AI config, prompt engine, subagent pipeline + harness + audit + feedback loop |
| `fdd_utils/pptx.py` | PPTX payload building, slide generation, executive summaries |
| `fdd_utils/ui.py` | Streamlit UI, processed view, AI panel, sidebar |
| `fdd_utils/mappings.yml` | Account definitions, aliases, Generator prompts |
| `fdd_utils/prompts.yml` | Auditor / Refiner / Validator prompts |
| `fdd_utils/config.yml` | Runtime config (AI providers, agent parameters, PPTX tuning) |

---

## The Subagents

Named `subagent_1`–`subagent_4` in code/config for historical reasons, but only
**3 stages run at runtime** — `subagent_3` (Refiner) is wired up, prompted, and
tested, but deliberately dormant (`SUBAGENT_SEQUENCE` in `ai.py` skips it). It
stays in the codebase because tightening-for-length is a real, recurring need
that's cheap to re-enable (one line) the moment a account type needs it again —
removing it outright would mean re-deriving the prompt from scratch later.

| Stage | Agent | Role | Runs by default? |
|-------|-------|------|---|
| 1 | **Generator** | Creates financial commentary from data + prompts | Yes |
| 2 | **Auditor** | Verifies figures, trend direction, and format accuracy | Yes |
| 3 | **Refiner** | Tightens length while preserving key facts and reasoning | No (dormant) |
| 4 | **Validator** | Final evidence check with clause-level hallucination detection | Yes |

A feedback loop retries Generator + Auditor + Validator when too many clauses
come back unsupported — see [AI Engineering](#ai-engineering-prompt-harness-audit-loop) below.

---

## AI Engineering: Prompt, Harness, Audit, Loop

Four methodology-level concerns, stacked on top of each other, that turn
"call an LLM once" into a reliable pipeline:

- **Prompt** — every account type gets tailored instructions (what to cover,
  tone, length), so the model lands close to final quality in one pass
  instead of relying on trial and error.
- **Harness** — LLM calls can fail, hang, or come back garbled. The system
  retries automatically, and if a stage keeps failing, falls back to a safe
  data-only summary so the report still finishes.
- **Audit** — after the model writes commentary, every number it cites is
  checked against the actual source data by code, not by asking the model to
  grade itself. A mismatch is treated as fact — it overrides whatever the
  model claims.
- **Loop** — when too much of a bullet fails the audit, the system
  automatically regenerates it with the specific problems fed back, up to a
  couple of retries, before accepting the result.

```mermaid
flowchart LR
    A[Generate] --> B["Harness:<br/>retry / fallback if the call fails"]
    B --> C["Audit:<br/>check every number against source data"]
    C -->|too many issues| D[Feed issues back]
    D --> A
    C -->|clean enough| E["Accept +<br/>highlight any remaining flags"]
```

---

## Text-to-Layout Utilisation: How the Packer Decides What Goes Where

AI-written commentary varies in length; the PPTX template's text boxes don't.
The packer's job is to decide which account's text goes in which box so every
box reads full without spilling over, using as few slides as possible.

Methodology, in four steps:

1. **Measure real space** — how much room each account's text actually needs,
   using real font rendering, not a word-count guess.
2. **Optimise the layout** — an algorithm decides which accounts go on which
   slide, favouring filling earlier slides fully and letting only the very
   last one run lighter (so the report doesn't end on an obviously
   under-filled page, but also doesn't spread everything evenly-thin).
3. **Clean up edge cases** — a handful of follow-up passes catch patterns the
   main algorithm doesn't fully solve on its own — e.g. one box ending up
   empty next to a full one, or a leftover sliver of text that should be
   folded into its neighbour instead of standing alone. This is the most
   iterated-on part of the system: each pass exists because a real generated
   report hit that exact pattern.
4. **A safety net, not the plan** — if content still doesn't quite fit,
   PowerPoint's own autofit shrinks the font slightly rather than cutting
   text off. Normal pages never need this; it only engages on a genuine edge
   case.

```mermaid
flowchart TD
    A[Text + boxes] --> B[Measure real space needed]
    B --> C["Optimise:<br/>fewest slides, fill earlier ones first"]
    C --> D["Clean-up passes<br/>for edge cases"]
    D --> E{Still doesn't fit?}
    E -- yes --> F[Shrink font slightly]
    E -- no --> G[Render as-is]
```

The main lever for controlling how full a page looks is upstream of all this:
each account's AI prompt carries its own word-count target, sized to that
account's real complexity. The packer can rearrange text but can't conjure
content the prompt didn't ask for — if a page looks under-filled, that's
usually the first place to look.

---

## Style Enforcement

Two layers ensure commentary matches the project's reference style:

1. **Prompt-level** (`mappings.yml`, `prompts.yml`) — formulaic openings, banned bloat patterns, length caps per account type, anti-hallucination rules.
2. **Post-process polish** (`polish_english_commentary` in `ai.py`) — deterministic regex pass that catches AI leaks the prompt missed:
   - ISO dates → `dd Month yyyy`
   - `The balance as at` → `the balance as at`
   - `CNY 7.90 million` → `CNY7.9 million` (no space, 1dp)
   - `CNY 78.2K` → `CNY78,200` (no K notation)
   - `CNY0` / `CNY0.0 million` → `nil`
   - Strips period-on-period filler, verbose cross-checks, advisory `You should…`, annualisation projections, calculated rates not in source, land residual hallucinations, and meta-commentary.

---

## Run

`streamlit run fdd_app.py`
