# Financial Due Diligence (FDD) Tool

Automated financial commentary generation from Excel databooks, powered by a 4-agent AI pipeline with reconciliation and PowerPoint export.

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
|  Generator ──> Auditor ──> Refiner ──> Validator
|  (create)     (verify)    (tighten)   (evidence check)
|      ^                                     |
|      |_____ feedback loop (if needed) _____|
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
| `fdd_utils/ai.py` | AI config, prompt engine, 4-subagent pipeline, feedback loop |
| `fdd_utils/pptx.py` | PPTX payload building, slide generation, executive summaries |
| `fdd_utils/ui.py` | Streamlit UI, processed view, AI panel, sidebar |
| `fdd_utils/mappings.yml` | Account definitions, aliases, Generator prompts |
| `fdd_utils/prompts.yml` | Auditor / Refiner / Validator prompts |
| `fdd_utils/config.yml` | Runtime config (AI providers, agent parameters, PPTX tuning) |

---

## The 4 Subagents

| Stage | Agent | Role |
|-------|-------|------|
| 1 | **Generator** | Creates financial commentary from data + prompts |
| 2 | **Auditor** | Verifies figures, trend direction, and format accuracy |
| 3 | **Refiner** | Tightens length while preserving key facts and reasoning |
| 4 | **Validator** | Final evidence check with clause-level hallucination detection |

A feedback loop retries Generator + Validator when too many clauses are unsupported.

---

## Reliability & Resilience

The pipeline assumes the LLM endpoint may be slow, rate-limited, or briefly
unavailable. Several layers cushion that:

```
        per-call          per-stage        per-export
        +-------+         +-------+        +---------+
LLM ->  | retry | -fail-> |breaker|-trip-> |fallback |  -> bullet rendered
        | 3x w/ |         | 4 in  |        |(data    |
        | 2s/5s |         | a row |        | only)   |
        +-------+         +-------+        +---------+
```

| Layer | Behaviour | Tunable |
|---|---|---|
| **Per-call timeout** | 30s on the HTTP client and the thread join. Failures throw fast. | `_run_ai_call(timeout=...)` in `ai.py` |
| **Retry with exponential backoff** | 3 attempts per call: `0s → 2s → 5s` between them. Lets the API recover instead of hammering it. | `retry_backoffs` in `process_single_agent_item` |
| **Circuit breaker (per stage)** | After 4 consecutive failures inside a stage, remaining calls in that stage skip the LLM and go straight to the fallback. Resets at the start of every new stage. | `_StageCircuitBreaker(threshold=4)` in `ai.py` |
| **Deterministic fallback** | When all retries are exhausted, the Generator falls back to a one-line data-only bullet built from the dataframe (`the balance as at <date> totalled CNY<X> million.`) marked as auto-summary. Subsequent stages reuse the previous successful output if they fail. | `_build_deterministic_fallback_bullet` in `ai.py` |
| **Concurrency cap** | At most **2 concurrent LLM calls** (across all three threadpools: `run_agent_stage`, feedback-loop, evaluator). Prevents the pipeline from saturating a stressed endpoint. | `max_workers = 2` in `ai.py` and `pptx.py` config |
| **PPTX-export bypass** | If section summaries weren't pre-generated during the AI phase, PPTX export does NOT make a fallback LLM call — it leaves the executive-summary box blank rather than burning 1–3 min per slide. Re-run the AI step to populate. | `pre_generated_summary` branch in `pptx.py` |

End-to-end the worst case is now bounded: each LLM call costs at most `(0+30) + (2+30) + (5+30) = 97s` before the breaker trips, after which the rest of the stage falls through to the data-only fallback in milliseconds.

---

## Slot Packing & Page Layout

PPTX commentary is distributed across template slides via a DP packer:

```
[ AI bullets ]                      [ PPTX template slots ]
   OI (140w)                          P3 R   (single, ~200w cap)
   OC (130w)         packing          P4 L   (~110w cap)
   T&S (60w)        ----------->      P4 R   (~110w cap)
   GA  (90w)                          P5 L   (~110w cap)
   Fin (70w)                          P5 R   (~110w cap)
   ...                                P6 L/R (~110w cap)
```

The packer:

1. **Estimates capacity** per shape from its rendered height (Pillow font metrics) divided by `line_height + PARA_SPACE_AFTER`.
2. **Multiplies capacity by `shape_height_utilization`** (currently **1.25**) — PPT autofit absorbs the small overflow at 9pt and lets bullets pack tighter.
3. **Runs a DP** that minimises `(num_nonempty_slots, max_fill_ratio)` lexicographically — uses the fewest slots possible, then balances fill among the slots used.
4. **Applies progressive relax** if the strict DP can't find a partition: capacity multiplier ramps `1.0 → 1.05 → 1.35 → 1.6 → 10×` until feasible. This guarantees a partition exists rather than dropping accounts on overflow.
5. **Pulls whole bullets forward** into half-empty earlier slots when `move_whole_min_fill_ratio` (0.50) is met.
6. **Hard length caps in prompts** keep individual bullets from blowing out a slot in the first place: BS bullets 30–80 words, IS bullets 80–150 words.

Tuning knobs live in `pptx.py` under `commentary_packing`:

| Setting | Default | Effect |
|---|---|---|
| `shape_height_utilization` | 1.25 | How aggressively the packer over-fills nominal slot height |
| `target_fill_min_ratio` | 0.95 | Lower bound the DP tries to hit per slot |
| `target_fill_max_ratio` | 1.00 | Upper bound the DP tries to hit per slot |
| `move_whole_min_fill_ratio` | 0.50 | Pull a whole bullet forward when current slot is at least this full |
| `minimum_slot_lines` | 22 | Floor for capacity heuristic when shape height can't be measured |
| `max_commentary_slides_per_statement` | 4 | How many slides to claim per statement (BS / IS) |

---

## PPTX Page Fill

Each account in `mappings.yml` has a word-count target in its AI prompt. Important accounts get longer commentary; minor ones get shorter. After generation, the tool automatically groups accounts across text boxes to balance how full each page is.

If a page looks under-filled, the fix is usually to increase the word budget in that account's `mappings.yml` prompt and re-run. If the template text boxes were resized, check whether the existing budgets still produce enough content — the layout adjusts automatically, but the AI needs enough words to fill the new space.

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

## Highlight Mechanism

The Validator returns clause-level annotations (`data-backed`, `reasoning`, `hallucination`) that are surfaced visually:

| Layer | Where | Colour |
|---|---|---|
| UI panel | Streamlit `build_highlighted_commentary_html` | inline orange (reasoning) / red (hallucination) |
| PPTX bullets | `_add_runs_for_line` in `pptx.py` | per-run RGB: orange `(213,94,0)` / red `(200,16,46)` |

Numbers are matched against source data with **rounding tolerance**: ±5% for million-scale, exact for sub-million comma-int, percentage rounding to 1dp accepted, period labels normalised (`FY24` ↔ `31 December 2024`, `1M26` ↔ `January 2026`).

---

## Run

`streamlit run fdd_app.py`
