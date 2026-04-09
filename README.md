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

`streamlit run fdd_app.py`
