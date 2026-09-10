# Financial Due Diligence (FDD) Tool

Automated financial commentary generation from Excel databooks, powered by a
multi-stage AI pipeline with reconciliation and PowerPoint export.

---

## Quick Start

```bash
pip install -r requirements.txt
streamlit run fdd_app.py
```

---

## Pipeline

**藍**是程式算出來的，**橙**是模型寫的，**灰虛線**是出錯時的退路。橙色只有四
塊，而且其中三塊都只是在寫字；模型唯一一次做裁決，是判斷一句因果講不講得通。
金額、對數、重試與否、留哪一次、版面怎麼排，全部藍色。

兩個位置值得留意。哪些科目配明細表，是在寫任何 prompt **之前**就全份 deck 一
次裁決完的，所以文字不可能承諾一張畫不出來的表。而決定一段評論可不可信的，
是對該科目自身數字做的算術；模型只被問一件事——它說的那個原因站不站得住。

```mermaid
flowchart TD
    classDef det fill:#E8F0FE,stroke:#1A56DB,stroke-width:2px,color:#0B2A6B
    classDef llm fill:#FFF1E0,stroke:#E8710A,stroke-width:2px,color:#7A3E00
    classDef edge fill:#FFFFFF,stroke:#5F6368,stroke-width:2px,color:#202124
    classDef weak fill:#F1F3F4,stroke:#9AA0A6,stroke-width:1px,color:#5F6368
    classDef gate fill:#FEF7E0,stroke:#B06000,stroke-width:2px,color:#5C3A00

    XL["Excel 底稿"]
    OUT["最終簡報 .pptx"]

    subgraph S1["① 先把事實算清楚 — 全程無模型"]
        direction TB
        RES["識別與對應<br/>分頁性質 · 科目"]
        NORM["標準化<br/>調整後各期"]
        REC["對數<br/>分頁合計 vs 財務報表"]
        SUB["表格入選裁決<br/>全份 deck 一次定"]
        FACT["事實表<br/>期間變動 · 重大性<br/>構成殘差 · 集中度"]
        RES --> NORM --> REC --> SUB --> FACT
    end

    subgraph S2["② 模型只負責語言"]
        direction TB
        GEN["生成初稿"]
        AUD["覆核<br/>冷讀，不帶上一輪對話"]
        GEN --> AUD
    end

    subgraph S3["③ 判決 — 算術先行"]
        direction TB
        CHK["逐句核對金額<br/>算術，不是模型"]
        VAL["宣稱因果時才判斷<br/>唯一交給模型的裁決"]
        CHK --> VAL
    end

    GATE{"有找不到<br/>來源的數字？"}
    RE["指名錯處<br/>只重寫該科目"]
    ARB["仲裁<br/>留缺陷最少那次<br/>不是最後一次"]
    HS["房屋風格<br/>零餘額 · 單位小數 · 公司名"]
    SUM["執行摘要"]
    PACK["量度真實高度並分欄<br/>字型度量，不是字數估算"]
    HAR["逾時重送 · 階段熔斷<br/>最後退回純資料摘要"]
    QA["匯出後檢查<br/>溢出 · 壓字 · 佔位符"]

    XL --> RES
    FACT --> GEN
    AUD --> CHK
    VAL --> GATE
    GATE -->|"有"| RE
    RE -->|"最多 3 次"| GEN
    GATE -->|"沒有"| ARB
    RE -.->|"次數用盡"| ARB
    ARB --> HS --> SUM --> PACK --> OUT

    GEN -.->|"呼叫失敗"| HAR
    HAR -.-> CHK
    OUT -.->|"僅 CLI"| QA

    class XL,RES,NORM,REC,SUB,FACT,CHK,RE,ARB,HS,PACK det
    class GEN,AUD,VAL,SUM llm
    class GATE gate
    class HAR,QA weak
    class OUT edge
```

---

## Architecture

The four big modules are **packages**, not single files. Each `__init__.py`
re-exports exactly what the flat module exposed, so `from fdd_utils.pptx import X`
is unchanged.

| Module | Responsibility |
|--------|---------------|
| `fdd_utils/workbook/` | Workbook profiling, sheet resolution, normalization, reconciliation, movement analysis |
| `fdd_utils/ai/` | AI config, prompt engine, subagent pipeline, harness, deterministic verification, feedback loop |
| `fdd_utils/pptx/` | PPTX payload building, packing, slide generation, executive summaries |
| `fdd_utils/ui/` | Streamlit UI, processed view, AI panel, sidebar, headless export helpers |
| `fdd_utils/mappings.yml` | Account definitions, aliases, Generator prompts |
| `fdd_utils/prompts.yml` | Auditor / Refiner / Validator prompts, plus the shared style pack |
| `fdd_utils/config.yml` | Runtime config (AI providers, agent parameters, PPTX tuning). Per-machine, not tracked |

Diagnostic entry points live at the repo root: `inspect_databook.py` (end-to-end,
optionally with `--run-ai --export-pptx`), `inspect_pptx.py` (geometry and
overflow checks on an exported deck), and `inspect_render_truth.py` (drives real
PowerPoint over COM on Windows — the only tool that reaches actual ground truth).

---

## The Subagents

Named `subagent_1`–`subagent_4` in code and config for historical reasons, but
only **three stages run** — `subagent_3` (Refiner) is wired up, prompted and
tested, yet deliberately dormant (`SUBAGENT_SEQUENCE` in
`fdd_utils/ai/pipeline.py` skips it). It stays because tightening-for-length is a
real recurring need that costs one line to re-enable; removing it would mean
re-deriving the prompt later.

| Stage | Agent | Role | Runs by default? |
|-------|-------|------|---|
| 1 | **Generator** | Writes commentary from the data and the account's prompt | Yes |
| 2 | **Auditor** | Re-reads the draft cold against the source data | Yes |
| 3 | **Refiner** | Tightens length while preserving facts and reasoning | No (dormant) |
| 4 | **Validator** | Judges causal claims; emits clause-level verdicts | Only when the text asserts a cause |

A typical account therefore costs two LLM calls, not four.

---

## Methodology

Five ideas do most of the work. They are ordered by how much they change the
output, not by where they sit in the pipeline.

### 1. Facts are compiled before prose

The model is kept out of arithmetic wherever the answer can be settled in code.
Movements, percentage changes, materiality flags, concentration shares and
verified parent/child hierarchies are computed first and handed over as settled
facts, with an explicit instruction to quote them rather than derive anything
further. A self-derived number has no source to match against, which is exactly
where invented figures come from. One arithmetic task is still asked of the
model: reconciling an itemised composition against its total, where the residual
is precomputed only once an account carries more than three top-level
components.

Two guards worth knowing: a percentage across a sign change is meaningless, so
those movements are described qualitatively; and a stub period is annualized to
a common basis before any cross-year comparison, and excluded outright from the
revenue benchmark an expense account is measured against.

### 2. The judge is arithmetic, not a model

After the model writes, every amount it cites is checked against the account's
own source data by code. The verdict outranks the model's in both directions: it
overrides a fabricated figure the model defended, and it dismisses a
"hallucination" the model flagged on a figure that does match. A model-judged
loop can be talked round by a confident rewrite; a sum cannot.

### 3. Attempts are bounded, and the best one wins

Tell a model its answer is wrong and it will agree and change it, whether or not
it was wrong. Each extra round drifts further toward whatever the feedback
implies and further from the data. So retries are capped, and the trigger is
deliberately narrow: a provable fabrication fires one, while an unprovable
inference on its own does not — that inference is the analysis the deliverable
exists for, and the deck renders it in orange on purpose. A second, looser
trigger survives for output that is broadly unsupported rather than specifically
wrong. An arbiter then keeps the attempt with the fewest defects rather than the
last one produced.

### 4. Nothing crosses a run boundary

There is no cache file, no learned preference, no memory of a previous
engagement. Every derived fact lives for the duration of one run and is rebuilt
next time. This is a decision rather than an omission: cross-run learning is
where one reviewer's wording preference silently becomes the next client's wrong
deck. Promoting a correction stays a human edit to the prompt files.

### 5. Every rule that matters lives in code, not only in a prompt

A rule that exists only as prompt text is a preference, not a guarantee. Where a
convention is load-bearing — zero balances reading as "no balance" rather than
"0", magnitude units and decimal places, company names shortened after first
mention — it is enforced deterministically after the model is finished. Two
attempts to hold the company-name rule through prompting alone did not survive
real runs.

---

## Known limits

Stated because a tool that hides its blind spots is harder to trust than one
that names them.

- **The number check discriminates, but not perfectly, and the history is worth
  knowing.** An amount used to be tested against a pool that folded in every
  other tab of the same statement, which made sibling tabs about 91% of it and
  the account's own cells under 1%. Feeding that pool a real figure multiplied
  by a random factor got it accepted around 88% of the time, and a tenfold unit
  error passed about 80% of the time — so `data-backed` asserted little beyond
  "a number of about this size exists somewhere in this statement". The pool is
  now the account's own data by default, and the same tests come back at roughly
  31% and 27%. That is a working check rather than a rubber stamp, but a third
  of deliberately wrong magnitudes still pass, so the verdict is evidence and
  not proof.
- **A `data-backed` verdict is also not one meaning.** A clause whose only
  defect is an unverifiable causal claim is demoted back to `data-backed` by a
  confidence floor rather than shown as flagged, so the label covers both "the
  numbers matched" and "nothing here was checkable".
- **What the check does still catch**, and why it is not worthless: a figure
  derived a way the pool has no route to — a difference, a sum of non-adjacent
  rows, a cross-statement reference — is genuinely absent and is flagged. Bare
  numbers and percentages are deliberately not extracted as amounts, so a wrong
  ratio is never caught here at all.
- **Unsupported clauses are rare, and most are judgement rather than
  arithmetic.** Across the archived runs, roughly one clause in sixty comes back
  unsupported. About a quarter of those are the arithmetic kind; the remaining
  three quarters are the model's own opinion that something is unsupported,
  which no amount of recomputation settles. Read that rate together with the
  first bullet: the arithmetic share is low partly because the arithmetic test
  is easy to pass. It is also measured *after* the confidence demotion above.
- **Extraction depth is the real ceiling on analysis.** When a workpaper's
  breakdown does not survive extraction, no amount of verification or prompting
  recovers it. This is a per-firm structural problem rather than an industry one.
- **Layout can only be confirmed in real PowerPoint.** Every check in this repo
  except `inspect_render_truth.py` compares the model against itself, so a clean
  `inspect_pptx.py` run is not evidence that PowerPoint agrees.
- **False-positive and false-negative rates for the verification layer have not
  been measured** — only how often each verdict occurs.

See [`methods.md`](methods.md) for the verification loop in detail, including why
the retry gate is shaped the way it is and what the arbiter is guarding against.

---

## Run

```bash
streamlit run fdd_app.py                                    # the app
python inspect_databook.py <databook.xlsx>                  # free diagnostics
python inspect_databook.py <databook.xlsx> --run-ai --export-pptx   # full run
python inspect_pptx.py <deck.pptx>                          # geometry checks
```
