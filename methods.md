# Methods — how commentary is generated, checked and grounded

Reference for the AI pipeline's verification design: which agent checks what,
where the check is an LLM judgement and where it is exact arithmetic, and what
the number-grounding pool can and cannot see.

**This file is committed. It must never contain client data** — no entity
names, no databook names, no monetary figures, no per-run account counts. The
repo's `.gitignore` excludes every `.md` but `README.md` for exactly that
reason (working notes name clients); this file carries an explicit negation
and therefore has to stay generic. Constants quoted below are code constants,
already public in this repo.

---

## 1. Pipeline shape

`SUBAGENT_SEQUENCE` is the truth: Generator → Auditor → Validator.
subagent_3 (Refiner) has prompts but is not in the sequence — it is dormant.

```
N accounts
   │
   ├─ Stage 1  Generator (subagent_1) ──────────────► N LLM calls
   │              writes the commentary
   │
   ├─ Stage 2  Auditor (subagent_2) ────────────────► N LLM calls
   │              verifies + tightens + finalises
   │
   └─ Stage 3  Validator (subagent_4)
          │
          │  processing.validator_mode
          │
          ├── "selective"  (see §6 for why this exists)
          │      │
          │      │  commentary_asserts_inference(text)?
          │      │  scans for causal markers, after stripping the
          │      │  "no cause available" disclaimers
          │      │
          │      ├─ asserts a cause ──► M accounts ──► M LLM calls
          │      │                          └─► verify_commentary(llm_reviews)
          │      │
          │      └─ no cause ──────────► N-M accounts ──► 0 LLM calls
          │                                  └─► _apply_deterministic_verification
          │                                       └─► verify_commentary(None)
          │
          └── "always"
                 every account ──────────► N LLM calls
                                             └─► verify_commentary(llm_reviews)
```

Both branches end in `verify_commentary`. **The number-grounding is identical
on both paths** — the only difference is whether an LLM opinion exists to be
merged in.

---

## 2. What each agent actually checks

| | Generator | Auditor | Validator |
|---|---|---|---|
| Writes commentary | yes | rewrites | rewrites |
| Checks numbers | — | **yes, by reading** | yes, by reading |
| Deterministic number-grounding | no | **no** | **yes** (`verify_commentary`) |
| Emits `clause_reviews` | no | **no** | yes |
| Can trigger the retry loop | no | **no** | yes |

### The Auditor does check numbers — but nothing records the result

Its prompt is explicit: *「验证先前代理输出中数字的准确性和重要性」*,
*「纠正任何不准确的数字或格式问题」*, *「趋势交叉验证：若评论称某科目"增加"
或"减少"，核实方向是否与源数据各期数值一致」*.

Three consequences follow, and they are the reason the Auditor's check does
not remove the need for the Validator's:

1. **No deterministic backstop.** `verify_commentary` is gated on
   `agent_name == "subagent_4"` (`pipeline.py`, in `process_single_agent_item`).
   The Auditor's number check is an LLM reading a table — the task LLMs are
   least reliable at, because it is exact lookup and arithmetic.
2. **No structured output.** The Auditor returns prose only. It emits no
   `clause_reviews`, so anything it silently fixes (or silently misses) is
   invisible to the deck's highlighting and to every downstream gate.
3. **It cannot drive the loop.** `_run_feedback_loop_for_key` reads
   `agent_4_validation.clause_reviews`. An Auditor objection is not
   representable in that structure, so it can never request a retry.

---

## 3. Inside `verify_commentary`

```
final_content
   │
   ├─ segment_clauses()  ──► one clause at a time
   │
   └─ per clause:
        │
        ├─ A. ground_amounts(clause, SourceIndex)
        │       extract_amounts(clause)
        │        ├─ no amount found ─────────► det = None      (defer)
        │        └─ amounts found ─► source.matches(x) for each
        │              ├─ all matched ───────► det = data-backed    conf 1.0
        │              └─ any unmatched ─────► det = hallucination  conf 0.9
        │
        ├─ B. _lookup_llm_review(clause)
        │       the LLM's own opinion, matched by clause overlap
        │       (always None on the deterministic-only path)
        │
        └─ C. _combine_verdict(det, llm)
```

`matches()` compares **magnitudes** — sign is dropped by `extract_amounts`, so
a negative source cell still matches a positively-written figure. The test is:

```
abs(target - value) <= max(500.0, 0.05 * value)
```

Two tiers, and the flat floor matters. Above CNY10k the 5% term dominates.
Below it the **flat 500 floor** does, and it is proportionally very wide — it
exists because Chinese commentary rounds sub-million amounts to one decimal of
万 (i.e. to the nearest thousand), a conventional rounding the old tight tier
flagged as hallucination. A genuine zero source cell is special-cased: it only
matches a target that also rounds to zero.

---

## 4. `_combine_verdict` precedence — the part that decides everything

```
if det == hallucination:
      ┌──────────────────────────────────────────────────────┐
      │  verdict = hallucination.  The LLM CANNOT override.   │  ★
      └──────────────────────────────────────────────────────┘
      code comment: "the model cannot override hard arithmetic"

elif det == data-backed:
      ├─ LLM says reasoning + unsupported ──► reasoning
      │                                       (numbers fine, inference doubted)
      └─ otherwise ────────────────────────► data-backed
                                              (an LLM "hallucination" claim
                                               here is dropped as a false
                                               positive)

else:                       # det is None — no checkable amount in the clause
      ├─ an LLM review exists ─────────────► use it
      └─ none ────────────────────────────► causal language ⇒ reasoning
                                              otherwise supported
```

### Consequence worth stating plainly

A figure the pool cannot find is flagged `hallucination` **before any LLM
opinion is consulted**. Therefore:

> Running the Validator on more accounts (`validator_mode: "always"`) does
> **not** rescue a correct-but-unfindable number. That path is closed by
> precedence, not by coverage.

The LLM's only decisive contribution is the middle branch: numbers all match,
but the *inference* is unsupported → `reasoning`. That is exactly the subset
`"selective"` already selects for.

---

## 5. The grounding pool (`SourceIndex`)

### What goes in

```
SourceIndex.from_df(df, sibling_dfs)
   │
   ├─ every numeric cell of the account's projection_df
   ├─ each column's total
   ├─ sums of every run of 2..4 CONSECUTIVE rows      (_adjacent_window_sums)
   ├─ all of the above for df.attrs["prompt_analysis_df"]  (the multi-period table)
   ├─ every number appearing in df.attrs text (notes / remarks / side columns)
   ├─ those text numbers × the annualisation factor, when 0 < months < 12
   └─ all of the above for each sibling df
                                    ▲
                                    └─ siblings are the SAME statement type only
```

### Where a legitimate figure can fall outside it

Derived from reading the code. **Not yet confirmed against a real run** —
confirm before treating any row as a known defect.

| Figure the commentary states | Findable? | Why |
|---|---|---|
| A single line item | yes | present as a cell |
| A column total | yes | totals are added explicitly |
| Sum of 2–4 **adjacent** rows | yes | window sums |
| Sum of **non-adjacent** rows | **no** | windows are consecutive runs only |
| Sum of **5 or more** rows | **no** | `max_window = 4` |
| A **difference** (A − B, "the remaining X") | **no** | pool holds no subtraction |
| A **ratio / percentage / per-unit** figure | **no** | pool holds no division |
| A figure from the **other statement** (BS clause citing IS) | **no** | siblings are same-type only |
| Annualised value of a table row | yes | the annualised column is in the df |
| Annualised value of a **remark** figure | yes | text numbers × factor |

### Both error directions are real

The pool is deliberately generous, and that cuts both ways:

- **False positive (flagged but correct).** Any row in the table above marked
  "no". This is authoritative per §4 and cannot be argued down by the LLM.
- **False negative (fabricated but passed).** Pool size grows roughly as
  `columns × (4 × rows)` per df, multiplied again across siblings, so it is
  large — and every entry carries a `max(500, 5%)` catchment around it. For
  small figures the flat 500 floor makes that catchment proportionally very
  wide. A fabricated figure therefore has a non-trivial chance of landing
  within tolerance of *something*. **A `data-backed` verdict is weak evidence
  of correctness**, not proof, and it is weakest exactly where the amounts
  are small.

Neither direction is currently measured. Quantifying the false-negative rate
would mean injecting known-fabricated figures and counting how many pass.

---

## 6. Why `"selective"` exists

Measured previously on a real run (figures deliberately omitted here — see the
private session notes): the Auditor does substantive work, while the
Validator's edits were mostly cosmetic, it stripped `➢` bullets, and it was
the single most expensive stage because it re-emits the full text *plus* a
per-clause review array that the deterministic layer then largely overwrites.

`commentary_asserts_inference` is a cheap deterministic pre-filter, and is
deliberately **over-inclusive**: a false positive costs one avoidable LLM
call, a false negative loses a real reasoning flag from the deck.

`主要为` / `主要包括` are deliberately **not** markers — they introduce a
composition, not a cause.

---

## 7. Feedback loop (`_run_feedback_loop_for_key`)

Already implemented and enabled by default.

```
attempt 0 = the original result
   │
   └─ up to max_retries times (default 2 → 3 attempts total):
         │
         ├─ gate: count_defective_clauses(clause_reviews)
         │        supported == False AND category == "hallucination"
         │        (RETRIABLE_CLAUSE_CATEGORIES)
         │        plus the unsupported-ratio as a secondary trigger
         │
         ├─ no defect ─► stop
         │
         └─ defect ────► format_validator_feedback_for_reprompt()
                         └─ re-run Generator → Auditor → Validator
                            with the feedback appended to the user comment
   │
   └─ final arbiter: score every attempt by hallucination count,
                     keep the BEST (ties break toward the latest).
                     Re-generation is not monotonically better — the last
                     attempt can be the worst.
```

### The gate is `hallucination` only, on purpose

An earlier design gated on `unsupported_clauses / total > threshold`. Every
unsupported clause observed on real runs was category `reasoning` — supportable
FDD inference, which the deliverable *wants* and the deck renders in orange.
Retrying on those spends tokens punishing good analysis and pushes the model
toward blander commentary.

A ratio is also the wrong **shape** of test at any threshold: one fabricated
figure inside a long correct bullet dilutes below any sane cut-off, so the
gate misses precisely the case it exists for.

---

## 8. Open questions

- Which pool gaps from §5 the reports actually hit. Answerable for free from
  an existing run's flagged clauses — no re-run needed.
- The false-negative rate (§5). Needs deliberate injection of known-bad
  figures.
- Whether the LLM Validator catches anything the deterministic layer does not,
  on accounts `"selective"` would have skipped. Answerable by diffing a
  `"always"` run against the deterministic verdicts for the same accounts.
- Both Chinese prompts (Auditor and Validator) carry `禁止使用要点、列表`, so
  `➢` bullets from the Generator get stripped inconsistently depending on
  which stage touches the account last. This affects rendered line counts,
  not just appearance.
