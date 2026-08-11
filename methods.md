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
   │              └─► verify_commentary()   ← deterministic, 0 LLM calls
   │                    every account now carries clause_reviews from here
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
| Checks numbers | — | yes, by reading | yes, by reading |
| Deterministic number-grounding | no | **yes** (`verify_commentary`) | yes |
| Emits `clause_reviews` | no | **yes** | yes |
| Can trigger the retry loop | no | **yes** | yes |
| Judges a causal claim | — | — | **yes — its unique job** |

### The Auditor is an independent reviewer, by construction

Every agent call is a **fresh, stateless** LLM call — `messages` is
`[system, user]` with no conversation history carried across stages. The
Auditor sees the Generator's text cold, not as its own prior turn.

It receives both sides of the comparison. Its user prompt carries
`先前输出: {output}` (the commentary) and `原始数据: {financial_data}` (the
data), and `render_prompt` builds that payload for **every** agent, not just
the Generator — so it includes the main table, the multi-period analysis table
with its breakdown component rows, supporting notes, and the side-column
remarks.

Its checklist is explicitly numbers-and-format: *「与原始数据相比, 所有数字是否
准确?」*, *「趋势交叉验证：若评论称某科目"增加"或"减少"，核实方向是否与源数据
各期数值一致」*, plus amount formatting, date formatting, materiality
thresholds and whether remarks are data-supported.

### What it could not do before, and now can

`verify_commentary` used to be gated on `agent_name == "subagent_4"`. That was
an accident of implementation rather than a design decision — it is pure
arithmetic over the account's own data, costs no tokens, and does not care
which agent last touched the text. It now runs after the Auditor too, filed
under the same single validation record. Consequences:

1. **Every account carries `clause_reviews` after stage 2**, with no third LLM
   call — including the accounts `"selective"` will skip.
2. **An Auditor-stage defect can trigger a retry.** The gate reads that record,
   so a fabricated figure surviving stage 2 is caught without waiting for a
   Validator that may never run for that account.
3. The Validator **overwrites** the record when it does run — its text is
   later, and it merges its own judgement in.

This leaves the Validator with only the job code cannot do: judging a causal
claim.

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

## 8. House style enforced as code

Two rules exist in the prompts *and* as deterministic post-processing in
`_finalize_agent_content`, because a rule that only lives in prose is not a
guarantee. Neither touches an amount, so number-grounding is unaffected.

**Zero balances** (`humanise_zero_balance`). A real report does not write
"余额为0元". Balance-sheet accounts read 无余额, income-statement accounts
未发生. The regex carries a negative lookahead so a real amount that merely
begins with a zero digit is never rewritten.

**Repeated names in an enumeration** (`dedupe_enumeration_prefix`). A cash
account list names the same bank and branch on every item, which reads as
padding. The name is stated once; later items keep only what distinguishes
them:

```
before:  …主要包括<bank><branch>#1#-A户 …、<bank><branch>#2#-B户 …
after:   …主要包括<bank><branch>#1#-A户 …、#2#-B户 …
```

It only fires when consecutive items share a long enough prefix, so ordinary
short overlaps are left alone, and it never strips an item to nothing.

---

## 9. Executive summary

The band is written by an LLM (`generate_section_summary`), but **only if a
summary was generated before export**. `export_pptx_from_structured_data_combined`
deliberately makes no LLM call of its own — an in-export call was reported to
hang for many minutes on a flaky API — so when nothing is passed to its
`pre_generated_summaries` argument it falls back to `_generate_page_summary`,
which splices each account's opening sentence together.

That fallback is easy to mistake for a real summary: it is grammatical, it
fills the band, and nothing in the export log says it happened. The tell is
that it reads as a verbatim copy of the page's first bullet.

Both the batch UI path and the CLI now call the shared `build_section_summaries`
before exporting. The CLI previously did not, so its decks never carried an AI
summary at all.

---

## 10. Open questions

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
