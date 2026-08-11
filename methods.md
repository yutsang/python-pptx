# Methods — how the commentary pipeline checks itself

**This file is committed. It must never contain client data** — no entity
names, no databook names, no monetary figures, no per-run counts. The repo's
`.gitignore` excludes every `.md` but `README.md` for exactly that reason
(working notes name clients); this file carries an explicit negation and has to
stay generic.

---

## The whole pipeline

One account's journey. Every account runs this independently, in parallel.

```
╔════════════════════════════════════════════════════════════════════╗
║  HARNESS ①   stateless calls · timeout resend · circuit breaker    ║
╚════════════════════════════════════════════════════════════════════╝

                    data for ONE account
                 (table · breakdown · remarks)
                              │
    ╭─────────────────────────▼──────────────────────────────╮
    │  ATTEMPT                          ② up to 3 attempts   │
    │                                                        │
    │   Generator ──► Auditor ──► [ Validator ]               │
    │    writes it    re-reads it   only when the text        │
    │                 against the   claims a CAUSE — judging  │
    │                 source data   that is all it adds ③     │
    │                      │             │                    │
    │                      ╰──────┬──────╯                    │
    │                             ▼                           │
    │                    verify_commentary            ④       │
    │                 arithmetic, not a model:                │
    │                 every amount vs the source              │
    │                             │                           │
    │                             ▼                           │
    │        data-backed  │  reasoning  │  hallucination      │
    ╰─────────────────────────────┬──────────────────────────╯
                                  │
                    any hallucination?  ⑤
                                  │
              ┌──── no ───────────┴─────────── yes ────┐
              │                                        │
              ▼                                  name the bad
           ship it                               amounts, feed
                                                 back, retry ⑥
                                                        │
                                            attempts exhausted
                                                        │
                                                        ▼
                                        ╭───────────────────────────╮
                                        │ ARBITER ⑦                 │
                                        │ keep the attempt with the │
                                        │ FEWEST hallucinations —   │
                                        │ NOT the last one          │
                                        ╰───────────────────────────╯
```

---

## What each marker is, and why

**① Harness.** Every LLM call is stateless — `[system, user]`, no conversation
carried between stages. The Auditor therefore reads the Generator's text
*cold*, like a reviewer, not as its own previous turn; a model reviewing its own
visible prior turn is much less likely to find fault with it. On top of that the
harness resends a call that times out (same prompt), and a circuit breaker stops
a stage that keeps failing instead of burning the whole run on retries.

**② Bounded attempts — the sycophancy limit.** This is the core reason the
budget is 3 and not "until clean". Tell a model its answer is wrong and it will
agree and change it, whether or not it was wrong. Each extra round pushes it
further toward whatever the feedback implies rather than toward the data, and
the text gets blander and less specific. So the loop is capped, and the three
mechanisms below all exist to stop agreeableness being mistaken for correction.

**③ Narrow scope.** The Validator only runs where the commentary asserts a
cause, because judging a causal claim is the one thing arithmetic cannot do.
Asking a model to re-review text that has nothing for it to judge invites it to
change something just to look useful.

**④ The judge is not a model.** Amounts are checked by arithmetic against the
account's own source data. This is what keeps the loop honest — a model-judged
loop can be talked round by a confident rewrite; a sum cannot. Its verdict on a
number outranks the model's, in both directions: it overrides a fabricated
figure the model defended, and it dismisses a "hallucination" the model flagged
on a figure that does match.

**⑤ A deliberately narrow gate.** Retry fires **only** on `hallucination`, never
on `reasoning`. A `reasoning` flag means the numbers are right but an inference
is not directly provable — which is the analysis an FDD deliverable is *for*
(the deck renders it in orange on purpose). Retrying on it spends tokens
punishing good writing and trains the output toward saying less. A ratio-based
gate was tried and is wrong at any threshold: one fabricated figure inside a
long correct paragraph dilutes below any sane cut-off, so it misses exactly the
case it exists for.

**⑥ Feedback names the defect.** The retry does not say "that was wrong, try
again" — it names the specific amounts that failed grounding, so the next
attempt has something concrete to correct rather than a mood to match.

**⑦ The arbiter is the real sycophancy guard.** Re-generation is **not**
monotonically better: attempt 3 can be the worst of the three, precisely because
each round drifts further under pressure to agree. So the pipeline never ships
"whatever came out last" — it scores every attempt by hallucination count and
keeps the best, breaking ties toward the later attempt (which has had the most
feedback applied). Without this, bounded retries would just guarantee shipping
the most-drifted version.

---

## Known limits

- The grounding pool holds each cell, column totals, sums of short runs of
  adjacent rows, and figures quoted in the notes. A legitimate figure derived
  another way — a difference, a ratio, a sum of non-adjacent rows, a
  cross-statement reference — is **not** in it and will be flagged.
- Matching carries a tolerance, and the pool is large, so a `data-backed`
  verdict is weak evidence of correctness rather than proof — weakest where the
  amounts are small.
- Neither error rate has been measured.
