# What changed, and why

A plain-language record of the analytical and wording work done on the
commentary pipeline, plus a standing assessment of what it would take to run
this outside real estate.

This file documents design decisions. It carries no entity names, no databook
names, and no engagement figures — keep it that way.

---

## 1. Analytical skills

### Shared across every account

**We stopped asking the model to do arithmetic.**

Before, the model got a table of numbers and was expected to notice what moved.
It frequently derived percentages itself, which is exactly where invented
figures come from — a self-derived number has no source to match against, so
the grounding check flags it.

Now the code computes every movement first and hands it over as a settled fact:
this account went from X to Y, a 43% drop, and that is material. The model only
has to *explain*, never to *calculate*.

A side effect that mattered more than expected: movements used to be ranked by
absolute size, so a small account that doubled was invisible next to a large one
that drifted a few percent. A line now qualifies on how far it moved *relative
to itself*, which surfaces small accounts that swung hard and correctly declines
to flag a large account whose percentage move is modest.

Two guards worth knowing about:

- A percentage across a sign change is meaningless (moving from negative to
  positive is not "a 336% increase"), so those are described qualitatively
  instead.
- A stub period against a full year reads as a ~90% collapse that is an artefact
  of period length. Partial tail periods are excluded from every comparison.

**We fixed a check that could never fire.**

The rule "flag an expense growing out of proportion to revenue" was gated behind
the account having *also* moved 30% on its own. But a fixed cost that stays flat
while revenue collapses is precisely the case worth flagging — and it never
moves 30%, so the gate made it unreachable. The two conditions are now
independent.

Thresholds for all of this live in `config.yml` under the `analysis:` block.
The code defaults are 30% for a material movement, 30 percentage points for a
disproportion gap, and at most 2 earlier period-pairs mentioned. An unconfigured
deployment behaves exactly as it did before the block existed.

**Component lists have to add up.**

Before, a composition read as bare categories with no amounts, and there was no
way to tell whether anything had been left out.

Now every item carries an amount, and the model must add the items up and
compare against the total *before* writing. Any difference is itself a component
it has to name and list last. Listing items that do not reach the total, with no
explanation of the rest, is treated as an incomplete disclosure.

Two related rules: never list a parent line and the lines that make it up
(adding both double-counts), and only ask for an itemised breakdown when
components are actually present — asking for one where there is none is an
invitation to invent it.

**Length is a ceiling, not a target.**

The old prompts read as word quotas, so the model padded to reach them. Padding
is the single biggest reason commentary reads as machine-generated.

The rule now carries a test: does this sentence tell the buyer something they
did not already know and that affects their view? If it merely restates a figure
visible in the table, leave it out. A small, single-component account with no
anomaly gets one or two sentences, and that is the correct answer.

The judgement rules also name what does *not* count as analysis — restating
period-by-period figures, narrating trends on immaterial balances, treating a
mechanical ratio such as "accounted for X% of the total movement" as insight,
and manufacturing a cause for every movement.

**English and Chinese deliberately differ on one point.**

When the remarks contain nothing that explains a movement:

- English states the movement and stops. It does not guess.
- Chinese reasons from other evidence in the same data and marks the result as
  judgement (主要系…所致, 预计系…). It is explicitly banned from writing
  「原因未在资料中说明」.

The reasoning: a clearly-labelled inference is worth more to the reader than a
blank space, and the review team can confirm or correct it. Either way, the
inference must start from a remark or a data point — never from assumed market,
competitive or macro causes.

**Where checking happens.**

The deterministic grounding pass now runs after the Auditor as well as the
Validator. It is pure arithmetic over the account's own data, costs no tokens,
and does not care which agent last touched the text — so an Auditor-stage defect
can trigger a retry instead of waiting for a Validator that may never run for
that account.

`validator_mode` is `selective`. It was briefly set to `always` and reverted the
same day: a deterministic hallucination verdict is consulted before any LLM
opinion and cannot be overridden by one, so coverage was never the lever. What
the Validator uniquely adds is judgement on a causal claim, which is what
`selective` already selects for.

### Balance sheet

- **Opening sentence discipline.** First sentence states the latest period-end
  balance and what it comprises. Nothing else. No listing every reporting period
  in the opening, no opening on a movement.
- **The nil-balance case.** When the latest period is zero, the model kept
  opening on the previous period's real number, which reads as if that were
  current. A feedback retry did not fix it. The computed dates are now written
  directly into the instruction rather than the rule being restated abstractly.
- **应收账款 ageing is mandatory.** It used to read "mention when supported",
  which in practice meant the topic was silently dropped whenever the data did
  not volunteer it. It is now addressed every time — state the ageing if
  supported, otherwise say plainly that none was provided. Inventing an ageing
  bucket or a bad-debt conclusion remains banned either way.
- **固定资产 had no instructions at all.** It was falling through to the generic
  default prompt, so any depreciation detail that appeared was riding on rich
  source data rather than on any standing instruction. It now has a dedicated
  prompt with depreciation method and useful life given *per category* — land,
  buildings and equipment differ, and one blanket figure across all three is
  wrong.
- **Advisory language.** Generic "you should confirm with management" is
  stripped. The one exception is a narrow, quantified recommendation tied to a
  specific finding already stated in the same bullet. That exception previously
  produced different figures across runs on identical data, because the model
  was constructing amounts from unlabelled raw balances; it now requires an
  explicitly labelled source figure.

### Income statement

- **Composition first.** Open with what the account consists of, not with an
  isolated trend sentence.
- **Do not repeat a table in prose.** Where the account already has a detail
  table on the page, the text gives one or two sentences of composition and
  hands off, rather than reciting every line again.
- **Know which costs are supposed to move.** Rigid costs — depreciation,
  property tax, professional-service fees — do not track revenue proportionally,
  and that is expected rather than something to explain away. Revenue-linked
  costs diverging from the revenue trend *is* worth flagging.
- **Drill-down caps per account type**, with one exemption: a material-movement
  or disproportion explanation does not count against the cap. That exemption is
  stated in both the prompt file and the code that generates the guidance, so
  neither can drift out of step with the other.
- **The income statement is bounded.** It had been swallowing the adjacent
  operating-KPI block.
- **The executive summary was never real in the CLI path.** The exporter makes
  no LLM call of its own by design, and passing it nothing silently fell back to
  splicing each account's opening sentence — which is why the summary band could
  read as a copy of the first bullet on the page, with nothing in the log saying
  so. The CLI now generates it through the same shared code path as the UI.

---

## 2. Wording

Best read as before and after.

### Currency prefix

| | |
|---|---|
| Before | 「人民币1,062.8万元…人民币398.2万元…人民币205.0万元」 |
| After | 「1,062.8万元…398.2万元…205.0万元」 |

Not a preference call — a style profile over the reference deliverable found
that essentially none of its amounts carried the prefix. 人民币 is now reserved
for telling currencies apart (as in 折合人民币), capped at one or two uses in a
deck, and banned twice within one sentence.

There was also a real defect here: the Auditor stripped the prefix and the
Validator added it straight back.

### Magnitude and precision

| Wrong | Right | Why |
|---|---|---|
| `77,930.0万元` | `7.8万元` | the raw figure had a unit stuck on it without being divided |
| `30,650.0万元` | `3.07亿元` | anything at or above 100 million must use 亿元 |
| `美元30.0 million` | `美元3,000.0万` | Chinese output never uses 百万 or "million", including for foreign-currency source figures |

Decimals are fixed: 万元 to one place, 亿元 to two, matching what the reference
reports do almost without exception. Below 万, the exact figure is written.

The English side is deliberately kept no more precise than its Chinese
counterpart: amounts from CNY10,000 up to just under a million are rounded to
the nearest thousand and written as `CNY238,000`, not `CNY238,366`, because the
Chinese 万-unit rule already rounds there. No space after `CNY`, no `K` suffix.

### Zero

| | |
|---|---|
| Before | 「余额为0元」 |
| After | 「无余额」 on a balance-sheet account, 「未发生」 on an income-statement one |

A subtlety that was being got wrong: if one period within a multi-period
sentence is zero, it must stay in the sentence and read as 未发生. Deleting it
makes the reader assume data is missing rather than that the value was zero.
Only a component that is zero in *every* period may be dropped from a
composition list. Those two cases were being conflated.

Both of these are now enforced in code, not only in prose — a rule that lives
only in a prompt is not a guarantee.

### Company names

| | |
|---|---|
| Before | 「某某物业管理有限公司」 repeated four times |
| After | 「某某物业」 |

A repeated full legal name costs most of a line every time it appears, and is
the most easily avoided waste of space in these bullets. The full form is
reserved for the first mention of a contract counterparty where the exact legal
entity matters.

This was attempted through the prompt twice and did not hold, so it is now done
deterministically in code after the model is finished. Commentary containing no
company name is returned unchanged.

### Where information came from

Banned:

> 「备注显示该款项将于交割前结清」

Required:

> 「该款项将于交割前结清」

But the reverse for anything management said:

> 「管理层表示，公区水电费由目标公司承担」

These look contradictory and are not. The reader does not care that we read a
remark — that is our filing system. The reader *does* care whether a fact was
verified by us or simply told to us, because that is how the strength of the
evidence is judged. Don't cite the data source; do cite the evidence source.

### Also banned

Dramatised description (急剧下降, 大幅跃升) — a DD report states facts and
judgements, it does not editorialise. Sentences that restate the account name
with no new information (「全部为XX明细余额」). Every bullet in a deck opening
with an identical phrase; the structural requirement is that the first sentence
covers the latest balance and composition, not that it uses one fixed form of
words.

### Related fixes

The canned sentences in `mappings.yml` were reframed as style examples rather
than facts to copy, and made conditional — they were at risk of constraining the
analysis rather than guiding the tone. A name repeated across items of one
enumeration (the same bank and branch on every cash line) is stated once.
Kinsoku is applied in the wrap simulation, and splits never land mid-company-name
or between a figure and its 万元/亿元 unit.

---

## 3. Running this outside real estate

**Feasible, but not a one-file swap.**

The machinery is industry-neutral. Extraction, packing, measurement, the
subagent pipeline and number grounding all only understand accounting structure.
Real-estate knowledge lives in the YAML, and only three places in the code
mention the sector at all — each of them a keyword list for "this row is a ratio,
not an amount" (出租率, 单位租金, 元/平方米).

The obstacle is how uneven the account knowledge is. `Inventory` — the heart of
a supply-chain business, where provisioning, goods in transit and turnover days
live — has roughly two orders of magnitude less written about it than 应收账款 or
投资性房地产, and no dedicated prompt at all. Filling that in is the main job.

Three other costs to budget for:

1. **Extraction depth is a workpaper problem, not an industry one.** The
   structural fixes made recently (dropped breakdown rows, marker characters,
   unexplained residuals) came from one firm's working habits. The same team's
   workpapers in a new sector inherit those fixes for free; a different team
   means a fresh set of structural surprises. This is the part that cannot be
   estimated in advance.
2. **The sheet-matching score floor was tuned on real-estate account names** and
   needs re-validating against a new vocabulary.
3. **The "how much is worth writing" calibration came from measuring a
   real-estate deliverable.** Another sector's sense of what matters is
   different — turnover, concentration, ageing — so it needs its own reference
   deck to calibrate against.

One item that is easy to miss: the analytical-lens text, the short paragraph
that primes the model with interpretive concepts, is entirely real-estate
vocabulary (tenant concentration, CIP-to-fixed-asset transfers, occupancy) and
would need rewriting from scratch. Against that, the composition/residual logic
and the deterministic movement calculation transfer unchanged.

**The cheap first step.** Run `inspect_databook.py` against one or two real
databooks from the target sector *without* `--run-ai`. Section 1 shows how many
sheets matched nothing; section 3 shows how deep the extraction got. Those two
sections alone distinguish "write more mappings" from "the extraction needs
rebuilding", and the run is free. Everything before that is speculation.
