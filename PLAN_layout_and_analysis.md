# Plan: table-slot empty space + FDD analytical capability

Written 2026-08-02 for a follow-up implementation session. Investigated against the
real deck photos (project team's own deliverable, same engagement) and the current
codebase at commit 3011558. Each workstream is independently shippable; within a
workstream the items are ordered.

User's asks, verbatim anchors:
1. "table + 字 + table 用盡空間 … 現在是隻要有subtable就沒有了那樣 大量empty區域"
   — fill the empty areas below presentation tables. Fallback if too hard:
   "subtable能夠轉換成sublist去處理 那樣文字計算上可能能夠繼承之前的modules".
2. IS items with large variation (e.g. >30%) must be explained from remarks; if no
   remark, model reasoning is allowed BUT must be highlighted (highlight mechanism
   already exists and is trusted). Expenses disproportionate to revenue growth
   (e.g. 管理费用) and big BS movements need independent analysis.
3. "有subagent懷疑是會把句子更加standardize了 這個或許會抹去我們想要的來自llm能力的分析"
   — find and fix the standardization that erases analytical content.
4. Generally increase LLM analytical capability "as a FDD consultant".

---

## Established facts (verified in code this session — do not re-derive)

- Active pipeline is **3 stages**: `SUBAGENT_SEQUENCE` (ai.py:3749) =
  Generator → Auditor → Validator. **3_Refiner is NOT active** (merged into
  Auditor; its prompts remain in prompts.yml and config.yml but never run).
- `_variance_analysis_guidance` (ai.py:2360) ALREADY implements most of ask #2:
  deterministic >30% movement computation on the LAST TWO full periods (partial
  tail excluded), remarks-present → reason-from-remarks-marked-as-judgement,
  remarks-absent (IS) → "state size/direction + 尚待与管理层确认, do not invent",
  plus a revenue-disproportion line via `peer_context`
  (`_build_peer_context`, ai.py:3995, wired at ai.py:4073).
- The disproportion line only fires INSIDE the `abs(pct) < 30: return ""` gate
  (ai.py:2423) — see B1, this is the biggest real gap vs the ask.
- The hallucination/highlight mechanism: Validator emits `clause_reviews`
  (data-backed / reasoning / hallucination), deterministic grounding combines in
  `_combine_verdict`/`verify_commentary` (ai.py:1249/1287), UI+PPTX highlight
  ungrounded clauses. User explicitly trusts this and WANTS unremarked reasoning
  to flow through it (highlighted), not be deleted upstream.
- The standardizer is the **Auditor** (prompts.yml `2_Auditor`). Specific
  conflicting instructions identified in B3 below. Validator is fine (it marks
  but preserves: Eng user_prompt items 8/11/12/13 explicitly preserve analytical
  content).
- Layout: the 4 IS table accounts are pulled out of the packing pool entirely
  (pptx.py `apply_structured_data_to_slides` ~5656-5685), each claims a whole
  L/R slot via `_append_table_accounts_to_distribution` (pptx.py:2091,
  slot-sharing between SMALL table accounts exists since 031c0db but real
  content rarely triggers it — every real pair exceeded cap by 31-197pt).
  Ordinary accounts NEVER enter a table slot's leftover space.
- Real-deck layout target (photos IMG_0228/0229/0230): one column flows
  lead-in → table → source → ➢bullets → NEXT account (table or plain text)
  continuously; page 4/4 has 财务费用 table followed by 投资收益 + 营业外支出
  plain-text bullets IN THE SAME column; a long account's ➢bullets can even
  continue into the next column (that continuation is explicitly OUT OF SCOPE
  here — noted as a known limitation, do not attempt).
- Packing history warning: `_distribute_content_across_slots`'s internals
  (greedy + DP + 6 rebalance passes) are the most regression-prone code in the
  repo — clause-DP redesign tried and deleted; two fixes reverted after
  real-world regressions. **Do not modify its internals for this work.** Both
  layout options below only shrink its input pool / consume its output, exactly
  like the existing table-account path already does.

---

## Workstream A — layout: use the empty space

### A1. Trailing-flow backfill (recommended first)

Ordinary accounts that come AFTER the first table account in statement order
(for the real IS: 投资收益, 营业外支出 — both 1-2 sentences) currently get packed
into early text slots, while the table slots below the tables sit empty. Instead,
withhold them from the normal packer pool and flow them, in statement order, into
the table slots' remaining space — reproducing the real deck's 4/4 page exactly.

- In `apply_structured_data_to_slides` (pptx.py ~5656): when splitting
  `table_items` / `normal_items`, any normal item positioned after the first
  table item in `structured_data` order goes into a new `trailing_items` list
  (only when `tables_enabled` and `table_items` non-empty).
- Extend `_append_table_accounts_to_distribution`: after placing table accounts
  (unchanged logic, `slot_fill_pt` dict), walk `trailing_items` in order; for
  each, estimate its text-block height (reuse
  `_estimate_table_account_block_height_pt`'s lead-in arm — category + key +
  commentary via `_calculate_content_lines(shape=None)` × std_lh ×
  `_TEXT_HEIGHT_SAFETY_FACTOR`); append to the slot where the LAST table
  account landed if it fits within `_TABLE_SLOT_CAPACITY_PT ×
  _TABLE_SLOT_PACK_THRESHOLD`, else the next table slot with room (in-order,
  not first-fit — preserves reading order), else FALL BACK to returning it to
  the normal packer pool (safe degradation; never drop).
  NOTE the fallback ordering constraint: the normal packer runs BEFORE this
  function. Simplest correct structure: decide the trailing placement BEFORE
  calling `_distribute_content_across_slots` (a small pre-pass that returns
  `(placed_trailing, returned_to_pool)`), so the packer's pool is final by the
  time it runs. Do NOT try to re-open the packer's finished distribution.
- Extend `_render_table_accounts_stack` (pptx.py:2159): it already walks
  `account_data_list` at a running EMU offset; for an account WITHOUT
  `_presentation_table`, render only the lead-in textbox arm (fresh textbox +
  `_fill_text_main_bullets_with_category_and_key`, which is proven safe with
  arbitrary fresh text_frames) and skip the table/source/explanation arm. The
  slot dispatch condition at ~5791 currently requires ALL accounts in the slot
  to be table-bearing (`all(a.get("_presentation_table"))`) — change to ANY,
  and keep the pure-text-slot path (no table accounts at all) going through the
  normal shared-text-frame path unchanged.
- Category header de-dup: `_render_table_accounts_stack` already tracks
  `current_category` — trailing items reuse it (投资收益 under 财务费用 both
  "Expenses"-adjacent; verify with real category values).
- Tests (scratchpad): extend `render_shared_slot_test.py` pattern — 4 table
  accounts + 2 short trailing text accounts; assert trailing text lands in the
  same column as the last table, below it, within slide bounds, no overlap;
  assert a LONG trailing account falls back to the normal pool (renders in a
  textMainBullets slot); Kunshan feature-off regression byte-identical.
- Real-run verification: `inspect_pptx_tables.py` shape dump should show the
  trailing bullets in the table column; `inspect_pptx.py` (post-3011558) checks
  the new boxes' own overflow automatically.

Known limitations to state honestly in the commit: text BEFORE the first table
account can't backfill (营业收入 overflow stays in its own slots); a table
account's ➢bullets still never continue across columns.

### A2. Sublist fallback mode (user's own suggestion; small, independent)

Config-gated alternative rendering: `pptx_settings.presentation_tables.style:
"table" (default) | "sublist"`. In sublist mode the account is NEVER pulled out
of the packing pool — the table dict is converted to text and the whole account
inherits every existing text module (packing, splitting, cross-column
continuation, overflow handling) for free.

- Implementation point: `apply_structured_data_to_slides` ~5660 — when style is
  sublist and `_presentation_table_for_account(item)` returns a table, rewrite
  `item["commentary"]` = lead-in + one line per top-level component
  (indent children with 、-joined names or skip children entirely — decide by
  eyeballing one render) + post_table_text, then treat as a NORMAL item (do not
  set `_presentation_table`). Format values with the existing
  `_format_table_value` ÷ `source_multiplier` (the 1000x lesson, cadbce8).
  Keep per-line period figures to LATEST period + total-row all-periods, not
  4 figures × every component — a 13-component × 4-period text dump would be
  worse than the empty space it replaces.
- Grounding: component values exist in `financial_data`, so the deterministic
  grounding check passes; no prompt changes needed.
- Tests: synthetic render in sublist mode → no table shapes, no pulled slots,
  text present, packer distributes normally; feature-off unchanged.

### A3. Full unified flow (DEFERRED — do not build now)

Sequential walk of ALL accounts with tables as atomic in-flow blocks would be
the true real-deck model, but it requires teaching the packer atomic blocks and
mixed-unit accounting — i.e. modifying the fragile packer internals. Only
revisit if A1+A2 measurably fail the user's expectation on a real run.

---

## Workstream B — analytical capability

### B1. Decouple the disproportion trigger from the 30% own-movement gate

`_variance_analysis_guidance` returns "" when the account's own movement <30%
(ai.py:2423), so the peer/revenue-disproportion line can never fire for exactly
the case the user named: 管理费用 flat (−4%) while revenue collapsed −36% — a
40pp gap that the real deck WOULD comment on (fixed-cost professional fees
don't scale with revenue). Fix: compute `peer_line` for IS non-revenue accounts
INDEPENDENTLY (own-|pct| may be small); if own movement <30% but |gap| ≥30pp,
emit a disproportion-only block (no "material movement" header — a
"本科目相对收入变动不成比例" header, with both percentages stated as computed
facts). Revenue-account self-comparison guard: skip when the account IS the
`revenue_key`.

### B2. Cover all adjacent full-period pairs, not just the latest

`series[-2:]` only (ai.py:2419). The user's ask is "IS variation比較大的items"
generally — 2023→2024 movements are invisible today. Loop adjacent full-period
pairs; emit one compact line per ≥30% pair (cap at the 2-3 most material to
bound prompt size); handle sign flips ("由净收益转为净支出" style — 财务费用
2023→2024 flips sign, a raw % there is nonsense; state the flip instead of a
percentage). Keep the existing latest-pair block as the primary, earlier pairs
as one-liners.

### B3. Stop the Auditor erasing analysis (the standardization fix)

Confirmed conflicts in prompts.yml `2_Auditor` — the variance guidance is
injected into the SAME prompt (shared rules, point 12/13) but the hard caps and
banned-pattern rules sit above it and win. Precise edits:

1. Eng system point 3 ("T&S, G&A, Fin Exp…: current-period figures only… Max 2
   sentences") and point 7 length caps; Chi system point 12 caps (财务费用/
   税金及附加/管理费用: 2-3句): add the exception "unless a 【重大变动提示】/
   [MATERIAL MOVEMENT] or disproportion instruction is present below — its
   required explanation (and the two figures it compares) sits OUTSIDE the cap,
   +1-2 sentences max".
   ALSO put the same sentence INSIDE the generated guidance text itself
   (`_variance_analysis_guidance` head), since the guidance is only present
   when it fires — self-carrying exception, no stale-cap risk.
2. Eng system point 8 "Invented drivers … DELETE unless data/remarks explicitly
   state them" and Chi point 7 "不得保留无数据支持的推断": narrow to
   UNMARKED external-cause claims stated as fact. Judgement-marked inference
   (预计/推测/主要系…所致/尚待与管理层确认; "mainly attributable to…",
   "expected to…") that reasons from remarks/data per the variance rule must be
   PRESERVED — it is flagged and highlighted downstream by design; deleting it
   here destroys the user-wanted analysis before the highlighter ever sees it.
3. Eng USER-prompt checklist item 7a strips "PoP filler" WITHOUT the system
   prompt's own "unless the movement is material AND explained" qualifier —
   add the qualifier there too (checklist wins in practice).
4. Do NOT touch Validator (already preserve+mark). 3_Refiner is inactive; align
   its "无支持，则删除或弱化" text only as a cheap future-proofing line item,
   lowest priority.

Verification for B3 is necessarily a REAL A/B: same databook, before/after
prompts, diff the 4 expense accounts' + 2-3 big-movement BS accounts' output for
(a) movement explanation present, (b) hedged reasoning survives to final_content,
(c) it arrives highlighted (clause_reviews category "reasoning"), (d) no bloat
regression on the no-movement accounts (they must stay short — spot-check
货币资金/预付款项 lengths unchanged).

### B4. FDD analytical-lens guidance block (the "as a consultant" ask)

New static block (e.g. `_analytical_lens_guidance`, same injection pattern as
`{variance_analysis_guidance}`; Generator + Auditor only, not Validator), ~10
lines, bilingual, giving the model the ANALYTICAL VOCABULARY the deck actually
uses rather than generic instructions:
- cost nature: rigid/fixed (depreciation, property tax, professional-service
  fees, insurance) vs revenue-linked (utilities net-off, property-management
  outsourcing) — a fixed cost NOT tracking revenue is normal and sayable; a
  variable cost diverging from revenue is a flag;
- one-off vs recurring (营业外支出 penalties; 汇兑损益 direction flips);
- related-party terms (pricing formula, interest-free, settlement-before-close)
  — always name the counterparty;
- stub-period comparison ONLY via annualization or same-length prior stub;
- BS: working-capital linkage (AR/预收 vs revenue trend), capex cycle
  (在建工程→固定资产 transfer explains both sides), reclassification signals
  (一年内到期 vs 长期借款).
All subject to the existing grounding rules — this block adds interpretive
frames, not license to invent facts. Keep it SHORT; every line here costs prompt
budget on every account.

### B5. Config

`analysis:` block in config.yml: `variance_threshold_pct: 30`,
`disproportion_gap_pp: 30`, `max_extra_pairs: 2` — read in
`_variance_analysis_guidance` (thread via existing kwargs path), defaults
matching today's hardcoded values so no-config behaves identically.

---

## Verification gates (CLAUDE.md-mandated, every commit)

1. Import check: `python -c "from fdd_utils.workbook import process_workbook_data; from fdd_utils.ui import render_sidebar_upload; from fdd_utils.ai import FDDConfig; print('OK')"`.
2. Kunshan databook regression (`kunshan_pptx_regression_test.py` in scratchpad
   pattern — rebuild if scratchpad rotated; feature off by default must stay
   byte-identical).
3. PPTX overflow: synthetic renders + `inspect_pptx.py` (now covers table-stack
   boxes) with zero new warnings; then a REAL run on the user's side (they run
   Windows, paste both inspect outputs back — established workflow).
4. B-workstream prompt changes additionally need the real A/B described in B3 —
   synthetic tests cannot verify prompt-behaviour changes; budget one full real
   pipeline run per prompt iteration and expect 1-2 iterations.

## Suggested order

A1 → A2 (both layout, no prompt risk, independently verifiable synthetically)
→ B1+B2 (deterministic code, unit-testable with synthetic DataFrames)
→ B3 (prompt edits + real A/B) → B4 (new block, same A/B run) → B5 (config).
Commit each separately. Memory files to update on completion:
`project_report_detail_tables.md` (A), `project_is_variance_and_remarks.md` (B).
