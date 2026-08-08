# Things that were tried and did not work

An index, not a store. Every entry points at a comment that is still in the
code — this file exists so you can find the warning **before** you make the
change, rather than after, when you happen to be reading that function.

The full reasoning stays at the call site. Nothing was moved here.

**Why this matters:** the grid-table header colour below was re-added twice in
one day, each time producing genuinely blank pages in real PowerPoint, because
the record of the first attempt was not where the second attempt was made.

Line numbers drift. The **function name is the stable anchor** — grep for that.

---

## Layout / packing — `fdd_utils/pptx/gen_packing.py`

**Do not remove the 1.05 floor in the relax ladder.**
`_optimize_slot_fill` · tried 2026-08-05 · made things materially worse.
Without it the next tier is 1.35, and a real 6-account export front-loaded
slide 1 to 109% while slide 2 sat at 63%/0%. Side effect to know about:
configuring `shape_height_utilization` below 1.05 therefore has no effect.
Closing the residual ~1.05x properly means making the DP add a slot (the S_min
expansion), not touching this ladder.

**Do not narrow the overflow move back to a single trailing account.**
`_rebalance_overflowing_boundaries`. When several trailing accounts collectively
overflow, `rest_used <= cur_cap` is never true, the branch falls through, and
`_try_partial_split_overflow_forward` bails on the same arithmetic — so nothing
moves at all. Confirmed on a real export: a slot at 109% next to a **completely
empty** next slot.

**Do not raise the split-gap floor back to 0.5.**
`_maximize_forward_fill`. It was lowered deliberately once the accurate-measurer
backoff/trim and the number/currency-safe split points landed. At 0.5 the
1.0-unit floors compound into a worst-case ~2-line leftover — the "some pages
2 lines short" report that survived the other fixes.

**The forward-fill attempt budget shrinks per failed candidate** (8 attempts).
`_maximize_forward_fill`. Read the comment before making the search broader.

**Table accounts continue in the packer's last slot, not a fresh one.**
`_append_table_accounts_to_distribution` / `flow`. Starting a new slot leaves
the previous one barely filled.

## Table rendering — `fdd_utils/pptx/gen_tables.py`

**Do not restyle the grid table's header band without incremental testing in
real PowerPoint.**
`_fill_table_placeholder` · reverted 2026-08-04 · **caused blank pages**.
Two consecutive real Chinese exports rendered completely blank in real
PowerPoint on exactly the two slides this function's table lands on (the BS/IS
overview pages), with abnormally slow open — while python-pptx read every shape
and cell back intact from the same files. The Commentary-band deletion was the
first suspect and was ruled out; the per-cell colour/border changes are what
remained. See also `:1364` for the blue-band styling this replaced, and
`_render_presentation_table` for the navy period-header row that is safe.

**Black column separators across a navy header band read as a grid pasted over
the title** — the complaint that the separators fuse visually with the blue band. Use white hairlines between
date columns; keep the black rule only under the band.

**A bare tint (#DCE6F1) is too light.** A user photo of a fresh export still
read as "hasn't taken effect". The current shade is deliberately more saturated.

## Commentary splitting — `fdd_utils/pptx/gen_splitting.py`

**Validate a split against the real measurer, not chars-per-line.**
`_split_commentary_at_boundary`. The crude estimate disagrees with what
PowerPoint renders.

**Chinese has no spaces, so Latin word-boundary logic does not transfer.**
`_split_commentary_at_boundary` — jieba is used for this, and it has its own
failure mode: it cannot see a company name that was never a dictionary word.

**Never split between a number and its unit.** `_snap_split_before_number` —
万元/亿元 are magnitude units attached to the number before them.

## Capacity model — `fdd_utils/pptx/payloads.py`

**Do not bump `shape_height_utilization` to compensate for low page fill.**
It was raised 1.08 → 1.15 → 1.25 (BS override to 1.47) chasing this, and the
real cause was elsewhere: the measurement assumed a 6-9pt paragraph gap and 0.9
Chinese line spacing that the renderer never applies (it hardcodes a flat 3pt
gap and 1.0 spacing), plus a capacity formula that floored away up to a full
line per box. Fixing those raised real capacity 40-50%.

## Extraction — `fdd_utils/workbook/`

**The mapping-score floor of 45.0 applies to every candidate, not just
`sheet_kind == "other"`.** `resolver.py` (module level). 18.0 and 42.0 were both
confirmed false positives on real client databooks; the same pattern appeared on
a sheet that *did* have a detectable stage row.

**A real breakdown can be two-level** — a real 营业成本 table has 折旧与摊销
under a parent. `schedules.py` (module level).

**Do not return a bare `None` from the detail-table guard** — it was
indistinguishable from "no table found". `extract_presentation_detail_table`.

**The IS used to run to the end of the sheet, unlike the BS.**
`extract_balance_sheet_and_income_statement`.

## Executive summary — `fdd_utils/pptx/generation.py`

**Every way the summary band ends up blank looks identical from outside** —
shape missing, no source text, generator returned nothing, blank pre-generated
summary. Each silently did nothing, and a blank summary band cost several
rounds of guessing. The cause is now named at WARNING level so
`inspect_databook.py`'s export-log analysis surfaces it.

**Four generator attributes are delegating methods on purpose.**
`PowerPointGenerator` (near `load_template`). `find_shape_by_name`,
`_prepare_structured_data_for_slides`, `_presentation_table_for_account` and
`_expand_commentary_to_cover_summary` live in `helpers.py` but are called on the
generator by `inspect_databook.py` and `ad-hoc/pptx-probes/`. Removing the
delegators reproduces an `AttributeError` that only surfaces at section 10 of a
`--run-ai` run — i.e. after ~12 minutes and real tokens have been spent.

---

## Adding to this file

Add an entry when you revert something, or when you find out why an earlier
approach failed. Keep the reasoning in the code comment; keep this to the
one-line "do not do X, because Y" plus the function name.
