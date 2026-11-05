# Methods: how commentary is produced, checked and delivered

**This file is committed. It must never contain client data**: no entity
names, no databook names, no monetary figures, no per-run counts. The repo's
`.gitignore` excludes every `.md` but `README.md` for exactly that reason
(working notes name clients); this file carries an explicit negation and has to
stay generic.

Every statement below carries one of four labels.

| Label | Meaning |
|---|---|
| **Verified** | A command in this repo reproduces it on the current revision. The command is given. |
| **Verified offline** | Reproduced with synthetic data, a stub provider or a scripted policy. It shows the code does what is described. It says nothing about what a real model does. |
| **Not verified** | Implemented or asserted, with no reproducible check here. Needs a provider, Windows PowerPoint, or the company machine. |
| **Not built** | Described as a limit or a plan. No code. |

Two paths exist. The fixed pipeline (Generator, Auditor, selective Validator)
is the production default and is unchanged except where stated. The agent path
(`fdd_utils/agent`, `fdd_agent.py`) is off by default. The checks in sections 2
to 6 serve both.

---

## 1. The fixed pipeline (production default)

```
data for ONE account ──► Generator ──► Auditor ──► [ Validator ]  (only if the text claims a cause)
                                          │
                                  verify_commentary: arithmetic against the account's own source data
                                          │
                  data-backed │ reasoning │ hallucination
                                          │
                 any hallucination? ── no ──► done
                         │ yes
                 name the bad amounts, retry (up to 3 attempts)
                         │ exhausted
                 arbiter keeps the attempt with the fewest hallucinations
                         │
                 delivery decision (section 3): a kept attempt that still has one is BLOCKED
```

- **Stateless calls.** Each stage gets `[system, user]` with no conversation, so
  the Auditor reads the Generator's text cold. A timed-out call is resent with
  the same prompt and a circuit breaker stops a stage that keeps failing.
  *Verified offline* (`ad-hoc/workbench/replay_pipeline_state.py`).
- **The judge of amounts is arithmetic.** `verify_commentary` checks every amount
  against the account's own source data and its verdict outranks the model's
  in both directions. *Verified offline* (`ad-hoc/workbench/test_validator_defects.py`).
- **The retry gate is `hallucination` only.** A `reasoning` flag (numbers right,
  inference unproven) never fires a retry by itself; the deck renders it in
  orange. A ratio of unsupported clauses stays as a backstop at the configured
  share. *Verified offline* (`pipeline.count_defective_clauses`, covered by
  `test_delivery_decision.py`).
- **Bounded attempts and the arbiter.** The budget of 3 attempts and the rule
  "keep the best attempt, not the last" rest on the observation that
  re-generation is not monotonically better. That observation is a design
  rationale. The explanation usually given for it (a model agrees with
  "that was wrong" whether or not it was) is **Not verified**: this repo has no
  experiment that separates it from other causes.
- **Narrow Validator.** It runs only where the text asserts a cause. Whether
  that saves quality as well as tokens is **Not verified**.

## 2. Which accounts are sent to commentary

`derive_reconciliation_matched_keys` used to end by re-admitting any account
whose tab had a non-zero period, including a balance-sheet line that
reconciliation had marked `❌ Diff`. It now delegates to `workbook/selection.py`,
which gives each account a route and a reason: reconciled, reconciled by policy
(income-statement Diff), prior period ties (partial latest period), unchecked
with history (Financials line nil in the two latest periods), or excluded
(`bs_diff`, `no_reconciliation_row`, `no_activity`, `no_data`).

- Admitted keys are never a superset of the old function's. **Verified**:
  `python ad-hoc/workbench/test_reconciliation_selection.py` (includes the old
  function as a negative control and an exhaustive status-combination property
  test), `python ad-hoc/workbench/selection_parity.py <databook>`.
- On local test workbooks the new rule excludes some accounts the old one
  admitted. All of them are accounts no Financials line reconciles to, or a
  `❌ Diff` line. Whether the project team wants those back, flagged instead of
  excluded, is a product decision that has not been made.

## 3. Delivery decision

`fdd_utils/delivery.py` answers `ready`, `review_required` or `blocked` for a
deck and for each account. It reads the reconciliation selection, the run's
results, the statement source (Financials sheet or synthesized from the tabs)
and the run identity.

- An unresolved hallucination blocks the deck, whatever the retry history and
  whichever attempt the arbiter kept. A `reasoning` flag does not.
- Silence is not a clean result. A result set with no account commentary is
  blocked. An account the selection admitted but the results lack is a finding,
  and when none of them has commentary the deck is blocked. A missing
  statement source asks for review. An empty or ungrounded account text blocks.
- A deck built on synthesized statements, an account left out for lack of a
  reconciliation row, a partial-period line, or a Financials line with no tab
  asks for review. A wrong-entity run, a dead run or placeholder text blocks.
- A blocked or review deck is still exported, as a draft: file-name suffix
  (`_DRAFT-BLOCKED`, `_REVIEW`), a `<deck>.delivery.json` sidecar, and the deck's
  own document properties. No slide is changed. A caller that did not assess the
  deck gets `review_required`. A combined deck takes its worst part.
- **Verified offline**: `test_delivery_decision.py`, the stub run
  (`ad-hoc/workbench/stub_inspect_run.py <databook> --export`) and
  `test_offline_reexport.py`.
- **Not verified**: how the labelled deck behaves in the Streamlit UI beyond
  import and syntax (the UI banner and the batch download label were not
  exercised in a browser).

## 4. Evidence contract and claim verification

`fdd_utils/ai/evidence_contract.py` extends `AccountEvidence`. For each account
it keeps three bases apart: the tab's own figure (`raw_schedule`), the Financials
figure, and the figure the deck uses after the "deck follows Financials" rule
(`authority_adjusted`). It records a check per account, period and kind (tab total
against Financials; detail against total) with statuses `pass`, `mismatch`,
`ambiguous`, `not_comparable`, `missing` or `not_checkable`, and derives what the
commentary may do in each period. Replacing a total does not make the detail
under it verified. Hidden sheets are never offered as evidence.

`fdd_utils/ai/claims.py` has Python compute movements, shares, explicit sums,
tab-minus-Financials gaps and linear annualisation. A claim cites the facts it
rests on. The verifier then checks that the text says what the citations say:
period, line, basis, sign, unit, rounding to the printed precision, percentages
and zero bases, annualisation labelled as such, and a stated cause resting on a
cited note or marked as inference. Verdicts stay separate by dimension, and a
dimension nothing was checked for reads `n/a`, never `ok`.

Which checks a sentence gets follows from what its text states and what it
cites, never from what its author called it. `kind`, `inference` and
`cross_account` are declarations by whoever wrote the sentence. `kind` is kept for
display and decides no check: the same text and citations give the same verdict
under every legal kind (and under an unknown one, which reads as `amount`). The
earlier version let a wrong direction pass as `descriptive`, a replaced total pass
as `difference`, and a detail line explaining an untied total pass as `amount`;
it also made every `cause` sentence demand a note. `inference=true` waives two
things, a cause with no note behind it and a linear annualisation worded as a
forecast; it never excuses a figure, period, unit, sign, direction or percentage.
`cross_account=true` lets another account's fact be cited and leaves the claim
`unverified` (`CROSS_ACCOUNT_UNBOUND`), because the text cannot say which account
a figure belongs to. A flag written as the word "false" is false.

What "checked" means here has four parts.

- **Period.** Every amount and every percentage is bound to the period the
  sentence writes it under (the nearest period before it, else after it); one that
  belongs to the other cited period fails, which catches two figures swapped
  between periods. An amount with no period in the sentence, when the claim
  declares several, is `unverified`, and a candidate with an unverified claim is
  not accepted.
- **Percentages.** A percentage is also bound to the line the sentence names and
  to the kind of ratio the words say, a share of the total, a change between
  periods, or a difference of two shares in percentage points, so a share written
  against the wrong line, or called a growth rate, fails too. A percentage keeps
  its sign: a written minus on a ratio that is positive, or a plus on one that
  fell, fails. A minus is not read after a figure (a range, 30%-40%). An unsigned
  percentage that fell, in a sentence with neither a direction word nor a neutral
  word such as 变动, is `unverified` (`PERCENT_SIGN_UNSTATED`). The share of a
  deduction line such as a provision is negative; written without a minus it is
  read as its size, with a minus as signed, and 升至 is judged on whichever the
  sentence uses.
- **Direction.** A direction word (增加, 下降, 升至, 降至 and the like) is checked
  against the quantity it describes. The nearest cue before it, after the previous
  direction word, says which: 金额/余额 for the amount, 占比/比重 for the share;
  failing that, a money amount after it means the amount, a percentage-point
  figure means the share, and a percentage means whatever the percentage check
  bound it to. An amount can rise while its share falls, so each word is compared
  with the sign of its own quantity, for the line the clause names (the account's
  own movement when it names none). A direction with no matching movement or pair
  of shares behind it fails (`DIRECTION_UNSUPPORTED`); the other quantity's
  movement does not stand in. The result clause and the cause clause are read
  apart. A word in the result clause is a statement about the account and is
  always checked. A word on the cause side (after 由于/主要系, before 所致/导致,
  inside 受…影响) is a reason and belongs to the cause check, unless it names a
  line and a quantity the citations can test ("主要系甲公司余额增加所致"), in which
  case it must agree with them. A word with no quantity of its own is compared
  with every cited reading: it fails when they agree against it, and is
  `unverified` (`DIRECTION_UNBOUND`) when the amount and the share would answer
  differently, when the sentence says the opposite elsewhere, or when the word
  cannot be placed on either side of a cause (after 导致, or in a clause that
  follows one giving a reason, "主要系A，B"). A correct sentence of that last kind,
  such as "余额减少，导致坏账风险上升", therefore waits for a person.
- **Lines.** A line the sentence names must have a figure behind it when the
  sentence makes a line-level statement: it cites a line's figure or share, says
  "其中" or "占比", or puts a figure right after the name. A name on the cause side
  of a clause, with no figure after it, is a counterparty mentioned as a reason.
  A detail line may not be quoted in a period where the tab's detail does not
  add up to the total, whatever the sentence is for.

A cause is not verified by a note that shares words with it. Word overlap shows
the note is about the same thing. The checker screens for contradiction around the
shared words (negation, completed against expected, all against part, opposite
direction, a different named line) and fails on one; a cause that passes is
marked as related-but-unchecked and the deck renders it as reasoning until a
person has read the note against the sentence. Entailment is not tested. Cause
wording the pattern list does not recognise is not screened at all; citing a note
(`cause_cites`) makes the check run whatever the wording.

- **Verified offline**: `test_evidence_contract.py`, `test_claims.py`.
- **Verified** against ground truth that this repo generates:
  `python ad-hoc/workbench/claim_eval.py <databook>` writes correct claims from a
  real workbook's contracts and breaks each one in a named way, including a
  flipped direction word on the amount or on the share of a line whose amount and
  share move opposite ways, a minus sign on a share, and a flipped sign on a
  change percentage. Every sentence is also run under all eight legal kinds, with
  inference off and on. On the four local test workbooks the verifier accepted none
  of the 2,757 broken claims and rejected none of the 502 correct ones; no correct
  claim changed verdict with its kind, and none of the 4,896 broken claims
  (each mutation, inference off and on) was accepted under any kind; about one
  in ten correct claims is accepted with a caveat because the detail's nesting is
  unknown. The mutations were
  written by the same author as the verifier. This is not a measurement of
  model-written text. **Not verified**: false-accept and false-reject rates on
  claims a real model writes.
- The older pool search (`SourceIndex.matches`) remains in the fixed pipeline.
  **Re-measured on the current revision**, `replay_verification.py --decoys`
  (seed 1234, one local test workbook): about a third of deformed real values
  (tenfold, one-tenth, transposed digits, jitter) are still accepted. The
  historical figures quoted in earlier versions of this file (about 88% and 80%
  before narrowing, about 31% and 27% after) came from a different workbook and
  are not reproducible here.

## 5. Tools and the controller

`fdd_utils/agent/tools.py` wraps existing capabilities: `inspect_sources`,
`get_account_evidence`, `compute_analysis`, `trace_claim`, `validate_candidate`,
`numerical_summary`, `export_candidate`, `refresh_sources`, `diff_sources`,
`read_deck`, `update_deck`. A result is `ok` only when the thing asked for
happened. A failed validation, an exception, a stale source and bad arguments
are all `ok: false` with a code. Every result carries the source version.

`fdd_utils/agent/controller.py` is one loop. The model replies with a tool call,
an answer, `done`, or a question for the user. Code enforces: citations only to
refs a tool returned in this run, independent verification of every claim, open
issues derived from tool results that must be investigated or named as
unresolved, questions to the user only when sources are ambiguous or in
conflict, a `done` with export accepted only on a successful export of the
current draft against the current sources (a blocked export, an export of an
earlier draft, or one made against older sources does not count), and budgets (steps, estimated tokens, seconds, steps without
progress, answers sent back) that end the run with a report.

- **Verified offline**: `test_agent_tools.py`, `test_agent_controller.py`
  (scripted policies; one policy takes three different tool paths in three
  situations because of what the tools returned) and `test_agent_flows.py` (the
  five flows on a real workbook, real template and real packing).
- **Not verified**: how a real model chooses tools, whether it uses the open
  issues, how often its answers are sent back, and its token and call counts.
  The development machine has no configuration for the company gateway. It
  holds a key for one direct provider of a different model family, which is not
  used as a stand-in: a result from it says nothing about the company's models.

### 5.1 The working session

`fdd_utils/agent/workspace.py` keeps one thread across requests: ask, follow up,
put the analysis in the report, change its wording, accept, export, update a
returned deck. It holds the source binding, the refs the model has been shown,
the figures Python derived (in the order derived, so D3 stays D3), the issues
the tool results raised, every turn with its tool path and usage, the drafts,
and the user's decisions (accepted, withdrawn, exported) as events. "Put that
in the report" with no change of wording makes a draft from the last answer's
verified claims with no model call; a request to change the wording goes to the
model as an edit with the evidence already gathered.

The thread lives in memory. If the user asks for a session file it is written
as one JSON file with mode 0600, only to a path git does not track. It holds the
workbook's figures and commentary, never a prompt, a key or the config.

A draft rests on more than the figures it quotes. When the workbook is saved,
the session reads it again and decides per draft in three layers:

1. What it cites: each figure by its semantic id (value compared), each computed
   figure by its recipe and the figures it was computed from, each note by its
   text and the period and kind it belongs to. Anything gone or moved
   invalidates the draft and says which.
2. Every sentence is verified again against the new reading.
3. The rest of what it rests on, recorded when it was written: every other
   figure of the periods it is about, what each check came out as, what each
   period permits, the remarks that apply, the period list. If any of it moved
   the draft is kept, acceptance is withdrawn, and the changes are named.
   Accepting again acknowledges them; the event keeps the list.

A session file saved without that record cannot show the account did not move,
and says so when it is opened against other bytes.

**Verified offline**: `test_workspace.py` (each layer, with a negative control
that fails without layer 3) and, on a real workbook with its paired run,
`test_workspace_flows.py`. **Not verified**: how a person uses the panel, and
a returned deck after real PowerPoint has re-saved it.

### 5.2 Routing and the budget

`fdd_utils/agent/routing.py` is not a second agent. There is one controller and
one loop. Python still computes and verifies. A configured second model
(`agent.routing: escalate`) answers the remaining steps of a turn only after
`escalate_after` consecutive steps the first model failed with the evidence it
had. The default is `single`.

What a failure asks for next is fixed per code, in `routing.CLASS`:

| class | meaning | goes to the second model? |
|---|---|---|
| rewrite | the sentence is wrong or cannot be bound and the evidence in hand settles it: wrong period, sign or unit, an unbound direction word, an unstated sign, a statement about another account, a total the permissions do not let it state | only when repeated |
| evidence | what the sentence needs is not in the conversation yet | never |
| human | the sources conflict or are ambiguous and no rule chooses | never |
| refresh | the workbook was saved | never (read it again) |
| format | the reply was not a usable action | only when repeated |
| system | a tool crashed | never |

A test scans the verifier, the tools and the controller for every code they can
raise and fails on one that has no class.

The second model is handed a task state (goal, the latest draft with the
verifier's objections, one line per cited figure, the tools already used, the
open issues, what is left of the budget), not the conversation. It is held to
`handoff_max_tokens`. A turn that ends is classed as `ok`, `budget`,
`needs_human`, `evidence_insufficient`, `capability_insufficient` (with or
without the second model having been tried), `tool_error` or `provider_error`.

`routing.SharedBudget` is one set of books for every model, turn and attempt of
a session. It counts failed calls and replies that were thrown away. Three kinds
of number are kept apart and never added: what the provider reported, this
code's estimate (for every call, so the two can be compared), and answers the
response cache served. A cap applies to the provider's number where one was
reported and to the estimate where not. Retries made inside the provider SDK
are not visible to this code and are listed as unmeasured. A cost cap needs
`price_in` / `price_out` on the model's `available_models` entry; a model
without them has no cost and nothing is borrowed from another provider.

Thinking is set per role (`off`, `on`, or an effort level). Only provider
branches that send the parameter act on it (`local`: `allow_thinking`;
`workbench`: `reasoning_effort`). On the others the call record says "not sent
by this provider branch". The agent's ledger lines are labelled `agent_primary`
and `agent_escalation`, not Validator, and carry reasoning tokens where the
provider reports them.

**Verified offline**: `test_routing.py`, with scripted models. **Not verified**:
whether a second model does better, what the company gateway accepts as a
thinking parameter, and what it reports as usage. Both need a paid probe.

### 5.3 Investigation tools

`read_source_block`, `search_notes`, `cross_account_evidence` and
`adjustment_evidence` return location, source version, what was and was not
covered, and an error state. Text they retrieve is context; a cause may cite a
searched remark only as an inference. `cross_account_evidence` lists the
relationships that passed (support) and, apart, those that were looked for and
did not tie, with the test, what was compared, and why. A failed check is never
citable and a sentence resting on one stays unverified.

### 5.4 Baseline material

`ad-hoc/workbench/baseline_pack.py` finds the archived run that describes a
databook and packs both, with the accepted deck and the run's ledger lines, under
a git-ignored path. A run belongs to a databook when its identity names the
databook's sha256, or, for a run with no identity file, when every account's data
fingerprint matches the workbook as it reads now. It does not pair by name, and a
databook with other bytes (re-saved, rebuilt from a text export, a demo copy) has
no run. A pack is `complete`, `usable_with_gaps` (what is missing is listed with
why and how to get it) or not written. Files are searched for keys, bearer
tokens, the values of the config's key and address fields and addresses before
anything is written; config, `processing.log` and the stage prompts are not
packed. The ledger lines are reduced to a field whitelist and the ledger's error
text is cut to its type. **Verified offline**: `test_baseline_pack.py`, and a dry
run on the local workbooks. **Not verified**: any real client's pack.

Getting the material to a reviewing machine goes through text when files cannot be
copied. `text_transfer.py` carries the exact bytes (tar, zlib, base64, parts with
a CRC each, every file's sha256 checked on arrival), so a workbook arrives with the
hash its run names and everything inside it, including cell indentation. Only if
the workbook's bytes cannot travel does `databook_dump.py` apply: it writes the
properties `fdd_utils/workbook` reads (cell values and types at their exact
positions, indent, bold, number format, row outline level, hidden rows and columns,
merged ranges, sheet state) and rebuilds a workbook from them. The dump ends with a
digest, the rebuild is read back with the same code and compared, and the rebuilt
file is marked as a rebuild in its properties and beside it. `baseline_pack.py` finds
the run through the file the rebuild came from, accepts it only by the per-account
fingerprint proof, and never calls it complete. Not carried: formulas, fonts and
fills other than bold, borders, widths, comments, charts, pivot tables, defined
names. **Verified offline**: `test_text_transfer.py`, `test_databook_dump.py`
(including that the extractor's own style readers return the same signals from the
rebuild), and a round trip of the four local workbooks (every sheet equal as pandas
reads it, every account's fingerprint equal). Those four hold no indentation, bold or
number formats at all, so the style path has only been exercised on a synthetic file.

### 5.5 The code map

`fdd_utils/agent/codemap.py` indexes the repo's own Python with the standard `ast`
module (no import of the code, no model): every function, method and class with its
signature, docstring, size, callers, callees and the tests that name it; what it
touches (Streamlit, files, a provider, config, global state, python-pptx objects,
PowerPoint COM, pandas); and the comment lines inside it that record a measurement,
a revert or a warning. The agent looks code up with `find_capability`,
`explain_capability` and `code_overview` instead of reading files, and decides for
itself whether to use a piece as it is, wrap it, or rewrite it. The `reuse` field is a
hint built only from those signals; it decides nothing.

Limits that matter: callers, callees and tests are matched by name, so a name defined
in several places lists all of them; tests are looked for under `ad-hoc/` only, and
most of the hand-tuned layout code (wrapping, packing, measurement) has none, so
its behaviour is pinned by its comments and by real PowerPoint output, not by a
test; English words match best.

```
python ad-hoc/workbench/code_map.py find "characters per line wrap width"   free, seconds
python ad-hoc/workbench/code_map.py explain <name> [--source 40]
```

## 6. Deck lineage and returned decks

An exported deck carries a manifest: per commentary paragraph, the account, slide,
shape, part (an account can span slides), a text fingerprint, and for
claim-written paragraphs the ids of the cited facts with a hash of each value.
The embedded part carries no figure. A sidecar `<deck>.lineage.json` holds the
baseline values and paragraph texts and stays with the working files.
`strip_manifest` produces the copy that goes outside.

Each claim-written paragraph also records what it depends on besides its
values, as hashes: what the cited periods permit, how their checks came out, and
the text of the notes a cause cites. A paragraph is flagged when any of these
change even though every figure it quotes is the same.

A returned deck is read against it: unchanged, moved, user-edited, ambiguous
(copied shape) or unlocated; user edits and source changes are told apart, a
conflict gets a sentence-level proposal and is not written, and only
paragraphs the user left alone whose cited facts changed are updated in place,
on a copy. Without a baseline the question "what changed since the last upload"
is refused. Without a manifest nothing is claimed about history.

- **Verified offline**: `test_deck_lineage.py` (copied shapes, reordered slides,
  a hand-edited statement table, a lost tag, a deleted paragraph, a stripped
  manifest, accounts spanning pages, several accounts in one box, a line
  inserted in the source).
- **Not verified**: that real PowerPoint preserves the custom XML part and the
  alt-text tags across a save (python-pptx does). The lost-tag and
  stripped-manifest paths exist so that the answer does not matter for safety.
  Layout after an in-place update was checked with `inspect_pptx`, which
  measures the model against itself; real PowerPoint was not run.
- Statement tables are fingerprinted, so a hand edit is reported
  (`table_edited`) and the table is kept as returned.
- **Not built**: in-place update of the embedded statement tables, reflow of an
  account that spans several paragraphs (reported as `reflow_required`), the
  project bundle (an encrypted, retained snapshot of the working papers). No
  unencrypted snapshot stands in for it.

Hidden sheets and the Financials sheet: one rule, in `workbook/visibility.py`,
applied wherever a Financials sheet is chosen or read. Picking by name keeps only
visible sheets and keeps nothing when sheet visibility cannot be read; naming a
sheet explicitly is refused when it is hidden or cannot be shown not to be; and
the function that reads the Financials sheet refuses a hidden one itself. The
places are the agent session, the CLI's auto-detection, its `--sheet` and
`--financials-sheet`, its roll-up matching, the resolver's reference sheet and
the reader. The UI's sheet list was already visible-only. Visible sheets are
chosen as before (the four local workbooks pick the same sheets). **Verified
offline**: `test_sheet_visibility.py` covers each entry point with hidden and
visible sheets and fails if a module starts matching sheet names against
"financials" without the rule. **Not verified**: a workbook whose sheet states
openpyxl cannot read is treated as unreadable, which this repo has no such file
to test on.

## 7. Switching and rolling back

The agent path is off. `FDD_AGENT=1`, or `agent.enabled: true` in `config.yml`,
turns on the model commands of `fdd_agent.py` (`ask`, `edit`, `chat`, `report`)
and the app panel. The free commands (`summary`, `diff`, `deck`, `candidates`,
`accept`, `export`, `strip`) always run. Routing is `single` unless
`agent.routing: escalate`, `FDD_AGENT_ROUTING` or `--routing` says otherwise.
Nothing in the fixed pipeline reads either setting. Rollback is unsetting them. **Verified offline**: the
flag and the CLI are covered in `test_agent_tools.py` and `test_agent_flows.py`.

## 8. Known limits

- Bare numbers and percentages are not extracted as amounts by the fixed
  pipeline's verifier, so a wrong ratio there is never caught. The claim
  verifier checks percentages that appear with `%`.
- A `data-backed` verdict in the fixed pipeline covers both "the numbers
  matched" and "nothing here was checkable".
- The frequency of verdicts across archived runs (roughly one clause in sixty
  comes back unsupported, about a quarter of those arithmetic) was read from
  archives that are not in this repository. **Not verified** on the current
  revision, and frequency is not accuracy.
- A tab's history-activity test (`any_period_nonzero_by_description`) is read
  from every block on a sheet, so a sheet holding several accounts can look
  active for an account that is not.
- Statement-table values are not part of the claim contract. A change in the
  source reaches the tables only by regenerating the deck.

## 9. Reproducing

```
python ad-hoc/workbench/migration_checks.py [--json out.json]   all offline suites, provider blocked
python ad-hoc/workbench/test_agent_flows.py <databook.xlsx>       five flows on a real workbook (needs the file)
FDD_FLOW_WORKBOOK=<databook.xlsx> FDD_FLOW_BASE_RUN=<run id> python ad-hoc/workbench/test_workspace_flows.py
python ad-hoc/workbench/baseline_pack.py --databook <databook.xlsx> [--out for_test/<pack>]   free; --verify <pack> checks one
python ad-hoc/workbench/claim_eval.py <databook.xlsx>             false accept / reject of the claim verifier
python ad-hoc/workbench/shadow_compare.py <databook.xlsx> --base-run <id>
PYTHONPATH=. python ad-hoc/workbench/replay_verification.py <databook.xlsx> --run <id> --decoys
```
