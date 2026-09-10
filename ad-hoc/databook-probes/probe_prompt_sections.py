"""M0 probe: which deterministic guidance sections actually reach the Generator prompt.

Free, no LLM. Prints a PRESENT/ABSENT matrix per account, the list of accounts
where the nested-frame attach was skipped, and two formatting assertions on the
rendered prompt (no scientific notation; trend-summary figures in the same unit
as the analysis table printed above them).

    python ad-hoc/databook-probes/probe_prompt_sections.py <databook.xlsx> [entity] [--before]

--before strips attrs["prompt_analysis_df"] from every account to reproduce the
pre-M0 behaviour, so the same run prints a comparable BEFORE matrix.
"""

import os
import sys

# Runs from the repo root on Windows too: every ad-hoc script documented
# "PYTHONPATH=. python ...", which is Unix shell syntax that cmd.exe rejects
# before python starts. Invoked by path, sys.path[0] is this script's own
# directory, so put the repo root on it here instead of asking the caller.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
import re
import sys

from fdd_utils.financial_display_format import format_in_unit
from fdd_utils.workbook import build_trend_summary, process_workbook_data
from fdd_utils.ai.prompts import PromptEngine

SCIENTIFIC = re.compile(r"\d[eE][+-]\d\d")

# (label, english marker, chinese marker)
MARKERS = [
    # Matched as the section HEADING _append_markdown_section emits (label,
    # optional unit suffix, colon). Two traps: bare "重大变动" also occurs in the
    # shared Chinese prompt prose and inside the 【重大变动提示】 control, so a
    # substring test reads PRESENT on a build that renders no section at all;
    # and the leading blank line cannot be anchored on either, because
    # normalize_english_text collapses it on the Eng side.
    ("trend_summary",   r"Trend summary( \(in [^)\n]*\))?:",         r"趋势摘要(（单位：[^）\n]*）)?:"),
    ("sig_movements",   r"Significant movements( \(in [^)\n]*\))?:", r"重大变动(（单位：[^）\n]*）)?:"),
    ("remainder",       re.escape("[REMAINDER ALREADY COMPUTED]"), re.escape("【余额差额已算好】")),
    ("comp_nature",     re.escape("[COMPONENT NATURE"),            re.escape("【构成项性质")),
    ("hierarchy",       re.escape("[VERIFIED HIERARCHY"),          re.escape("【本科目已核对的层级关系】")),
    # controls: these read df/attrs directly and should be alive today
    ("CTRL material",   re.escape("[MATERIAL MOVEMENT]"),          re.escape("【重大变动提示】")),
    ("CTRL insight",    re.escape("[DATA INSIGHT"),                re.escape("【数据洞察")),
]


def render(pe, language, key, df, data_format):
    sys_p, usr_p = pe.render_prompt("subagent_1", language, key, df, data_format=data_format)
    return (sys_p or "") + "\n" + (usr_p or "")


def prompt_sizes(dfs, language, pe, limits=(8192, 16384, 32768)):
    """How big is each account's Generator prompt, and does it fit?

    This is the only place the question can be answered for an account whose
    call FAILS. A failed call writes no usage record, so the token table in
    inspect_run.py is built entirely from calls that succeeded -- and the
    prompt worth measuring is always the one that was rejected. A real run
    came back "Range of input length should be [1, 32768]" while every prompt
    that report could see topped out at 10,425.

    Same estimator the client uses before every call, so these numbers are
    comparable with the ones the run log prints. Free: renders prompts, calls
    nothing.
    """
    from fdd_utils.ai.client import AIClient
    est = AIClient._estimate_text_tokens

    rows = []
    for key, df in dfs.items():
        try:
            blob = render(pe, language, key, df, "markdown")
        except Exception as exc:  # noqa: BLE001
            rows.append((-1, key, f"RENDER FAILED: {type(exc).__name__}: {exc}", 0))
            continue
        cjk = len(AIClient._CJK_CHAR_RE.findall(blob))
        rows.append((est(blob), key, "", len(blob), cjk / max(len(blob), 1)))
    rows.sort(reverse=True)

    print("\n" + "=" * 78)
    print("  PROMPT SIZE PER ACCOUNT (subagent_1, markdown) — estimated tokens")
    print("=" * 78)
    print(f"  {'tokens':>8}  {'chars':>8}  {'CJK':>5}  account")
    for tokens, key, err, chars, share in rows:
        if err:
            print(f"  {'-':>8}  {'-':>8}  {'-':>5}  {key}  {err}")
            continue
        over = [f"{lim // 1024}k" for lim in limits if tokens > lim]
        flag = ("   *** OVER " + ", ".join(over) + " ***") if over else ""
        print(f"  {tokens:>8,}  {chars:>8,}  {share:>4.0%}  {key}{flag}")
    real = [r for r in rows if r[0] >= 0]
    if real:
        top = real[0]
        # The CJK share is printed because it is the one number the estimate
        # turns on and the one nobody had measured: len/4 assumed there was no
        # CJK at all, and put a prompt the provider rejected at 32,768 tokens
        # comfortably inside it. See AIClient._estimate_text_tokens.
        print(f"\n  largest: {top[1]} at {top[0]:,} tokens ({top[3]:,} chars, {top[4]:.0%} CJK).")
        print("  A prompt over the provider's input limit comes back as 'Range of input")
        print("  length exceeds limited' and that account ships a deterministic bullet.")


def matrix(dfs, language, pe, keys):
    rows = []
    for k in keys:
        df = dfs.get(k)
        try:
            blob = render(pe, language, k, df, "markdown")
        except Exception as exc:  # noqa: BLE001
            rows.append((k, f"RENDER FAILED: {type(exc).__name__}: {exc}", None))
            continue
        hits = {}
        for label, eng, chi in MARKERS:
            hits[label] = ("E" if re.search(eng, blob) else "") + ("C" if re.search(chi, blob) else "")
        rows.append((k, hits, len(blob)))
    return rows


def show(title, rows):
    print(f"\n=== {title} ===")
    labels = [m[0] for m in MARKERS]
    print(f"{'account':<28} " + " ".join(f"{l:>14}" for l in labels) + "   chars")
    for k, hits, size in rows:
        if isinstance(hits, str):
            print(f"{k:<28} {hits}")
            continue
        cells = " ".join(f"{(hits[l] or '-'):>14}" for l in labels)
        print(f"{k:<28} {cells}   {size}")


def check_formatting(dfs, language, pe):
    """Assertion 1: no scientific-notation token anywhere in the rendered prompt.
    Assertion 2: trend-summary figures are in the same unit as the analysis table.
    """
    sci_hits = []
    unit_hits = []
    checked = 0
    for key, df in dfs.items():
        blobs = {}
        for data_format in ("markdown", "json"):
            try:
                blobs[data_format] = render(pe, language, key, df, data_format)
            except Exception as exc:  # noqa: BLE001
                sci_hits.append((key, f"RENDER FAILED ({data_format}): {type(exc).__name__}: {exc}"))
        for data_format, blob in blobs.items():
            found = SCIENTIFIC.findall(blob)
            if found:
                sci_hits.append((key, f"{data_format}: {sorted(set(found))[:4]}"))

        analysis_df = pe._build_analysis_prompt_df(df)
        if analysis_df is None or analysis_df.empty:
            continue
        formatted = pe._format_analysis_prompt_df(analysis_df, language)
        divisor = float(formatted.attrs.get("display_unit_divisor") or 1.0)
        decimals = int(formatted.attrs.get("display_unit_decimals") or 0)
        raw = build_trend_summary(analysis_df)
        if not raw:
            continue
        checked += 1
        blob = blobs.get("markdown", "")
        for field in ("start_value", "end_value", "net_change"):
            value = raw.get(field)
            if value is None:
                continue
            expected = format_in_unit(value, divisor, decimals)
            if f"{field}: {expected}" not in blob:
                unit_hits.append((key, f"{field} expected {expected!r}, not found"))
            elif divisor != 1.0 and f"{field}: {value}" in blob:
                unit_hits.append((key, f"{field} still printed in base units ({value})"))

    print(f"\nASSERT no scientific notation (\\d[eE][+-]\\d\\d): "
          f"{'FAIL' if sci_hits else 'PASS'} ({len(sci_hits)} account-renders hit)")
    for key, detail in sci_hits[:10]:
        print(f"    {key}: {detail}")
    print(f"ASSERT trend summary in the analysis table's unit: "
          f"{'FAIL' if unit_hits else 'PASS'} ({checked} accounts with a trend summary)")
    for key, detail in unit_hits[:10]:
        print(f"    {key}: {detail}")


def main(path, entity="", before=False):
    state = process_workbook_data(temp_path=path, entity_name=entity, selected_sheet=None)
    dfs = state["dfs"]
    language = state.get("language") or "Eng"
    print(f"language={language}  accounts={len(dfs)}  mode={'BEFORE (attach stripped)' if before else 'AFTER (shipped code)'}")
    pe = PromptEngine()

    if before:
        for df in dfs.values():
            df.attrs.pop("prompt_analysis_df", None)

    # The attach is skipped for any account whose normalized payload had no
    # prompt_analysis_df at all (Non-operating income in the reference file):
    # expected, not a regression, so list them instead of demanding 100%.
    skipped = [k for k, df in dfs.items() if df.attrs.get("prompt_analysis_df") is None]
    print(f"nested frame attached on {len(dfs) - len(skipped)}/{len(dfs)} accounts; skipped: {skipped or 'none'}")

    keys = list(dfs.keys())[:8]
    first = dfs[keys[0]]
    print("variant:", first.attrs.get("selected_variant"),
          "| has nested prompt_analysis_df:", first.attrs.get("prompt_analysis_df") is not None,
          "| has significant_movements:", first.attrs.get("significant_movements") is not None,
          "| component_descriptions:", len(first.attrs.get("component_descriptions") or []))

    show("BEFORE (attach stripped)" if before else "AFTER (shipped code)",
         matrix(dfs, language, pe, keys))

    prompt_sizes(dfs, language, pe)

    check_formatting(dfs, language, pe)

    # nesting must not make a frame uncopyable (Python 3.13 deepcopy recursion)
    import copy
    d = dfs[keys[0]]
    for op, fn in (("copy", lambda x: x.copy()),
                   ("deepcopy", lambda x: copy.deepcopy(x)),
                   ("reset_index", lambda x: x.reset_index(drop=True))):
        try:
            out = fn(d)
            print(f"{op}: ok, nested preserved =", out.attrs.get("prompt_analysis_df") is not None)
        except RecursionError:
            print(f"{op}: RecursionError")

    # the invariants the plan names, on the accounts where the attach happened
    bad = []
    for k, df in dfs.items():
        nested = df.attrs.get("prompt_analysis_df")
        if nested is None:
            continue
        if nested is df:
            bad.append((k, "self-reference"))
        elif "prompt_analysis_df" in nested.attrs:
            bad.append((k, "nested frame carries its own nested frame"))
    print(f"ASSERT no attrs cycle: {'FAIL ' + str(bad) if bad else 'PASS'}")


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a != "--before"]
    main(args[0], args[1] if len(args) > 1 else "", before="--before" in sys.argv)
