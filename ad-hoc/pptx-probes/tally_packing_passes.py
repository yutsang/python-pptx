"""Which pass first pushes a slot past capacity?

gen_packing.py writes dp_packing_report.txt on every export, unconditionally,
and each statement gets a per-pass trace of every slot's fill with '!' marking
over-capacity. The answer to "did the DP overflow, or did a later pass do it"
is therefore already on disk after any export -- this just counts it.

    python ad-hoc/pptx-probes/tally_packing_passes.py [dp_packing_report.txt]

Two outcomes mean different work. Overflow already present at 'after DP' is
the DP having climbed its relax ladder because the content does not fit; no
objective change helps. Overflow that appears at a later pass is that pass
pulling content forward into a slot that was fitting a moment earlier, and
those traces are printed in full because they are the ones worth reading.

Reads only. Writes nothing.
"""
import collections
import re
import sys

PASSES = (
    "after DP",
    "rebalance_lopsided_lr_pairs",
    "consolidate_tiny_stub_lr_pairs",
    "rebalance_underfilled_boundaries",
    "rebalance_overflowing_boundaries",
    "maximize_forward_fill",
    "consolidate_trailing_near_empty_slot",
)
_ROW = re.compile(r"\s{2}(\S[^ ]*(?: \S+)*?)\s{2,}(s\d.*)")


def _traces(path):
    """Each 'per-pass slot fill' block as (statement_label, [(pass, rows, over)])."""
    out, cur, label = [], None, "?"
    for line in open(path, encoding="utf-8", errors="replace"):
        if "per-pass slot fill" in line:
            label = line.strip().lstrip("- ").split(":")[0]
            cur = []
            out.append((label, cur))
            continue
        if cur is None:
            continue
        m = _ROW.match(line)
        if m and m.group(1).strip() in PASSES:
            cur.append((m.group(1).strip(), m.group(2).rstrip(), "!" in m.group(2)))
    return [(lab, rows) for lab, rows in out if rows]


def main() -> int:
    path = sys.argv[1] if len(sys.argv) > 1 else "dp_packing_report.txt"
    try:
        traces = _traces(path)
    except OSError as exc:
        print(f"Could not read {path}: {exc}")
        print("It is written into the working directory at export time, so run this")
        print("from wherever the export ran.")
        return 2
    if not traces:
        print(f"{path} holds no per-pass traces. It predates the tracing in")
        print("gen_packing.py, or no statement was packed.")
        return 1

    first = collections.Counter()
    late = []
    clean = 0
    for label, rows in traces:
        hit = next((n for n, _r, over in rows if over), None)
        if hit is None:
            clean += 1
            continue
        first[hit] += 1
        if hit != "after DP":
            late.append((label, rows))

    print(f"{len(traces)} statement trace(s) in {path}")
    print(f"  {clean} never exceed capacity at any pass")
    print(f"  {len(traces) - clean} do -- first pass where '!' appears:\n")
    for name, count in first.most_common():
        tag = "  <- content does not fit; the DP relaxed" if name == "after DP" else \
              "  <- fitted after the DP, then did not"
        print(f"    {count:5d}  {name:<38}{tag}")

    if late:
        print(f"\n{len(late)} statement(s) were inside capacity after the DP and over it"
              f" later.\nThose traces in full:\n")
        for label, rows in late:
            print(f"  --- {label} ---")
            for name, row, _over in rows:
                print(f"    {name:<38} {row}")
            print()
    else:
        print("\nNo statement was pushed over capacity by a pass after the DP.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
