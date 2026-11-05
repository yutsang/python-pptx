"""Print the two packing knobs exactly as the exporter reads them."""
import sys as _sys

# Windows swaps the console encoding away from UTF-8 the moment output is piped
# or redirected: attached to a terminal, printing goes through the console's own
# Unicode writer, but `... | findstr` / `... > out.txt` falls back to
# locale.getpreferredencoding() -- cp1252 on this user's box -- and every CJK
# character this tool prints (主表, 固定资产, the account names themselves) then
# raises UnicodeEncodeError mid-print. It crashed a real `--help | findstr`.
# errors="replace" so a genuinely undisplayable glyph degrades to a placeholder
# rather than killing a 12-minute run on its last line.
for _stream in (_sys.stdout, _sys.stderr):
    if hasattr(_stream, "reconfigure"):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except (ValueError, OSError):
            pass

from fdd_utils.pptx import PowerPointGenerator
import os, time
g = PowerPointGenerator.__new__(PowerPointGenerator)
from fdd_utils.pptx.payloads import _load_pptx_settings
g.pptx_settings = _load_pptx_settings()
cfg = "fdd_utils/config.yml"
print(f"  config.yml 最後修改：{time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(os.path.getmtime(cfg)))}")
for st in (None, "BS", "IS"):
    p = g._packing_settings(st)
    util = float(p.get("shape_height_utilization", 1.00) or 1.00)
    tol  = g._tail_overflow_tolerance_units(st)
    eff  = max(1.05, util)
    print(f"  [{st or 'default'}]  shape_height_utilization={util}  ->  DP 第二階實際用 {eff}"
          f"   |  tail_overflow_tolerance_lines={tol}  (DP 讀不到)")
print(f"\n  DP relax 階梯 = [1.0, max(1.05, util), 1.35, 1.6, 10.0]")
print(f"  26 行的 slot：1.05 -> 27.3 行（+1.3）；若要 +2 行需 util >= {28/26:.3f}")
print(f"  23.9 行的 slot：1.05 -> 25.1 行（+1.2）；若要 +2 行需 util >= {25.9/23.9:.3f}")
