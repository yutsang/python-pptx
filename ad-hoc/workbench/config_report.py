"""Free; no model call; never prints a secret. What this machine's config.yml says, set against config.example.yml, so two machines can be
aligned by reading one short report from each.

    python ad-hoc/workbench/config_report.py                      read-only report (paste it back)
    python ad-hoc/workbench/config_report.py --set-flash-default  also make Flash the model a new session starts on (backs up first)
    python ad-hoc/workbench/config_report.py --section agent      also every value of one section (--verbose: every difference)

What the report holds:
    FILE        which config.yml is read, its size and age, duplicate top-level keys (YAML keeps only the last one, silently)
    ENVIRONMENT the library versions and the per-machine files git does not carry (config.yml, template.pptx), by size and a short hash
    EFFECTIVE   what the app will do with this file: the model a new session starts on, the dropdown in order with ready / not ready,
                the agent switch and the models it writes with, the agent flags that are not at their code default
    VS EXAMPLE  keys the example has and this file lacks (the code default applies, shown), values that differ, keys only this file has
                (a typo shows up here: it is read by nothing), and the settings that hold a secret as set / empty only
    SECTIONS    one short hash per top-level section over its non-secret values: two machines with the same hash have the same section

A value is masked (shown as set / empty and its length) when its key is api_key, api_base, charge_code, region_override, header_names,
a token or a secret, or when it looks like a URL or a key. Nothing else is masked, so a pasted report holds no credential, no host and
no client name. --set-flash-default edits one line (default.model) of the file it reports on, after copying it to config.yml.bak-<time>,
and reads the file back; it refuses when it cannot find exactly one Flash model in the dashscope list.
"""
from __future__ import annotations

import argparse
import datetime
import hashlib
import os
import platform
import re
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import yaml

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

SECRET_KEYS = {"api_key", "apikey", "api_base", "base_url", "endpoint", "charge_code", "region_override", "header_names", "secret",
               "password", "passwd", "token", "access_token", "unlock_secret", "authorization"}
_URLISH = re.compile(r"://|^sk-|^[A-Za-z0-9_\-]{32,}$")


class _Loader(yaml.SafeLoader):
    duplicates: List[Tuple[str, int]] = []


def _construct(loader: "_Loader", node: yaml.MappingNode, deep: bool = False) -> Dict[Any, Any]:
    seen: Dict[Any, int] = {}
    for key_node, _value in node.value:
        key = loader.construct_object(key_node, deep=True)
        if key in seen:
            loader.duplicates.append((str(key), key_node.start_mark.line + 1))
        seen[key] = key_node.start_mark.line + 1
    return yaml.SafeLoader.construct_mapping(loader, node, deep)


_Loader.add_constructor(yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG, _construct)


def load(path: Path) -> Tuple[Dict[str, Any], List[Tuple[str, int]]]:
    loader = _Loader(path.read_text(encoding="utf-8"))
    loader.duplicates = []
    try:
        data = loader.get_single_data() or {}
    finally:
        loader.dispose()
    return data, list(loader.duplicates)


def secret(path: str, value: Any) -> bool:
    last = path.rsplit(".", 1)[-1].lower()
    if last in SECRET_KEYS or any(part.lower() in ("header_names",) for part in path.split(".")):
        return True
    if last in {"model", "chat_model", "id", "label", "name", "provider", "ai_provider", "api_version"}:
        return False                                   # a model id is as long as a key and is not one
    return isinstance(value, str) and bool(_URLISH.search(value))


def flatten(obj: Any, prefix: str = "") -> Dict[str, Any]:
    """path -> leaf. A list of mappings with an id is keyed by that id (the order of a dropdown is reported on its own)."""
    out: Dict[str, Any] = {}
    if isinstance(obj, dict):
        for k, v in obj.items():
            out.update(flatten(v, f"{prefix}.{k}" if prefix else str(k)))
        if not obj and prefix:
            out[prefix] = {}
    elif isinstance(obj, list):
        if obj and all(isinstance(x, dict) and "id" in x for x in obj):
            for x in obj:
                out.update(flatten({k: v for k, v in x.items() if k != "id"}, f"{prefix}[{x['id']}]"))
            out[f"{prefix}#ids"] = [str(x["id"]) for x in obj]
        else:
            out[prefix] = list(obj)
    else:
        out[prefix] = obj
    return out


def shown(path: str, value: Any) -> str:
    if secret(path, value):
        if value in (None, "", {}, []):
            return "empty"
        n = len(str(value)) if not isinstance(value, dict) else len(value)
        return f"set ({n})"
    return repr(value)


def section_hashes(flat: Dict[str, Any]) -> Dict[str, Tuple[str, int]]:
    parts: Dict[str, List[str]] = {}
    for path, value in sorted(flat.items()):
        if secret(path, value):
            continue
        parts.setdefault(re.split(r"[.\[]", path, 1)[0], []).append(f"{path}={value!r}")
    return {k: (hashlib.sha256("\n".join(v).encode("utf-8")).hexdigest()[:8], len(v)) for k, v in parts.items()}


def short_hash(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()[:10] if path.exists() else "-"


def environment() -> List[str]:
    lines = [f"os {platform.system()} {platform.release()}  python {sys.version.split()[0]}"]
    libs = []
    for name in ("pandas", "openpyxl", "pptx", "yaml", "openai", "httpx", "streamlit"):
        try:
            mod = __import__(name)
            libs.append(f"{name} {getattr(mod, '__version__', '?')}")
        except Exception:
            libs.append(f"{name} MISSING")
    lines.append("  ".join(libs))
    try:
        import pandas as pd
        lines.append(f"pandas new string dtype: {bool(getattr(pd.options.future, 'infer_string', False))}")
    except Exception:
        pass
    try:
        head = subprocess.run(["git", "rev-parse", "--short", "HEAD"], cwd=ROOT, capture_output=True, text=True, timeout=10).stdout.strip()
        dirty = bool(subprocess.run(["git", "status", "--porcelain", "--untracked-files=no"], cwd=ROOT, capture_output=True, text=True,
                                    timeout=10).stdout.strip())
        lines.append(f"git {head or '?'}{' (uncommitted changes)' if dirty else ''}")
    except Exception:
        lines.append("git: not available")
    try:
        sys.path.insert(0, str(ROOT / "ad-hoc" / "workbench"))
        import entity_trial
        lines.append(f"code fingerprint {entity_trial.code_fingerprint()['short']}")
    except Exception as exc:                                              # noqa: BLE001
        lines.append(f"code fingerprint: not available ({type(exc).__name__})")
    for label, rel in (("template.pptx", "fdd_utils/template.pptx"), ("font metrics eng", "fdd_utils/font_metrics/arial_eng.json"),
                       ("font metrics chi", "fdd_utils/font_metrics/msyh_chi.json")):
        p = ROOT / rel
        lines.append(f"{label}: " + (f"{p.stat().st_size} bytes, sha {short_hash(p)}" if p.exists() else "MISSING"))
    return lines


def effective(config: Dict[str, Any]) -> List[str]:
    from fdd_utils.agent import flag
    from fdd_utils.ai import list_model_choices
    from fdd_utils.ai.config import default_model_choice_key, is_provider_ready
    lines = []
    default = config.get("default") or {}
    choices = list_model_choices(config)
    start = default_model_choice_key(config, choices)
    lines.append(f"default.ai_provider={default.get('ai_provider')!r}  default.model={default.get('model')!r}")
    lines.append("dropdown, in order (* = a new session starts on it):")
    for c in choices:
        lines.append(f"  {'*' if c['key'] == start else ' '} {c['label']:<26s} {c['key']:<52s} {'ready' if c['ready'] else 'NOT READY (api_key / api_base empty)'}")
    env = os.environ.get("FDD_AGENT", "")
    lines.append(f"agent.enabled={flag.agent_enabled(config)}  (config says {bool((config.get('agent') or {}).get('enabled', False))}; "
                 f"environment FDD_AGENT={env!r})")
    for role in ("primary", "escalation"):
        r = flag.role_config(config, role)
        lines.append(f"agent.models.{role}: " + (f"{r['provider']}::{r['model']}  thinking={r['thinking']!r} effort={r['reasoning_effort']!r} "
                                                f"max_output={r['max_output']}" if r["provider"] else "not set (the dropdown's model writes)"))
    flags = {"prepared_write": flag.prepared_write(config), "lead_write": flag.lead_write(config), "plan_first": flag.plan_first(config),
             "tool_loop_on_repeat": flag.tool_loop_on_repeat(config), "research": flag.research(config),
             "repair_inference": flag.repair_inference(config), "judge_binding": flag.judge_binding(config),
             "seed_facts": flag.seed_facts(config), "review_meaning": flag.review_meaning(config)}
    lines.append("agent flags: " + "  ".join(f"{k}={v}" for k, v in flags.items()))
    lines.append(f"agent budget: {flag.agent_budget(config)}  max_model_calls={flag.agent_max_model_calls(config)}")
    return lines


def compare(mine: Dict[str, Any], example: Dict[str, Any]) -> Dict[str, List[str]]:
    out: Dict[str, List[str]] = {"missing": [], "different": [], "only_mine": [], "secrets": []}
    keys = sorted(set(mine) | set(example))
    for path in keys:
        if path.endswith("#ids"):
            if path in mine and path in example and mine[path] != example[path]:
                out["different"].append(f"{path[:-4]} order/members: yours {mine[path]}  example {example[path]}")
            continue
        a, b = mine.get(path, KeyError), example.get(path, KeyError)
        masked = secret(path, a if a is not KeyError else b) or secret(path, b if b is not KeyError else a)
        if masked:
            if a is not KeyError:
                out["secrets"].append(f"{path}: {shown(path, a)}")
            elif b is not KeyError and b not in (None, "", {}):
                out["missing"].append(f"{path}: (a credential or host; example {shown(path, b)})")
            continue
        if a is KeyError:
            out["missing"].append(f"{path}: example {b!r}")
        elif b is KeyError:
            out["only_mine"].append(f"{path}: {a!r}")
        elif a != b:
            out["different"].append(f"{path}: yours {a!r}  example {b!r}")
    return out


def flash_ids(config: Dict[str, Any]) -> List[str]:
    block = config.get("dashscope") or {}
    return [str(e.get("id")) for e in block.get("available_models") or [] if isinstance(e, dict) and e.get("kind", "chat") == "chat"
            and "flash" in str(e.get("id")).lower()]


def set_flash_default(path: Path) -> str:
    """Make default.model the one Flash model of the dashscope list, editing that one line only. Returns what happened."""
    before, _dup = load(path)
    ids = flash_ids(before)
    if len(ids) != 1:
        return f"NOT CHANGED: expected exactly one Flash chat model in dashscope.available_models, found {ids or 'none'}"
    target = ids[0]
    lines = path.read_text(encoding="utf-8").splitlines(keepends=True)
    start = next((i for i, ln in enumerate(lines) if re.match(r"^default:\s*(#.*)?$", ln)), None)
    if start is None:
        return "NOT CHANGED: no top-level 'default:' block found"
    end = next((i for i in range(start + 1, len(lines)) if lines[i].strip() and not lines[i].startswith((" ", "\t", "#"))), len(lines))
    idx = [i for i in range(start + 1, end) if re.match(r"^\s+model:\s", lines[i])]
    if len(idx) != 1:
        return f"NOT CHANGED: expected one 'model:' line inside default:, found {len(idx)}"
    old = lines[idx[0]]
    indent = re.match(r"^(\s+)", old).group(1)
    comment = re.search(r"\s+#.*$", old)
    new = f'{indent}model: "{target}"{comment.group(0) if comment else ""}\n'
    if old == new:
        return f"already set: default.model = {target}"
    backup = path.with_name(f"{path.name}.bak-{datetime.datetime.now():%Y%m%d-%H%M%S}")
    shutil.copyfile(path, backup)
    lines[idx[0]] = new
    path.write_text("".join(lines), encoding="utf-8")
    after, _ = load(path)
    flat_a, flat_b = flatten(before), flatten(after)
    changed = sorted(k for k in set(flat_a) | set(flat_b) if flat_a.get(k) != flat_b.get(k))
    if changed != ["default.model"] or after["default"]["model"] != target:
        shutil.copyfile(backup, path)
        return f"RESTORED from {backup.name}: the read-back differed in {changed}"
    return f"CHANGED default.model: {before.get('default', {}).get('model')!r} -> {target!r} (one line; copy kept as {backup.name})"


def top(path: str) -> str:
    return re.split(r"[.\[]", path, 1)[0]


def report(path: Path, example_path: Path, *, verbose: bool = False, section: str = "") -> str:
    mine_cfg, dups = load(path)
    example_cfg, _ = load(example_path)
    mine, example = flatten(mine_cfg), flatten(example_cfg)
    age = datetime.datetime.now() - datetime.datetime.fromtimestamp(path.stat().st_mtime)
    out = ["CONFIG REPORT (no secrets, hosts or client names in it)", "", "FILE",
           f"  {path.name}: {path.stat().st_size} bytes, edited {age.days}d {age.seconds // 3600}h ago, {len(mine)} settings",
           f"  read from {'the repo folder' if path.resolve().parent == (ROOT / 'fdd_utils').resolve() else str(path.resolve().parent)}"]
    out.append("  duplicate keys (YAML keeps the last, the earlier block is ignored): "
               + (", ".join(f"'{k}' again at line {n}" for k, n in dups) if dups else "none"))
    out += ["", "ENVIRONMENT"] + [f"  {x}" for x in environment()]
    out += ["", "EFFECTIVE"] + [f"  {x}" for x in effective(mine_cfg)]
    diff = compare(mine, example)
    # the example leaves out per-model prices on purpose (they are a price sheet, written per machine): not a difference
    diff["only_mine"] = [x for x in diff["only_mine"] if not re.search(r"\.price_(in|out|cached_in):", x)]
    hashes = section_hashes(mine)
    counts = {k: {n: sum(1 for x in diff[n] if top(x.split(":")[0].split(" ")[0]) == k) for n in ("missing", "different", "only_mine")}
              for k in sorted(set(hashes) | {top(x.split(":")[0].split(" ")[0]) for n in ("missing", "different", "only_mine") for x in diff[n]})}
    out += ["", "SECTIONS (same hash on two machines = same non-secret values)    missing / different / only here, against config.example.yml"]
    for k, c in counts.items():
        h, n = hashes.get(k, ("-", 0))
        out.append(f"  {k:<14s} {h:<9s} {n:>3d} values   {c['missing']:>3d} / {c['different']:>3d} / {c['only_mine']:>3d}")
    shown_different = diff["different"] if verbose else diff["different"][:25]
    out += ["", f"DIFFERENT from config.example.yml (not secret): {len(diff['different'])}"] + [f"  {x}" for x in shown_different]
    if len(shown_different) < len(diff["different"]):
        out.append(f"  ... {len(diff['different']) - len(shown_different)} more (--verbose)")
    out += ["", f"SECRET OR HOST SETTINGS (set / empty only): {len(diff['secrets'])}"] + [f"  {x}" for x in diff["secrets"]]
    if verbose:
        for title, key in (("MISSING from this file (the code default applies)", "missing"), ("ONLY in this file", "only_mine")):
            out += ["", f"{title}: {len(diff[key])}"] + [f"  {x}" for x in diff[key]]
    if section:
        out += ["", f"SECTION {section} in this file, every value (secrets masked)"]
        out += [f"  {p_} = {shown(p_, v)}" for p_, v in sorted(mine.items()) if top(p_) == section]
    todo = []
    if dups:
        todo.append("merge the duplicate top-level block(s) named above into one")
    ids = flash_ids(mine_cfg)
    default_model = str((mine_cfg.get("default") or {}).get("model") or "")
    if len(ids) == 1 and default_model != ids[0]:
        todo.append(f"default.model is {default_model!r}: re-run with --set-flash-default (or edit that one line to {ids[0]!r})")
    if not ids:
        todo.append("dashscope.available_models has no Flash chat model")
    out += ["", "TO CHANGE" + ("" if todo else ": nothing the report can see")] + [f"  - {t}" for t in todo]
    return "\n".join(out)


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--config", default=str(ROOT / "fdd_utils" / "config.yml"))
    ap.add_argument("--example", default=str(ROOT / "fdd_utils" / "config.example.yml"))
    ap.add_argument("--set-flash-default", action="store_true")
    ap.add_argument("--verbose", action="store_true", help="list every missing / different / only-here setting")
    ap.add_argument("--section", default="", help="list every value of one top-level section (agent, dashscope, processing, ...)")
    args = ap.parse_args(argv)
    path = Path(args.config)
    if not path.exists():
        print(f"no config at {path}")
        return 1
    if args.set_flash_default:
        print(set_flash_default(path) + "\n")
    print(report(path, Path(args.example), verbose=args.verbose, section=args.section))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
