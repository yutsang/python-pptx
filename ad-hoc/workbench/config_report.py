"""Free; no model call; never prints a secret. What this machine's config.yml says, set against config.example.yml, so two machines can be
aligned by reading one short report from each.

    python ad-hoc/workbench/config_report.py                      read-only report (paste it back)
    python ad-hoc/workbench/config_report.py --set-flash-default  also make Flash the model a new session starts on (backs up first)
    python ad-hoc/workbench/config_report.py --drop-ignored-blocks  delete the earlier copy of a top-level key written twice (backs up first)
    python ad-hoc/workbench/config_report.py --add-section agent  append that block from the example to a file that lacks it (backs up first)
    python ad-hoc/workbench/config_report.py --set processing.read_cell_styles=true  change one existing setting by its one line (backs up first)
    python ad-hoc/workbench/config_report.py --section agent pptx also every value of those sections (--verbose: every difference)

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
no client name. --set-flash-default edits two lines (default.ai_provider and default.model) of the file it reports on, after copying it to
config.yml.bak-<time>, reads the file back and checks that a new session now starts on Flash; it refuses when it cannot find exactly one Flash
model in the dashscope list, and puts the copy back when anything else moved. --drop-ignored-blocks does the same for a top-level key
written twice: the app already reads only the last copy, so removing the earlier one changes nothing the app does (the report says what it held).
"""
from __future__ import annotations

import argparse
import datetime
import hashlib
import json
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
        parts.setdefault(re.split(r"[.\[]", path, maxsplit=1)[0], []).append(f"{path}={value!r}")
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
    for role in ("primary", "escalation"):
        r = flag.role_config(config, role)
        lines.append(f"agent.models.{role}: " + (f"{r['provider']}::{r['model']}  thinking={r['thinking']!r} effort={r['reasoning_effort']!r} "
                                                f"max_output={r['max_output']}" if r["provider"] else "not set (the dropdown's model writes)"))
    flags = {"prepared_write": flag.prepared_write(config), "lead_write": flag.lead_write(config), "plan_first": flag.plan_first(config),
             "tool_loop_on_repeat": flag.tool_loop_on_repeat(config), "research": flag.research(config),
             "repair_inference": flag.repair_inference(config), "judge_binding": flag.judge_binding(config),
             "seed_facts": flag.seed_facts(config), "review_meaning": flag.review_meaning(config)}
    lines.append("agent flags: " + "  ".join(f"{k}={v}" for k, v in flags.items()))
    lines.append(f"agent budget: {flag.agent_budget(config)}  max_model_calls (one ask/edit session)={flag.agent_max_model_calls(config)}  "
                 f"max_calls_per_account (a whole book)={flag.max_calls_per_account(config)}")
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


def _backup(path: Path) -> Path:
    """A copy of the file under a name nothing else has: two edits in one command land in the same second, and the second must not
    overwrite the copy of the original with a copy of the first edit."""
    stamp = f"{datetime.datetime.now():%Y%m%d-%H%M%S}"
    target, n = path.with_name(f"{path.name}.bak-{stamp}"), 1
    while target.exists():
        n += 1
        target = path.with_name(f"{path.name}.bak-{stamp}-{n}")
    shutil.copyfile(path, target)
    return target


def starts_on(config: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """The dropdown entry a new session starts on, as the app resolves it."""
    from fdd_utils.ai import list_model_choices
    from fdd_utils.ai.config import default_model_choice_key
    choices = list_model_choices(config)
    key = default_model_choice_key(config, choices)
    return next((c for c in choices if c["key"] == key), None)


def _replace_in_block(lines: List[str], path: List[str], rendered: str) -> Tuple[Optional[int], str]:
    """The index of the one line that holds the setting at <path> (['processing', 'read_cell_styles'], or deeper), found by walking the
    indentation from the last copy of the top-level key, and what that line would read with the value changed to <rendered>
    (a trailing comment stays)."""
    blocks = [blk for blk in top_blocks(lines) if blk[0] == path[0]]
    if not blocks:
        return None, f"no top-level '{path[0]}:' block found"
    _k, lo, hi = blocks[-1]
    walked, child = path[0], ""
    for key in path[1:]:
        child = next((re.match(r"^(\s+)", lines[i]).group(1) for i in range(lo + 1, hi)
                      if lines[i].strip() and lines[i][0] in " \t" and not lines[i].lstrip().startswith("#")), None)
        if child is None:
            return None, f"'{walked}' has no settings under it"
        found = [i for i in range(lo + 1, hi) if re.match(rf"^{re.escape(child)}{re.escape(key)}:(\s|$)", lines[i])]
        if len(found) != 1:
            return None, f"expected one '{key}:' line directly under {walked}, found {len(found)}"
        line = found[0]
        end = next((j for j in range(line + 1, hi) if lines[j].strip() and not lines[j].lstrip().startswith("#")
                    and len(lines[j]) - len(lines[j].lstrip()) <= len(child)), hi)
        lo, hi, walked = line, end, f"{walked}.{key}"
    comment = re.search(r"\s+#.*$", lines[lo])
    return lo, f'{child}{path[-1]}: {rendered}{comment.group(0) if comment else ""}\n'


def _replace_in_default(lines: List[str], key: str, value: str) -> Tuple[Optional[int], str]:
    return _replace_in_block(lines, ["default", key], f'"{value}"')


def set_flash_default(path: Path) -> str:
    """Make a new session start on the one Flash model of the dashscope list: default.ai_provider and default.model, those two lines only.
    The provider has to move with the model: with another provider there (workbench is not in the dropdown) the app starts on the first
    entry of the list whatever default.model says. Returns what happened."""
    before, _dup = load(path)
    ids = flash_ids(before)
    if len(ids) != 1:
        return f"NOT CHANGED: expected exactly one Flash chat model in dashscope.available_models, found {ids or 'none'}"
    target, provider = ids[0], "dashscope"
    lines = path.read_text(encoding="utf-8").splitlines(keepends=True)
    edits = []
    for key, value in (("ai_provider", provider), ("model", target)):
        i, new = _replace_in_default(lines, key, value)
        if i is None:
            return f"NOT CHANGED: {new}"
        if str((before.get("default") or {}).get(key) or "") != value:      # by what the file says, not by how it is quoted
            edits.append((i, new))
    if not edits:
        now = starts_on(before)
        return (f"already set: default.ai_provider = {provider}, default.model = {target}" if now and now["model_name"] == target
                else f"NOT CHANGED: the two lines already say {provider} / {target} but a new session starts on {now['label'] if now else 'nothing'} "
                     "(is Flash listed: false, or the provider missing from the file?)")
    backup = _backup(path)
    for i, new in edits:
        lines[i] = new
    path.write_text("".join(lines), encoding="utf-8")
    try:
        after, _ = load(path)
    except yaml.YAMLError:
        shutil.copyfile(backup, path)
        return f"RESTORED from {backup.name}: the file did not read after the edit"
    flat_a, flat_b = flatten(before), flatten(after)
    changed = sorted(k for k in set(flat_a) | set(flat_b) if flat_a.get(k) != flat_b.get(k))
    now = starts_on(after)
    if not set(changed) <= {"default.ai_provider", "default.model"} or not now or now["model_name"] != target or now["model_type"] != provider:
        shutil.copyfile(backup, path)
        return (f"RESTORED from {backup.name}: the read-back differed in {changed} and a new session would start on "
                f"{now['label'] if now else 'nothing'}")
    was = (before.get("default") or {})
    return (f"CHANGED default.ai_provider {was.get('ai_provider')!r} -> {provider!r}, default.model {was.get('model')!r} -> {target!r} "
            f"(two lines; a new session starts on {now['label']}; copy kept as {backup.name})")


def top_blocks(lines: List[str]) -> List[Tuple[str, int, int]]:
    """(key, first line, end) of every top-level 'key:' block. Blank lines and column-0 comments after a block belong to the next one."""
    starts = [(i, m.group(1)) for i, ln in enumerate(lines) if (m := re.match(r"^([A-Za-z_][\w.\-]*):", ln))]
    blocks = []
    for n, (s, key) in enumerate(starts):
        end = starts[n + 1][0] if n + 1 < len(starts) else len(lines)
        while end > s + 1 and (not lines[end - 1].strip() or lines[end - 1].startswith("#")):
            end -= 1
        blocks.append((key, s, end))
    return blocks


def ignored_blocks(path: Path) -> List[Tuple[str, int, int, Dict[str, Any]]]:
    """The earlier copies of a top-level key written more than once, which the app never reads: (key, first line, last line, their flat values)."""
    lines = path.read_text(encoding="utf-8").splitlines(keepends=True)
    blocks = top_blocks(lines)
    last = {key: (s, e) for key, s, e in blocks}
    out = []
    for key, s, e in blocks:
        if (s, e) == last[key]:
            continue
        try:
            values = flatten(yaml.safe_load("".join(lines[s:e])) or {})
        except yaml.YAMLError:
            values = {}
        out.append((key, s + 1, e, values))
    return out


def drop_ignored_blocks(path: Path) -> str:
    """Delete the earlier copy of every top-level key written twice. YAML keeps the last, so the app already reads without them: the file is
    parsed again afterwards and put back from the copy unless the parsed result is exactly what it was."""
    before, dups = load(path)
    if not dups:
        return "NOT CHANGED: no top-level key is written twice"
    lines = path.read_text(encoding="utf-8").splitlines(keepends=True)
    blocks = top_blocks(lines)
    last = {key: (s, e) for key, s, e in blocks}
    drop = [(key, s, e) for key, s, e in blocks if (s, e) != last[key]]
    if not drop:
        return "NOT CHANGED: could not find the earlier block in the text"
    backup = _backup(path)
    kept = [ln for i, ln in enumerate(lines) if not any(s <= i < e for _k, s, e in drop)]
    path.write_text("".join(kept), encoding="utf-8")
    try:
        after, dups_after = load(path)
    except yaml.YAMLError:
        after, dups_after = None, ["unreadable"]
    if after != before or dups_after:
        shutil.copyfile(backup, path)
        return f"RESTORED from {backup.name}: the file read differently after the removal"
    return ("REMOVED " + ", ".join(f"the earlier '{k}' block (lines {s + 1}-{e})" for k, s, e in drop)
            + f"; what the app reads is unchanged (copy kept as {backup.name})")


def _render(value: Any) -> str:
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (int, float)):
        return repr(value)
    return json.dumps(str(value), ensure_ascii=False)


def set_value(path: Path, assignment: str) -> str:
    """--set section.key=value (any depth, dots between): change one existing setting (a true/false, a number or a
    word), by that one line. The file is read again afterwards and put back from the copy unless that setting is the only thing that moved."""
    if "=" not in assignment or "." not in assignment.split("=", 1)[0]:
        return "NOT CHANGED: write it as section.key=value, for example processing.read_cell_styles=true"
    dotted, raw = assignment.split("=", 1)
    dotted = dotted.strip()
    parts = dotted.split(".")
    if "[" in dotted or not all(parts):
        return "NOT CHANGED: name the setting as section.key (or section.key.key); a list entry cannot be set this way"
    try:
        value = yaml.safe_load(raw)
    except yaml.YAMLError:
        return f"NOT CHANGED: cannot read {raw!r} as a value"
    if isinstance(value, (dict, list)) or value is None:
        return "NOT CHANGED: only a true/false, a number or a word can be set this way"
    if secret(dotted, value):
        return "NOT CHANGED: that setting holds a credential or a host: edit it by hand"
    before, dups = load(path)
    if parts[0] in {d for d, _n in dups}:
        return f"NOT CHANGED: '{parts[0]}' is written twice in the file; settle that first (--drop-ignored-blocks)"
    current: Any = before
    for part in parts:
        current = current.get(part, KeyError) if isinstance(current, dict) else KeyError
        if current is KeyError:
            return f"NOT CHANGED: {dotted} is not in your file (the code default applies); add the line by hand"
    if isinstance(current, (dict, list)):
        return f"NOT CHANGED: {dotted} holds more than one value"
    if current == value and type(current) is type(value):
        return f"already set: {dotted} = {_render(value)}"
    lines = path.read_text(encoding="utf-8").splitlines(keepends=True)
    i, new = _replace_in_block(lines, parts, _render(value))
    if i is None:
        return f"NOT CHANGED: {new}"
    backup = _backup(path)
    lines[i] = new
    path.write_text("".join(lines), encoding="utf-8")
    try:
        after, _ = load(path)
    except yaml.YAMLError:
        shutil.copyfile(backup, path)
        return f"RESTORED from {backup.name}: the file did not read after the edit"
    flat_a, flat_b = flatten(before), flatten(after)
    changed = sorted(k for k in set(flat_a) | set(flat_b) if flat_a.get(k) != flat_b.get(k))
    got: Any = after
    for part in parts:
        got = got.get(part, KeyError) if isinstance(got, dict) else KeyError
    if changed != [dotted] or got != value or type(got) is not type(value):
        shutil.copyfile(backup, path)
        return f"RESTORED from {backup.name}: the read-back moved {changed}"
    return f"CHANGED {dotted}: {_render(current)} -> {_render(value)} (one line; copy kept as {backup.name})"


#: blocks that hold credentials or hosts: never copied from the example by a tool
_CREDENTIAL_SECTIONS = {"dashscope", "workbench", "local", "deepseek", "openai"}


def add_sections(path: Path, example_path: Path, names: List[str]) -> str:
    """--add-section agent: append the example's top-level block (with the comment lines above it) to a file that lacks it, so the file
    says what the code would assume. Left alone when the file already has the key. Read back, and put back from the copy on any difference."""
    before, dups_before = load(path)
    example, _ = load(example_path)
    ex_lines = example_path.read_text(encoding="utf-8").splitlines(keepends=True)
    ex_blocks = {key: (s, e) for key, s, e in top_blocks(ex_lines)}
    notes, add = [], []
    for name in names:
        if name in _CREDENTIAL_SECTIONS:
            notes.append(f"{name}: NOT ADDED (it holds credentials or hosts: yours, by hand)")
        elif name in before:
            notes.append(f"{name}: already in your file, left alone")
        elif name not in ex_blocks:
            notes.append(f"{name}: NOT ADDED (config.example.yml has no such block)")
        else:
            add.append(name)
    if not add:
        return "NOT CHANGED: " + "; ".join(notes)
    text = path.read_text(encoding="utf-8")
    chunk = []
    for name in add:
        s, e = ex_blocks[name]
        c = s
        while c > 0 and ex_lines[c - 1].startswith("#"):
            c -= 1
        chunk.append("".join(ex_lines[c:e]).rstrip("\n") + "\n")
    backup = _backup(path)
    path.write_text(text.rstrip("\n") + "\n\n" + "\n".join(chunk), encoding="utf-8")
    try:
        after, dups = load(path)
    except yaml.YAMLError:
        after, dups = None, [("unreadable", 0)]
    expected = dict(before)
    expected.update({name: example[name] for name in add})
    if after != expected or {k for k, _n in dups} != {k for k, _n in dups_before}:
        shutil.copyfile(backup, path)
        return f"RESTORED from {backup.name}: the file read differently after the append"
    return ("ADDED " + ", ".join(add) + f" from config.example.yml (copy kept as {backup.name})" + ("; " + "; ".join(notes) if notes else ""))


def top(path: str) -> str:
    return re.split(r"[.\[]", path, maxsplit=1)[0]


def report(path: Path, example_path: Path, *, verbose: bool = False, section: Any = "") -> str:
    mine_cfg, dups = load(path)
    example_cfg, _ = load(example_path)
    mine, example = flatten(mine_cfg), flatten(example_cfg)
    age = datetime.datetime.now() - datetime.datetime.fromtimestamp(path.stat().st_mtime)
    out = ["CONFIG REPORT (no secrets, hosts or client names in it)", "", "FILE",
           f"  {path.name}: {path.stat().st_size} bytes, edited {age.days}d {age.seconds // 3600}h ago, {len(mine)} settings",
           f"  read from {'the repo folder' if path.resolve().parent == (ROOT / 'fdd_utils').resolve() else str(path.resolve().parent)}"]
    out.append("  duplicate keys (YAML keeps the last, the earlier block is ignored): "
               + (", ".join(f"'{k}' again at line {n}" for k, n in dups) if dups else "none"))
    for key, first, last_line, values in ignored_blocks(path) if dups else []:
        active = {p_: v for p_, v in mine.items() if top(p_) == key}
        only_there = sorted(p_ for p_ in values if p_ not in active)
        other = sorted(p_ for p_ in values if p_ in active and values[p_] != active[p_])
        out.append(f"    ignored '{key}' block, lines {first}-{last_line}: {len(values)} setting(s); not in the block the app reads: "
                   + (", ".join(f"{p_}={shown(p_, values[p_])}" for p_ in only_there[:12]) + (" ..." if len(only_there) > 12 else "") if only_there else "none")
                   + ("; same key, other value: " + ", ".join(f"{p_}={shown(p_, values[p_])} (read: {shown(p_, active[p_])})" for p_ in other[:8]) if other else ""))
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
    sections = [section] if isinstance(section, str) and section else list(section or [])
    for name in sections:
        out += ["", f"SECTION {name} in this file, every value (secrets masked)"]
        out += [f"  {p_} = {shown(p_, v)}" for p_, v in sorted(mine.items()) if top(p_) == name]
    todo = []
    if dups:
        todo.append("a top-level key is written twice: read the 'ignored' line above; if it holds nothing you want, re-run with --drop-ignored-blocks")
    ids = flash_ids(mine_cfg)
    now = starts_on(mine_cfg)
    if len(ids) == 1 and not (now and now["model_name"] == ids[0]):
        todo.append(f"a new session starts on {now['label'] if now else 'nothing'}, not Flash: re-run with --set-flash-default "
                    f"(it sets default.ai_provider to dashscope and default.model to {ids[0]!r})")
    if not ids:
        todo.append("dashscope.available_models has no Flash chat model")
    out += ["", "TO CHANGE" + ("" if todo else ": nothing the report can see")] + [f"  - {t}" for t in todo]
    return "\n".join(out)


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--config", default=str(ROOT / "fdd_utils" / "config.yml"))
    ap.add_argument("--example", default=str(ROOT / "fdd_utils" / "config.example.yml"))
    ap.add_argument("--set-flash-default", action="store_true")
    ap.add_argument("--drop-ignored-blocks", action="store_true", help="delete the earlier copy of a top-level key written twice")
    ap.add_argument("--add-section", nargs="+", default=[], metavar="NAME", help="append these top-level blocks from the example to a file that lacks them (agent, analysis)")
    ap.add_argument("--set", dest="assignments", nargs="+", default=[], metavar="SECTION.KEY=VALUE", help="change one existing setting by its one line")
    ap.add_argument("--verbose", action="store_true", help="list every missing / different / only-here setting")
    ap.add_argument("--section", nargs="*", default=[], help="list every value of these top-level sections (agent, dashscope, processing, ...)")
    args = ap.parse_args(argv)
    path = Path(args.config)
    if not path.exists():
        print(f"no config at {path}")
        return 1
    if args.drop_ignored_blocks:
        print(drop_ignored_blocks(path) + "\n")
    if args.set_flash_default:
        print(set_flash_default(path) + "\n")
    if args.add_section:
        print(add_sections(path, Path(args.example), args.add_section) + "\n")
    for assignment in args.assignments:
        print(set_value(path, assignment) + "\n")
    print(report(path, Path(args.example), verbose=args.verbose, section=args.section))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
