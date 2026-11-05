"""A map of the code, so the agent looks a capability up instead of reading files.

Much of this repo's Python was tuned over a long time against real PowerPoint output (wrapping, character counts, packing) and cannot be
re-derived from a prompt. This index lets an agent find what already exists, see how far to trust it, and decide: use it as it is, wrap
it, or rewrite it. It reads no file the agent has not asked about, and the answer is a few hundred tokens, not a module.

For each function, method and class (standard library `ast`, no import of the code, no model):
  what it is     qualified name, file:line, signature, first paragraph of its docstring, size in lines
  how it is used callers and callees (by name; a name defined in several places is counted as ambiguous), tests that mention it
  what it touches  signals: streamlit, files, a model/provider, config, global state, python-pptx objects, PowerPoint COM, pandas
  how it was earned the comment lines inside it that record a measurement, a revert or a warning ("tried", "do not", "measured", ...)

Nothing here decides. `reuse` is a hint built from those signals, to be read together with the markers and the source when it matters.
"""
from __future__ import annotations

import ast
import re
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Set, Tuple

ROOT = Path(__file__).resolve().parents[2]
SKIP_DIRS = {"__pycache__", ".git", ".venv", "venv", "logs", "output", "for_test", "node_modules", "ad-hoc", "docs", "pptx_previews"}

_MARKER = re.compile(r"(?i)\b(failed attempts?|tried|revert(?:ed)?|do not|don't|never|measured|verified|ground truth|regress(?:ion)?|root cause|"
                     r"real powerpoint|powerpoint|bug|because|why)\b")
_SIGNALS = {
    "streamlit": re.compile(r"\bst\.[a-z_]+\(|import streamlit"),
    "files": re.compile(r"\bopen\(|\.read_text\(|\.write_text\(|\.read_bytes\(|\.write_bytes\(|shutil\.|os\.(?:remove|rename|replace)\(|\.mkdir\("),
    "model_or_network": re.compile(r"AIClient|get_response\(|requests\.|httpx\.|urlopen\(|openai"),
    "config": re.compile(r"load_config|config\.yml|yaml\.safe_load|\bos\.environ\b"),
    "global_state": re.compile(r"^\s*global\s|\b_[A-Z_]*(?:CACHE|STATE|REGISTRY|SEEN)\b", re.M),
    "pptx_objects": re.compile(r"\bPresentation\(|\bpptx\b|\bshape\b|\btext_frame\b"),
    "powerpoint_com": re.compile(r"win32com|Dispatch\("),
    "pandas": re.compile(r"\bpd\.|DataFrame"),
}
_STOP = {"the", "a", "an", "of", "to", "in", "for", "and", "or", "is", "it", "that", "this", "with", "on", "by", "as", "be", "are", "at", "from"}


@dataclass
class Symbol:
    qname: str
    kind: str                                  # function | method | class
    path: str
    line: int
    end: int
    signature: str
    doc: str
    calls: Set[str] = field(default_factory=set)
    signals: List[str] = field(default_factory=list)
    markers: List[str] = field(default_factory=list)
    tests: List[str] = field(default_factory=list)
    callers: List[str] = field(default_factory=list)
    ambiguous: bool = False                    # the simple name is defined in more than one place, so callers/tests may belong to another
    module_doc: str = ""

    @property
    def simple(self) -> str:
        return self.qname.rsplit(".", 1)[-1]

    @property
    def lines(self) -> int:
        return self.end - self.line + 1

    def summary(self, n: int = 200) -> str:
        first = (self.doc or "").strip().split("\n\n")[0].replace("\n", " ")
        return first[:n]

    def reuse(self) -> str:
        """A hint from the signals only. It is not a verdict: read the markers and, when it matters, the source."""
        risky = [s for s in self.signals if s in ("streamlit", "model_or_network", "powerpoint_com")]
        if self.kind == "class":
            return "class: read explain_capability before using"
        if risky:
            return "wrap or rewrite: it touches " + ", ".join(risky)
        inst = "; a method, so it needs an instance of its class" if self.kind == "method" else ""
        if not self.tests and not self.callers:
            return "unproven: no test and no caller found; read the source first; no test names it" + inst
        touched = [s for s in self.signals if s in ("files", "config", "global_state")]
        if touched:
            verdict = "usable with care: it touches " + ", ".join(touched) + ("; has tests" if self.tests else "; used by others")
        else:
            verdict = "likely usable as it is: " + ("tested" if self.tests else "used by others") + (f", {len(self.callers)} caller(s)" if self.callers else "")
        verdict += inst
        if len(self.markers) >= 3:
            verdict += f"; carries {len(self.markers)} warning/measurement comments: read them before changing its logic"
        if not self.tests:
            verdict += "; no test names it, so its behaviour is pinned only by those comments and by real output"
        return verdict


# ---------------------------------------------------------------------------------------------------------------------------
# building
# ---------------------------------------------------------------------------------------------------------------------------

def _py_files(root: Path) -> Iterable[Path]:
    for p in sorted(root.rglob("*.py")):
        rel = p.relative_to(root)
        if any(part in SKIP_DIRS for part in rel.parts[:-1]) or p.name.startswith("test_"):
            continue
        yield p


def _signature(node: ast.AST) -> str:
    args = node.args  # type: ignore[attr-defined]
    parts: List[str] = []
    positional = list(args.posonlyargs) + list(args.args)
    defaults = [None] * (len(positional) - len(args.defaults)) + list(args.defaults)
    for a, d in zip(positional, defaults):
        text = a.arg + (f": {ast.unparse(a.annotation)}" if a.annotation else "")
        parts.append(text + (" = ..." if d is not None else ""))
    if args.vararg:
        parts.append("*" + args.vararg.arg)
    elif args.kwonlyargs:
        parts.append("*")
    for a, d in zip(args.kwonlyargs, args.kw_defaults):
        parts.append(a.arg + (f": {ast.unparse(a.annotation)}" if a.annotation else "") + (" = ..." if d is not None else ""))
    if args.kwarg:
        parts.append("**" + args.kwarg.arg)
    ret = f" -> {ast.unparse(node.returns)}" if getattr(node, "returns", None) else ""  # type: ignore[attr-defined]
    return f"{node.name}({', '.join(parts)}){ret}"[:240]  # type: ignore[attr-defined]


def _called_names(node: ast.AST) -> Set[str]:
    out: Set[str] = set()
    for sub in ast.walk(node):
        if isinstance(sub, ast.Call):
            f = sub.func
            if isinstance(f, ast.Name):
                out.add(f.id)
            elif isinstance(f, ast.Attribute):
                out.add(f.attr)
    return out


def _markers(source_lines: List[str], node: ast.AST) -> List[str]:
    """Comment lines (and docstring lines) inside the function, plus the comment block directly above it, that record why it is the way it is."""
    start = node.lineno - 1  # type: ignore[attr-defined]
    end = node.end_lineno  # type: ignore[attr-defined]
    above = start
    while above > 0 and source_lines[above - 1].strip().startswith("#"):
        above -= 1
    found: List[str] = []
    for line in source_lines[above:end]:
        text = line.strip()
        if text.startswith("#") or (text and not text.startswith(("def ", "class ", "return ", "if ", "for ", "while ", "else", "elif ", "try", "except", "with "))
                                    and re.match(r"^[A-Za-z一-鿿\"'(]", text) and " " in text and not re.search(r"[=\[\]{}]\s*$", text)):
            if _MARKER.search(text) and len(text) > 12:
                found.append(text.lstrip("# ").strip()[:150])
    return found[:12]


class CodeMap:
    def __init__(self, root: Path = ROOT, tests_dirs: Optional[List[Path]] = None):
        self.root = Path(root)
        self.symbols: Dict[str, Symbol] = {}
        self.modules: Dict[str, str] = {}
        self._build(tests_dirs if tests_dirs is not None else [self.root / "ad-hoc"])

    # -- construction ------------------------------------------------------------------------------------------------------

    def _build(self, tests_dirs: List[Path]) -> None:
        for path in _py_files(self.root):
            rel = path.relative_to(self.root).as_posix()
            try:
                source = path.read_text(encoding="utf-8")
                tree = ast.parse(source)
            except (SyntaxError, UnicodeDecodeError, OSError):
                continue
            lines = source.splitlines()
            mod_doc = (ast.get_docstring(tree) or "").strip().split("\n\n")[0].replace("\n", " ")[:200]
            self.modules[rel] = mod_doc
            module_name = rel[:-3].replace("/", ".")
            for node in tree.body:
                if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                    self._add(module_name, "", node, "function", rel, lines, mod_doc)
                elif isinstance(node, ast.ClassDef):
                    cls = Symbol(f"{module_name}.{node.name}", "class", rel, node.lineno, node.end_lineno or node.lineno, f"class {node.name}",
                                 ast.get_docstring(node) or "", module_doc=mod_doc)
                    cls.markers = _markers(lines, node)[:6]
                    self.symbols[cls.qname] = cls
                    for sub in node.body:
                        if isinstance(sub, (ast.FunctionDef, ast.AsyncFunctionDef)):
                            self._add(module_name, node.name, sub, "method", rel, lines, mod_doc)
        self._link(tests_dirs)

    def _add(self, module: str, cls: str, node: ast.AST, kind: str, rel: str, lines: List[str], mod_doc: str) -> None:
        name = node.name  # type: ignore[attr-defined]
        qname = f"{module}.{cls + '.' if cls else ''}{name}"
        body = "\n".join(lines[node.lineno - 1:node.end_lineno])  # type: ignore[attr-defined]
        sym = Symbol(qname, kind, rel, node.lineno, node.end_lineno or node.lineno, _signature(node),  # type: ignore[attr-defined]
                     ast.get_docstring(node) or "", calls=_called_names(node), module_doc=mod_doc)
        sym.signals = [s for s, rx in _SIGNALS.items() if rx.search(body)]
        sym.markers = _markers(lines, node)
        self.symbols[qname] = sym

    def _link(self, tests_dirs: List[Path]) -> None:
        by_simple: Dict[str, List[str]] = defaultdict(list)
        for q, s in self.symbols.items():
            by_simple[s.simple].append(q)
        for q, s in self.symbols.items():
            s.ambiguous = len(by_simple[s.simple]) > 1
        callers: Dict[str, Set[str]] = defaultdict(set)
        for q, s in self.symbols.items():
            for name in s.calls:
                for target in by_simple.get(name, []):
                    if target != q:
                        callers[target].add(q)
        for q, s in self.symbols.items():
            s.callers = sorted(callers.get(q, set()))
        tests: Dict[str, str] = {}
        for d in tests_dirs:
            if d.is_dir():
                for p in sorted(d.rglob("test_*.py")):
                    try:
                        tests[p.relative_to(self.root).as_posix()] = p.read_text(encoding="utf-8", errors="ignore")
                    except (OSError, ValueError):
                        continue
        for q, s in self.symbols.items():
            if len(s.simple) < 5 and "_" not in s.simple:
                continue                                           # too generic to attribute a test to by name
            rx = re.compile(r"\b" + re.escape(s.simple) + r"\b")
            s.tests = [t for t, text in tests.items() if rx.search(text)][:6]

    # -- searching ---------------------------------------------------------------------------------------------------------

    @staticmethod
    def tokens(text: str) -> List[str]:
        text = re.sub(r"([a-z])([A-Z])", r"\1 \2", text or "")
        out: List[str] = []
        for word in re.findall(r"[A-Za-z0-9]+", text.lower()):
            if word not in _STOP and len(word) > 1:
                out.append(word)
        for run in re.findall(r"[一-鿿]+", text):
            out += [run] if len(run) == 1 else [run[i:i + 2] for i in range(len(run) - 1)]
        return out

    def find(self, query: str, k: int = 8) -> List[Dict[str, Any]]:
        q = set(self.tokens(query))
        if not q:
            return []
        scored: List[Tuple[float, Symbol, List[str]]] = []
        for s in self.symbols.values():
            name_t = set(self.tokens(s.simple)) | set(self.tokens(s.qname.rsplit(".", 2)[0]))
            doc_t = set(self.tokens(s.doc))
            mark_t = set(self.tokens(" ".join(s.markers)))
            path_t = set(self.tokens(s.path)) | set(self.tokens(s.module_doc))
            call_t = set(self.tokens(" ".join(s.calls)))
            hit, score = [], 0.0
            for t in q:
                w = (3.0 if t in name_t else 0) + (2.0 if t in doc_t else 0) + (1.0 if t in mark_t else 0) + (1.0 if t in path_t else 0) + (0.5 if t in call_t else 0)
                if w:
                    hit.append(t)
                    score += w
            if score:
                score = score * (0.5 + len(hit) / len(q)) + (0.5 if s.tests else 0) + min(len(s.callers), 5) * 0.1 - (0.3 if s.kind == "class" else 0)
                scored.append((score, s, hit))
        scored.sort(key=lambda t: (-t[0], t[1].qname))
        return [self.card(s, matched=hit) for _, s, hit in scored[:max(1, min(int(k), 20))]]

    def card(self, s: Symbol, matched: Optional[List[str]] = None) -> Dict[str, Any]:
        return {"name": s.qname, "where": f"{s.path}:{s.line}", "signature": s.signature, "summary": s.summary(), "lines": s.lines,
                "signals": s.signals, "markers": len(s.markers), "tested_by": s.tests[:3], "callers": len(s.callers),
                "ambiguous_name": s.ambiguous, "reuse": s.reuse(), "matched": sorted(matched or [])}

    # -- one symbol --------------------------------------------------------------------------------------------------------

    def resolve(self, name: str) -> List[Symbol]:
        name = str(name or "").strip()
        if name in self.symbols:
            return [self.symbols[name]]
        exact = [s for q, s in self.symbols.items() if q.endswith("." + name)]
        return exact

    def explain(self, name: str, source_lines: int = 0) -> Dict[str, Any]:
        found = self.resolve(name)
        if not found:
            return {"found": False, "message": f"no function, method or class named {name!r}; try find_capability with words describing it"}
        if len(found) > 1:
            return {"found": False, "ambiguous": [self.card(s) for s in found[:8]], "message": "several match; use the full name"}
        s = found[0]
        by_simple: Dict[str, List[str]] = defaultdict(list)
        for q, sym in self.symbols.items():
            by_simple[sym.simple].append(q)
        callees = sorted({t for n in s.calls for t in by_simple.get(n, []) if t != s.qname})
        out = dict(self.card(s), found=True, doc=(s.doc or "")[:1200], markers_text=s.markers, called_by=s.callers[:10], calls=callees[:15],
                   module=s.path, module_summary=s.module_doc,
                   note="callers, callees and tests are matched by name only: a name defined in several places lists all of them")
        if source_lines:
            lines = (self.root / s.path).read_text(encoding="utf-8").splitlines()[s.line - 1:s.end]
            out["source"] = "\n".join(lines[:max(1, min(int(source_lines), 120))])
            out["source_truncated"] = len(lines) > int(source_lines)
        return out

    def overview(self, package: str = "") -> Dict[str, Any]:
        prefix = package.strip("/").replace(".", "/")
        mods = [(p, d) for p, d in sorted(self.modules.items()) if not prefix or p.startswith(prefix)]
        counts: Dict[str, int] = defaultdict(int)
        for s in self.symbols.values():
            counts[s.path] += 1
        return {"package": package or "(all)", "modules": [{"file": p, "symbols": counts.get(p, 0), "about": d[:140]} for p, d in mods[:60]],
                "omitted": max(0, len(mods) - 60)}


_MAP: Optional[CodeMap] = None


def get_map() -> CodeMap:
    global _MAP
    if _MAP is None:
        _MAP = CodeMap()
    return _MAP
