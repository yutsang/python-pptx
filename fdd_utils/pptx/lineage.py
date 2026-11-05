"""Which paragraph of which slide says what, from which facts -- and what a returned deck has done to it.

A deck leaves the tool with a manifest. Each commentary paragraph is an ELEMENT:
the account it belongs to, the slide and shape it sits in, a fingerprint of its
text, and -- when it was written from claims -- the semantic ids of the facts it
cites with a hash of each fact's value. An account that continues across slides
is several elements; a slide that holds many accounts is many.

Two halves, kept apart on purpose:

  PUBLIC    embedded in the .pptx as a custom XML part, and as alt-text tags on the
            shapes. Carries no figure: ids, fingerprints, value HASHES. A deck is
            sent to people the working papers are not (strip_manifest removes it).
  INTERNAL  a sidecar <deck>.lineage.json that stays with the working files: the
            baseline fact values, the paragraph texts, the claims. It is what turns
            "something changed" into "this changed from X to Y" and what makes a
            sentence-level three-way merge possible.

What each half buys, stated rather than implied:
  both present        edits are told from source changes, a conflict gets a concrete
                      proposal, history of the source is known
  public only         edits and source changes are still detected (fingerprints and
                      hashes), but not WHAT they were; no merge proposal
  neither             nothing is known about history. The deck can be checked against
                      the current numbers; it cannot be said what changed since.

A returned deck is read back, never trusted: shapes copied, slides reordered, a
paragraph rewritten, a tag lost, the whole part stripped by another tool. Whatever
cannot be matched with enough confidence is reported for a person to relink; it is
not edited.
"""
from __future__ import annotations

import difflib
import hashlib
import json
import os
import re
import uuid
from dataclasses import dataclass, field
from typing import Any, Dict, Iterable, List, Optional, Tuple

from pptx import Presentation
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI

MANIFEST_VERSION = 1
CUSTOM_XML_RELTYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
MANIFEST_PARTNAME = "/customXml/fdd_manifest.xml"
MANIFEST_NS = "urn:fdd-tool:manifest:1"
TAG_PREFIX = "fdd:"
INTERNAL_SUFFIX = ".lineage.json"
MANAGED_SHAPE = re.compile(r"^textMainBullets")
BULLET = "■"
CONT_MARKS = ("(续)", "(cont'd)", "（续）")
SEPARATOR = re.compile(r"^\s*[-–—:：]\s*")


def norm(text: str) -> str:
    return re.sub(r"\s+", "", str(text or ""))


def fingerprint(text: str) -> str:
    return hashlib.sha1(norm(text).encode("utf-8")).hexdigest()[:12]


def value_hash(value: Any) -> str:
    try:
        return hashlib.sha1(f"{float(value):.2f}".encode()).hexdigest()[:10]
    except (TypeError, ValueError):
        return hashlib.sha1(str(value).encode()).hexdigest()[:10]


# ---------------------------------------------------------------------------
# Reading paragraphs
# ---------------------------------------------------------------------------

@dataclass
class ParaRef:
    slide: int
    shape_id: int
    shape_name: str
    index: int                 # paragraph index inside the shape
    text: str
    label: Optional[str] = None
    body: str = ""
    continuation: bool = False
    tags: List[str] = field(default_factory=list)       # element ids the shape's alt text names
    shape_pos: int = 0                                   # position in the slide's shape tree (a copy can repeat shape_id)

    def to_dict(self) -> Dict[str, Any]:
        return dict(self.__dict__)


def shape_tags(shape: Any) -> Tuple[Optional[str], List[str]]:
    """(deck id, element ids) from a shape's alt text, or (None, [])."""
    try:
        descr = shape._element.xpath(".//p:cNvPr")[0].get("descr") or ""
    except Exception:
        return None, []
    m = re.search(re.escape(TAG_PREFIX) + r"([\w-]+);([\w,]+)", descr)
    return (m.group(1), [x for x in m.group(2).split(",") if x]) if m else (None, [])


def parse_label(text: str, labels: Iterable[str]) -> Tuple[Optional[str], str, bool]:
    """(label, body, is_continuation) for a '■ <label> - <body>' paragraph."""
    stripped = text.strip()
    if not stripped.startswith(BULLET):
        return None, stripped, False
    rest = stripped[len(BULLET):].strip()
    for label in sorted(set(labels), key=len, reverse=True):
        if rest.startswith(label):
            tail = rest[len(label):]
            cont = False
            for mark in CONT_MARKS:
                if tail.lstrip().startswith(mark):
                    tail, cont = tail.lstrip()[len(mark):], True
                    break
            return label, SEPARATOR.sub("", tail, count=1).strip(), cont
    m = re.match(r"(.+?)(?:\s*(?:\(续\)|\(cont'd\)|（续）))?\s+[-–—]\s+(.*)$", rest, re.S)
    if m:
        return m.group(1).strip(), m.group(2).strip(), "续" in rest[: m.end(1) + 8] or "cont'd" in rest[: m.end(1) + 12]
    return None, rest, False


def read_paragraphs(deck: Any, labels: Optional[Iterable[str]] = None) -> List[ParaRef]:
    """Every paragraph of every commentary box, in reading order (slide, then box, then paragraph)."""
    prs = deck if hasattr(deck, "slides") else Presentation(deck)
    out: List[ParaRef] = []
    known = list(labels or [])
    for number, slide in enumerate(prs.slides, 1):
        everything = list(slide.shapes)        # one pass: proxy objects are rebuilt on each iteration
        managed = [(i, s) for i, s in enumerate(everything) if getattr(s, "has_text_frame", False) and MANAGED_SHAPE.match(s.name or "")]
        managed.sort(key=lambda item: (0 if (item[1].name or "").endswith("_L") or not (item[1].name or "").endswith("_R") else 1, item[1].top or 0))
        for position, shape in managed:
            _deck_id, tags = shape_tags(shape)
            for i, para in enumerate(shape.text_frame.paragraphs):
                text = para.text
                if not text.strip():
                    continue
                label, body, cont = parse_label(text, known)
                out.append(ParaRef(number, shape.shape_id, shape.name, i, text, label, body or text.strip(), cont, tags, position))
    return out


# ---------------------------------------------------------------------------
# Tables (the embedded statement tables): fingerprinted so a hand edit is seen, never rewritten here
# ---------------------------------------------------------------------------

def table_fingerprint(table: Any) -> str:
    cells = [norm(table.cell(r, c).text) for r in range(len(table.rows)) for c in range(len(table.columns))]
    return hashlib.sha1("|".join(cells).encode("utf-8")).hexdigest()[:12]


def read_tables(deck: Any) -> List[Dict[str, Any]]:
    prs = deck if hasattr(deck, "slides") else Presentation(deck)
    out: List[Dict[str, Any]] = []
    for number, slide in enumerate(prs.slides, 1):
        for position, shape in enumerate(list(slide.shapes)):
            if getattr(shape, "has_table", False):
                _deck_id, tags = shape_tags(shape)
                table = shape.table
                out.append({"slide": number, "shape_id": shape.shape_id, "shape_pos": position, "tags": tags,
                            "title": table.cell(0, 0).text.strip()[:60], "rows": len(table.rows), "cols": len(table.columns),
                            "fingerprint": table_fingerprint(table)})
    return out


def locate_tables(manifest_tables: List[Dict[str, Any]], current: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    """Each exported table, found again by tag, else by title and size when that is unique."""
    result: Dict[str, Dict[str, Any]] = {}
    claimed: set = set()
    for t in manifest_tables:
        pool = [c for c in current if (c["slide"], c["shape_pos"]) not in claimed]
        tagged = [c for c in pool if t["id"] in c["tags"]]
        if len(tagged) > 1:
            home = [c for c in tagged if c["slide"] == t["slide"]]
            tagged = home if len(home) == 1 else []
            if not tagged:
                result[t["id"]] = {"status": "ambiguous", "reason": "the table was copied; no copy is clearly the original"}
                continue
        found = tagged or [c for c in pool if c["title"] == t["title"] and (c["rows"], c["cols"]) == (t["rows"], t["cols"])]
        if len(found) != 1:
            result[t["id"]] = {"status": "unlocated" if not found else "ambiguous",
                               "reason": "no table with this tag or title and size" if not found else "several tables match by title and size"}
            continue
        c = found[0]
        claimed.add((c["slide"], c["shape_pos"]))
        result[t["id"]] = {"status": "unchanged" if c["fingerprint"] == t["fingerprint"] else "user_edited",
                           "slide_now": c["slide"], "moved": c["slide"] != t["slide"]}
    return result


# ---------------------------------------------------------------------------
# What a claim depends on besides the values it quotes
# ---------------------------------------------------------------------------

PERMISSION_FIELDS = ("describe_total", "describe_difference", "explain_with_detail", "detail_basis",
                     "state_nil_or_first_arose", "compute_movement")


def _h(obj: Any) -> str:
    return hashlib.sha1(json.dumps(obj, sort_keys=True, ensure_ascii=False, default=str).encode("utf-8")).hexdigest()[:10]


def permission_view(contract: Any) -> Dict[str, Dict[str, Any]]:
    return {p: {k: (perm or {}).get(k) for k in PERMISSION_FIELDS} for p, perm in (contract.permissions or {}).items()}


def note_hashes(contract: Any, searched: bool = True) -> List[str]:
    """Hashes of the notes' text. A note a search filed on the contract (origin "search") is a working artefact of one
    session, not content of the workbook: a snapshot of the workbook leaves them out (searched=False)."""
    return sorted(_h(n["text"]) for n in contract.notes if searched or n.get("origin") != "search")


def dependency_digest(contract: Any, cite_ids: Iterable[str], cause_note_hashes: Iterable[str]) -> Dict[str, Any]:
    """What a paragraph rests on other than the quoted values, as hashes: what each cited period permits,
    how each cited period's checks came out, and the text of the notes a cause cites. A figure can stay
    the same while the reason it may be stated, or the note that explained it, does not."""
    periods = sorted({f["period"] for f in list(contract.facts) + list(contract.derived)
                      if f["id"] in set(cite_ids) and f.get("period")} |
                     {p for f in contract.derived if f["id"] in set(cite_ids) for p in (f.get("periods") or [])})
    permissions, matrix = permission_view(contract), contract.status_matrix()
    return {"permissions": _h({p: permissions.get(p) for p in periods}),
            "checks": _h({p: matrix.get(p) for p in periods}),
            "notes": sorted(set(cause_note_hashes))}


def element_dependencies(contract: Any, claims: List[Dict[str, Any]], cite_ids: Iterable[str]) -> Optional[Dict[str, Any]]:
    if contract is None or not claims:
        return None
    hashes = []
    for claim in claims:
        for ref in claim.get("cause_cites") or []:
            note = contract.note(ref)
            if note is not None:
                hashes.append(_h(note["text"]))
    return dependency_digest(contract, cite_ids, hashes)


def dependency_changes(contract: Any, deps: Dict[str, Any], cite_ids: Iterable[str]) -> List[str]:
    """Why a stored dependency digest no longer matches the sources, in words."""
    now = dependency_digest(contract, cite_ids, [])
    why = []
    if now["permissions"] != deps.get("permissions"):
        why.append("what the cited periods permit changed")
    if now["checks"] != deps.get("checks"):
        why.append("the checks on the cited periods came out differently")
    current_notes = set(note_hashes(contract))
    gone = [h for h in deps.get("notes") or [] if h not in current_notes]
    if gone:
        why.append(f"{len(gone)} note(s) the cause rested on changed or no longer exist")
    return why


# ---------------------------------------------------------------------------
# Building and stamping a manifest
# ---------------------------------------------------------------------------

def build_manifest(
    deck_path: str,
    items: List[Dict[str, Any]],
    *,
    is_chinese: bool,
    source_version: str,
    contracts: Optional[Dict[str, Any]] = None,
    claims_by_account: Optional[Dict[str, List[Dict[str, Any]]]] = None,
    binding: Optional[Dict[str, Any]] = None,
) -> Tuple[Dict[str, Any], Dict[str, Any]]:
    """(public, internal) for an exported deck. `items` are the payload items the
    exporter was given (mapping_key, display names, commentary)."""
    from .helpers import _rendered_bullet_label
    # The account is the key the run and the contracts use (account_name, the tab); mapping_key is the
    # canonical mapping name and differs for some accounts ('Other current assets' -> 'Other CA').
    def key_of(item: Dict[str, Any]) -> str:
        return str(item.get("account_name") or item.get("mapping_key") or "")
    labels = {_rendered_bullet_label(item, is_chinese): key_of(item) for item in items}
    texts = {key_of(item): str(item.get("commentary") or "") for item in items}
    paragraphs = read_paragraphs(deck_path, labels)
    deck_id = uuid.uuid4().hex[:10]
    elements: List[Dict[str, Any]] = []
    internal_elements: Dict[str, Dict[str, Any]] = {}
    seen_parts: Dict[str, int] = {}
    last_account: Optional[str] = None
    for para in paragraphs:
        account = labels.get(para.label or "")
        if account is None and para.label is None and last_account and norm(para.body) and norm(para.body) in norm(texts.get(last_account, "")):
            account = last_account                               # an unlabelled continuation line of the account above
        if account is None:
            last_account = None
            continue
        last_account = account
        seen_parts[account] = seen_parts.get(account, 0) + 1
        eid = f"E{len(elements) + 1}"
        claims = [c for c in (claims_by_account or {}).get(account, []) if norm(c.get("text", "")) and norm(c["text"]) in norm(para.body)]
        contract = (contracts or {}).get(account)
        cites, hashes = [], {}
        for claim in claims:
            for ref in list(claim.get("cites") or []):
                fact = contract.fact(ref) if contract is not None else None
                if fact is not None and fact["id"] not in hashes and fact.get("value") is not None:
                    cites.append(fact["id"])
                    hashes[fact["id"]] = value_hash(fact["value"])
        elements.append({
            "id": eid, "kind": "commentary", "account": account, "label": para.label, "slide": para.slide,
            "shape_id": para.shape_id, "shape_name": para.shape_name, "paragraph": para.index,
            "part": seen_parts[account], "continuation": para.continuation, "fingerprint": fingerprint(para.body),
            "cites": cites, "cite_hashes": hashes, "claim_ids": [c.get("id") for c in claims if c.get("id")],
            "deps": element_dependencies(contract, claims, cites)})
        internal_elements[eid] = {"text": para.body, "claims": claims}
    for e in elements:
        e["parts"] = seen_parts.get(e["account"], 1)
    account_digest = {}
    for account, contract in (contracts or {}).items():
        totals = sorted((f["period"], f["basis"], round(f["value"], 2)) for f in contract.facts if f["role"] == "total")
        account_digest[account] = hashlib.sha1(json.dumps(totals).encode()).hexdigest()[:12]
    tables = [{"id": f"T{i}", "slide": t["slide"], "shape_id": t["shape_id"], "title": t["title"], "rows": t["rows"], "cols": t["cols"],
               "fingerprint": t["fingerprint"]} for i, t in enumerate(read_tables(deck_path), 1)]
    public = {"version": MANIFEST_VERSION, "deck_id": deck_id, "source_version": source_version, "elements": elements,
              "tables": tables, "account_digest": account_digest}
    baseline = {}
    for account, contract in (contracts or {}).items():
        baseline[account] = {"facts": {f["id"]: f["value"] for f in list(contract.facts) + list(contract.derived) if f.get("value") is not None},
                             "checks": {f"{c['period']}|{c['kind']}": c["status"] for c in contract.checks},
                             "periods": list(contract.periods), "permissions": permission_view(contract),
                             "notes": note_hashes(contract)}
    internal = {"version": MANIFEST_VERSION, "deck_id": deck_id, "binding": binding or {}, "elements": internal_elements,
                "baseline": baseline, "texts": texts}
    return public, internal


def stamp_deck(deck_path: str, public: Dict[str, Any], internal: Optional[Dict[str, Any]] = None) -> None:
    """Tag the shapes, embed the public manifest, write the internal sidecar."""
    prs = Presentation(deck_path)
    by_shape: Dict[Tuple[int, int], List[str]] = {}
    for e in list(public["elements"]) + list(public.get("tables") or []):
        by_shape.setdefault((e["slide"], e["shape_id"]), []).append(e["id"])
    for number, slide in enumerate(prs.slides, 1):
        for shape in slide.shapes:
            ids = by_shape.get((number, shape.shape_id))
            if ids:
                shape._element.xpath(".//p:cNvPr")[0].set("descr", f"{TAG_PREFIX}{public['deck_id']};{','.join(ids)}")
    _set_part(prs, public)
    prs.save(deck_path)
    if internal is not None:
        with open(deck_path + INTERNAL_SUFFIX, "w", encoding="utf-8") as fh:
            json.dump(internal, fh, ensure_ascii=False, indent=1)


def _set_part(prs: Any, public: Dict[str, Any]) -> None:
    for rel in list(prs.part.rels.values()):
        if rel.reltype == CUSTOM_XML_RELTYPE and not rel.is_external and str(rel.target_part.partname) == MANIFEST_PARTNAME:
            prs.part.drop_rel(rel.rId)
    xml = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<m xmlns="{MANIFEST_NS}"><![CDATA['
           + json.dumps(public, ensure_ascii=False) + "]]></m>").encode("utf-8")
    part = Part(PackURI(MANIFEST_PARTNAME), "application/xml", prs.part.package, xml)
    prs.part.relate_to(part, CUSTOM_XML_RELTYPE)


def read_manifest(deck_path: str) -> Optional[Dict[str, Any]]:
    """The embedded public manifest, or None (never made, or stripped by another tool)."""
    try:
        prs = Presentation(deck_path)
        for rel in prs.part.rels.values():
            if rel.reltype == CUSTOM_XML_RELTYPE and not rel.is_external and str(rel.target_part.partname) == MANIFEST_PARTNAME:
                m = re.search(r"<!\[CDATA\[(.*)\]\]>", rel.target_part.blob.decode("utf-8"), re.S)
                return json.loads(m.group(1)) if m else None
    except Exception:
        return None
    return None


def read_internal(deck_path: str) -> Optional[Dict[str, Any]]:
    path = deck_path + INTERNAL_SUFFIX
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def strip_manifest(deck_path: str, out_path: Optional[str] = None) -> str:
    """The deck without the embedded manifest or its alt-text tags: the copy that goes outside."""
    prs = Presentation(deck_path)
    for rel in list(prs.part.rels.values()):
        if rel.reltype == CUSTOM_XML_RELTYPE and not rel.is_external and str(rel.target_part.partname) == MANIFEST_PARTNAME:
            prs.part.drop_rel(rel.rId)
    for slide in prs.slides:
        for shape in slide.shapes:
            cnv = shape._element.xpath(".//p:cNvPr")
            if cnv and (cnv[0].get("descr") or "").startswith(TAG_PREFIX):
                cnv[0].attrib.pop("descr")
    target = out_path or deck_path
    prs.save(target)
    return target


# ---------------------------------------------------------------------------
# Reading a returned deck
# ---------------------------------------------------------------------------

def sentences(text: str) -> List[str]:
    parts = re.split(r"(?<=[。；;！!？?])|(?<=\.)\s+", str(text or ""))
    return [p.strip() for p in parts if p and p.strip()]


def three_way(base: Optional[str], user: str, new: str) -> Dict[str, Any]:
    """Sentence-level merge of what the user left and what the sources now say, against what was exported.

    A sentence only the sources changed takes the new text; one only the user changed
    keeps theirs; one both changed differently is a conflict and keeps the user's
    until a person decides. With no base text nothing can be told apart, and every
    difference is a conflict."""
    user_s, new_s = sentences(user), sentences(new)
    if base is None:
        return {"mode": "no_base", "merged": user, "conflicts": [{"user": u, "new": n} for u, n in zip(user_s, new_s) if norm(u) != norm(n)]
                or ([{"user": user, "new": new}] if norm(user) != norm(new) else [])}
    base_s = sentences(base)
    merged: List[str] = []
    conflicts: List[Dict[str, str]] = []
    used_new: set = set()
    for b in base_s:
        u = next((x for x in user_s if difflib.SequenceMatcher(None, norm(x), norm(b)).ratio() > 0.85), None)
        user_changed = u is None or norm(u) != norm(b)
        n = next(((i, x) for i, x in enumerate(new_s) if difflib.SequenceMatcher(None, norm(x), norm(b)).ratio() > 0.55), None)
        if n:
            used_new.add(n[0])
        new_text = n[1] if n else None
        source_changed = new_text is not None and norm(new_text) != norm(b)
        if user_changed and source_changed:
            if u is not None and new_text is not None and norm(u) == norm(new_text):
                merged.append(u)
            else:
                conflicts.append({"base": b, "user": u or "(deleted)", "new": new_text})
                if u:
                    merged.append(u)
        elif user_changed:
            if u:
                merged.append(u)
        elif source_changed:
            merged.append(new_text)
        else:
            merged.append(b)
    merged.extend(x for x in user_s if not any(difflib.SequenceMatcher(None, norm(x), norm(b)).ratio() > 0.85 for b in base_s))
    merged.extend(x for i, x in enumerate(new_s) if i not in used_new and norm(x) not in {norm(m) for m in merged})
    return {"mode": "three_way", "merged": "".join(merged), "conflicts": conflicts}


def locate(public: Dict[str, Any], internal: Optional[Dict[str, Any]], paragraphs: List[ParaRef]) -> Dict[str, Any]:
    """Match each manifest element to a paragraph of the returned deck.

    Order of trust: the shape's alt-text tag with the account's label; then the same
    label and the same text anywhere (a slide moved); then, if the baseline text is
    known, the most similar paragraph under that label, taken only when it clearly
    beats the runner-up. Anything else is reported, not guessed."""
    deck_id = public.get("deck_id")
    elements = public["elements"]
    by_label: Dict[str, List[ParaRef]] = {}
    for p in paragraphs:
        by_label.setdefault(p.label or "", []).append(p)
    claimed: set = set()
    result: Dict[str, Dict[str, Any]] = {}
    duplicates: List[Dict[str, Any]] = []
    shapes_with_tag: Dict[str, set] = {}
    for p in paragraphs:
        for t in p.tags:
            shapes_with_tag.setdefault(t, set()).add((p.slide, p.shape_pos))
    for eid, shapes in shapes_with_tag.items():
        if len(shapes) > 1:
            duplicates.append({"element": eid, "shapes": sorted(shapes)})
    dup_ids = {d["element"] for d in duplicates}
    for e in elements:
        label = e.get("label") or ""
        pool = [p for p in by_label.get(label, []) if id(p) not in claimed]
        base_text = ((internal or {}).get("elements") or {}).get(e["id"], {}).get("text")
        tagged = [p for p in pool if e["id"] in p.tags]
        method, candidates = "", []
        if tagged:
            # copies carry the same tag: the one still where it was exported is the original
            at_home = [p for p in tagged if p.slide == e["slide"]]
            candidates, method = (at_home or tagged), "tag"
            if e["id"] in dup_ids and len(at_home) != 1:
                result[e["id"]] = {"status": "ambiguous", "reason": "the shape was copied; no copy is clearly the original",
                                   "candidates": [(p.slide, p.shape_pos) for p in tagged]}
                continue
            if len(candidates) > 1:
                candidates = [p for p in candidates if p.index == e["paragraph"]] or candidates[:1]
        if not candidates:
            same = [p for p in pool if fingerprint(p.body) == e["fingerprint"]]
            if len(same) == 1:
                candidates, method = same, "text"
            elif len(same) > 1:
                result[e["id"]] = {"status": "ambiguous", "reason": "the same text appears more than once under this label",
                                   "candidates": [(p.slide, p.shape_id) for p in same]}
                continue
        if not candidates and base_text:
            scored = sorted(((difflib.SequenceMatcher(None, norm(p.body), norm(base_text)).ratio(), p) for p in pool),
                            key=lambda x: -x[0])
            if scored and scored[0][0] >= 0.6 and (len(scored) == 1 or scored[0][0] - scored[1][0] >= 0.1):
                candidates, method = [scored[0][1]], "similarity"
            elif scored and scored[0][0] >= 0.6:
                result[e["id"]] = {"status": "ambiguous", "reason": "two paragraphs are about equally like the original",
                                   "candidates": [(p.slide, p.shape_id) for _, p in scored[:2]]}
                continue
        if not candidates:
            result[e["id"]] = {"status": "unlocated", "reason": "no paragraph under this label matches by tag, text or similarity"}
            continue
        p = candidates[0]
        claimed.add(id(p))
        edited = fingerprint(p.body) != e["fingerprint"]
        moved = p.slide != e["slide"]
        result[e["id"]] = {"status": "user_edited" if edited else ("moved" if moved else "unchanged"), "method": method,
                           "slide_now": p.slide, "shape_id": p.shape_id, "paragraph": p.index, "text_now": p.body, "moved": moved}
    unclaimed = [p for p in paragraphs if id(p) not in claimed and p.label]
    return {"elements": result, "added_paragraphs": [p.to_dict() for p in unclaimed], "duplicate_tags": duplicates}


def restamp(deck_path: str, session: Any, applied: List[Dict[str, Any]], texts: Dict[str, str],
            claims_by_element: Dict[str, List[Dict[str, Any]]], fully_reconciled: bool) -> None:
    """After an in-place update: advance the manifest for what was changed, and only that.

    An element updated gets its new fingerprint and citations. The source version and the
    baseline of an account move forward only when every element of the deck is reconciled with
    the sources; otherwise the paragraphs still waiting (a conflict, a relink) would look current."""
    public = read_manifest(deck_path)
    if public is None:
        return
    internal = read_internal(deck_path) or {"version": MANIFEST_VERSION, "deck_id": public["deck_id"], "elements": {}, "baseline": {}}
    by_id = {e["id"]: e for e in public["elements"]}
    for item in applied:
        element = by_id[item["element"]]
        element["fingerprint"] = fingerprint(texts[element["id"]])
        contract = session.contracts.get(element["account"])
        claims = claims_by_element.get(element["id"], [])
        cites, hashes = [], {}
        for claim in claims:
            for ref in claim.get("cites") or []:
                fact = contract.fact(ref) if contract is not None else None
                if fact is not None and fact.get("value") is not None and fact["id"] not in hashes:
                    cites.append(fact["id"])
                    hashes[fact["id"]] = value_hash(fact["value"])
        element["cites"], element["cite_hashes"] = cites, hashes
        element["deps"] = element_dependencies(contract, claims, cites)
        internal.setdefault("elements", {})[element["id"]] = {"text": texts[element["id"]], "claims": claims}
    if fully_reconciled:
        public["source_version"] = session.source_version
        internal["baseline"] = session.snapshot()
    stamp_deck(deck_path, public, internal)
