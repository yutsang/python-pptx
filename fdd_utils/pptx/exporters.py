from __future__ import annotations

# re-added: bound by an import in another section of the pre-split module
from typing import Any, Dict, Iterable, List, Optional
import re


import copy
import logging
import os
import posixpath
import time
import traceback
from typing import Dict, List, Optional

from pptx import Presentation
from pptx.oxml.ns import qn

logger = logging.getLogger(__name__)


class ReportGenerator:
    """Report generator that orchestrates PPTX creation from markdown."""

    def __init__(
        self,
        template_path: str,
        markdown_file: str,
        output_path: str,
        project_name: Optional[str] = None,
        language: str = "english",
        row_limit: int = 20,
    ):
        self.template_path = template_path
        self.markdown_file = markdown_file
        self.output_path = output_path
        self.project_name = project_name
        self.language = language
        self.row_limit = row_limit

    def generate(self):
        from .generation import PowerPointGenerator  # local: exporters imports the later generation; module-level would be a cycle
        logger.info("Starting PPTX generation...")
        logger.info("Template: %s", self.template_path)
        logger.info("Markdown: %s", self.markdown_file)
        logger.info("Output: %s", self.output_path)
        logger.info("Language: %s", self.language)
        logger.info("Project: %s", self.project_name)

        with open(self.markdown_file, "r", encoding="utf-8") as handle:
            md_content = handle.read()

        logger.info("Content length: %s characters", len(md_content))
        generator = PowerPointGenerator(self.template_path, self.language, self.row_limit)

        try:
            generator.generate_full_report(md_content, None, self.output_path)
            if self.project_name:
                generator.update_project_titles(self.project_name, "BS")
        except Exception as exc:
            logger.error("Report generation failed: %s", exc)
            raise

        logger.info("PPTX generation completed: %s", self.output_path)


def export_pptx(
    template_path: str,
    markdown_path: str,
    output_path: str,
    project_name: Optional[str] = None,
    _excel_file_path: Optional[str] = None,
    language: str = "english",
    statement_type: str = "BS",
    row_limit: int = 20,
    model_type: Optional[str] = None,
):
    from .generation import PowerPointGenerator  # local: exporters imports the later generation; module-level would be a cycle
    generator = ReportGenerator(template_path, markdown_path, output_path, project_name, language, row_limit)
    generator.generate()

    if not os.path.exists(output_path):
        raise FileNotFoundError(f"PPTX file was not created at {output_path}")

    if project_name:
        temp_presentation = Presentation(output_path)
        pptx_gen = PowerPointGenerator(template_path, language, row_limit, model_type=model_type)
        pptx_gen.presentation = temp_presentation
        pptx_gen.update_project_titles(project_name, statement_type)
        temp_presentation.save(output_path)

    logger.info("PowerPoint presentation successfully exported to: %s", output_path)
    return output_path


def export_pptx_from_structured_data_combined(
    template_path: str,
    bs_data: List[Dict],
    is_data: List[Dict],
    output_path: str,
    project_name: Optional[str] = None,
    language: str = "english",
    temp_path: Optional[str] = None,
    selected_sheet: Optional[str] = None,
    is_chinese_databook: bool = False,
    bs_is_results: Optional[Dict[str, Any]] = None,
    model_type: Optional[str] = None,
    model_name: Optional[str] = None,
    skip_summary_ai: bool = False,  # AI summary needed for coSummaryShape; parallelized at max_workers=4
    pre_generated_summaries: Optional[Dict[str, str]] = None,  # {"BS": str, "IS": str} — bypass AI in PPTX export
    mappings: Optional[Dict[str, Any]] = None,  # for translating the embedded BS/IS table's row labels when Chinese
    financials_from: Optional[str] = None,  # workbook the Financials SHEET lives in, when not temp_path
):
    from .generation import PowerPointGenerator  # local: exporters imports the later generation; module-level would be a cycle
    try:
        export_started_at = time.perf_counter()
        def _stage_log(msg: str) -> None:
            logger.info(msg)

        _stage_log(f"Starting export | BS={len(bs_data)} IS={len(is_data)} skip_summary_ai={skip_summary_ai}")

        generator = PowerPointGenerator(template_path, language, row_limit=20, model_type=model_type, model_name=model_name)
        if skip_summary_ai:
            generator.pptx_settings.setdefault("executive_summary", {})["enable_ai"] = False
        stage_started_at = time.perf_counter()
        generator.load_template()
        _stage_log(f"load_template: {time.perf_counter() - stage_started_at:.2f}s")

        # Rule on every subtable in the deck before either statement is
        # planned -- the cap is per DECK and the income statement outranks the
        # balance sheet, neither of which can be honoured from inside the
        # per-statement planner that runs BS first. See _select_deck_subtables.
        from .helpers import _select_deck_subtables
        for rejected_item, reason in _select_deck_subtables(
            [("BS", bs_data), ("IS", is_data)], generator.pptx_settings,
        ):
            _stage_log(
                f"Subtable not drawn for "
                f"{rejected_item.get('mapping_key') or rejected_item.get('account_name')}: {reason}"
            )

        pre_summaries = pre_generated_summaries or {}
        if bs_data:
            stage_started_at = time.perf_counter()
            generator.apply_structured_data_to_slides(
                bs_data, 1, project_name, "BS",
                is_chinese_databook=is_chinese_databook,
                pre_generated_summary=pre_summaries.get("BS"),
            )
            _stage_log(f"apply_bs_slides: {time.perf_counter() - stage_started_at:.2f}s")
        if is_data:
            stage_started_at = time.perf_counter()
            generator.apply_structured_data_to_slides(
                is_data, 5, project_name, "IS",
                is_chinese_databook=is_chinese_databook,
                pre_generated_summary=pre_summaries.get("IS"),
            )
            _stage_log(f"apply_is_slides: {time.perf_counter() - stage_started_at:.2f}s")
        # bs_is_results being already-computed is sufficient on its own --
        # requiring selected_sheet too silently skipped the embedded table
        # whenever the caller had no sheet name to give (roll-up-sourced
        # financials with a blank own-file sheet, or a synthesized BS/IS
        # built purely from schedule tabs with no Financials sheet at all)
        # even though there was real BS/IS data ready to embed.
        if temp_path and (selected_sheet or bs_is_results):
            stage_started_at = time.perf_counter()
            generator.embed_financial_tables(
                temp_path,
                selected_sheet,
                project_name,
                language,
                bs_is_results=bs_is_results,
                mappings=mappings,
                financials_path=financials_from,
            )
            _stage_log(f"embed_financial_tables: {time.perf_counter() - stage_started_at:.2f}s")
        if hasattr(generator, "_unused_slides_to_remove") and generator._unused_slides_to_remove:
            stage_started_at = time.perf_counter()
            unused_slides_sorted = sorted(set(generator._unused_slides_to_remove), reverse=True)
            generator._remove_slides(unused_slides_sorted)
            _stage_log(f"remove_unused_slides ({len(unused_slides_sorted)}): {time.perf_counter() - stage_started_at:.2f}s")
        if project_name:
            stage_started_at = time.perf_counter()
            generator.refresh_project_placeholders(project_name)
            _stage_log(f"refresh_project_placeholders: {time.perf_counter() - stage_started_at:.2f}s")

        stage_started_at = time.perf_counter()
        generator.save(output_path)
        _stage_log(f"save_presentation: {time.perf_counter() - stage_started_at:.2f}s")
        _stage_log(f"TOTAL export: {time.perf_counter() - export_started_at:.2f}s")
        logger.info("Combined PPTX generation completed: %s", output_path)
        return output_path
    except Exception as exc:
        logger.error("PPTX generation failed: %s", exc)
        logger.error(traceback.format_exc())
        raise


def export_pptx_from_structured_data(
    template_path: str,
    structured_data: List[Dict],
    output_path: str,
    project_name: Optional[str] = None,
    language: str = "english",
    statement_type: str = "BS",
    start_slide: int = 1,
    model_type: Optional[str] = None,
):
    from .generation import PowerPointGenerator  # local: exporters imports the later generation; module-level would be a cycle
    try:
        logger.info("Starting PPTX generation from structured data...")
        logger.info("Template: %s", template_path)
        logger.info("Output: %s", output_path)
        logger.info("Language: %s", language)
        logger.info("Statement type: %s, Start slide: %s", statement_type, start_slide)
        logger.info("Accounts to process: %s", len(structured_data))

        generator = PowerPointGenerator(template_path, language, row_limit=20, model_type=model_type)
        generator.load_template()
        generator.apply_structured_data_to_slides(structured_data, start_slide, project_name, statement_type)
        generator.save(output_path)

        logger.info("PPTX generation completed: %s", output_path)
        return output_path
    except Exception as exc:
        logger.error("PPTX generation failed: %s", exc)
        raise


def merge_presentations(bs_presentation_path: str, is_presentation_path: str, output_path: str):
    try:
        logger.info("🔄 Starting presentation merge...")
        logger.info("   BS: %s", bs_presentation_path)
        logger.info("   IS: %s", is_presentation_path)

        merged_prs = Presentation(bs_presentation_path)
        is_prs = Presentation(is_presentation_path)

        from copy import deepcopy

        for slide_idx, slide in enumerate(is_prs.slides):
            try:
                slide_layout = slide.slide_layout
                new_slide = merged_prs.slides.add_slide(slide_layout)

                source_slide_xml = slide._element
                target_slide_xml = new_slide._element

                shapes_to_remove = list(new_slide.shapes)
                for shape in shapes_to_remove:
                    try:
                        sp_tree = target_slide_xml.get_or_add_spTree()
                        sp_tree.remove(shape._element)
                    except Exception:
                        pass

                source_sp_tree = source_slide_xml.get_or_add_spTree()
                target_sp_tree = target_slide_xml.get_or_add_spTree()
                for shape_element in source_sp_tree:
                    target_sp_tree.append(deepcopy(shape_element))

            except Exception as exc:
                logger.error("Error copying slide %s, using fallback method: %s", slide_idx, exc)
                slide_layout = slide.slide_layout
                new_slide = merged_prs.slides.add_slide(slide_layout)
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for new_shape in new_slide.shapes:
                            if (
                                hasattr(new_shape, "name")
                                and hasattr(shape, "name")
                                and new_shape.name == shape.name
                                and new_shape.has_text_frame
                            ):
                                new_shape.text_frame.text = shape.text_frame.text
                                break

        merged_prs.save(output_path)
        del merged_prs
        del is_prs

        import gc

        gc.collect()
        logger.info("✅ Presentation merge completed successfully")
    except Exception as exc:
        logger.error("Presentation merge failed: %s", exc)
        raise


def _dedupe_part_name(dest_prs: "Presentation", target_part, renamed_part_ids: set) -> None:
    """Rename `target_part` in-place if its partname collides with a part
    already present in dest_prs's package.

    python-pptx's Package.save() writes every part reachable from the
    package's own relationship graph using each Part object's OWN
    `.partname` -- it never re-derives a name. When _copy_slide_into()
    relates a destination slide directly to a Part object still owned by a
    DIFFERENT source Presentation (e.g. a picture's blipFill target), that
    part keeps the partname it was assigned in ITS OWN package (e.g.
    "/ppt/media/image3.png"). Since every batch entity's deck is built by
    the same export code, two different source decks landing on the same
    numbered partname is common, not a corner case -- and when that
    happens, the combined package ends up with two different parts both
    claiming "/ppt/media/image3.png", which produces a zip with a
    duplicate member name: invalid OPC, which is exactly what makes
    PowerPoint prompt "repair this presentation" (the media is
    unrecoverable/misattributed, not merely cosmetically wrong).
    Renaming the incoming part to a partname that's actually free in the
    destination package's namespace (via next_partname, the same
    mechanism python-pptx itself uses when adding new parts) avoids the
    collision. Only checked once per distinct source Part object
    (tracked by id() in `renamed_part_ids`, shared across an entire
    combine_presentations() call) -- once resolved, a part's identity/
    partname pairing is stable for the rest of the run.
    """
    if id(target_part) in renamed_part_ids:
        return
    renamed_part_ids.add(id(target_part))
    existing_partnames = {p.partname for p in dest_prs.part.package.iter_parts()}
    if target_part.partname not in existing_partnames:
        return
    partname = target_part.partname
    name_part = re.sub(r"\d+$", "", posixpath.splitext(partname.filename)[0]) or "part"
    tmpl = posixpath.join(partname.baseURI, f"{name_part}%d.{partname.ext}") if partname.ext else posixpath.join(partname.baseURI, f"{name_part}%d")
    target_part.partname = dest_prs.part.package.next_partname(tmpl)


def _copy_slide_into(dest_prs: "Presentation", source_slide, renamed_part_ids: Optional[set] = None) -> None:
    """Deep-copy one slide from a DIFFERENT Presentation (built from the
    same template.pptx) onto the end of dest_prs, preserving every shape
    including images and native tables.

    python-pptx has no built-in "append an existing slide" API, so this
    clones the slide's shape-tree XML directly -- the same technique
    merge_presentations() above uses. The one thing that technique is
    missing (and why it's not reused as-is here): the copied XML still
    references relationship IDs (r:embed / r:id / r:link, used by
    pictures and hyperlinks) that only exist in the SOURCE file's part.
    Left unmapped, those would point at nothing in the destination part --
    copied images would come through as silently broken/missing rather
    than raising an error. Every non-slideLayout relationship the source
    slide owns is re-created on the destination slide's own part first,
    and every r:embed/r:id/r:link attribute in the copied XML is
    rewritten to the new relationship id.

    Embedded/linked OLE objects (MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT /
    LINKED_OLE_OBJECT -- e.g. a "TCLayout.ActiveDocument.1" marker some
    add-ins like ThinkCell/UpSlide leave on every slide) are deliberately
    SKIPPED entirely, not copied or relationship-remapped: a real batch
    combine produced blank/whited-out pages specifically where these
    existed, and this codebase has no template with such an object to
    debug the exact OLE relationship mechanics against locally. These
    markers are consistently 0.001in x 0.001in (invisible, carry no
    reader-facing content) in every template seen so far, so dropping them
    trades an add-in bookkeeping artifact for guaranteed-correct visible
    content -- the safer side of that tradeoff.
    """
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    renamed_part_ids = renamed_part_ids if renamed_part_ids is not None else set()

    layout_name = source_slide.slide_layout.name
    dest_layout = next(
        (layout for layout in dest_prs.slide_layouts if layout.name == layout_name),
        dest_prs.slide_layouts[0],
    )
    dest_slide = dest_prs.slides.add_slide(dest_layout)

    # The layout auto-populates placeholder shapes -- clear them, the
    # source slide's own shape tree (copied below) already carries
    # everything that should be on the page.
    for shape in list(dest_slide.shapes):
        shape._element.getparent().remove(shape._element)

    r_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    ole_shape_elements = set()
    ole_rel_ids = set()
    for shape in source_slide.shapes:
        try:
            is_ole = shape.shape_type in (MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT, MSO_SHAPE_TYPE.LINKED_OLE_OBJECT)
        except (ValueError, NotImplementedError):
            is_ole = False
        if is_ole:
            ole_shape_elements.add(shape._element)
            for el in shape._element.iter():
                for attr_name in ("embed", "link", "id"):
                    rid = el.get(f"{r_ns}{attr_name}")
                    if rid:
                        ole_rel_ids.add(rid)

    rel_id_map: Dict[str, str] = {}
    for rel_id, rel in source_slide.part.rels.items():
        if rel.reltype.endswith("/slideLayout") or rel_id in ole_rel_ids:
            continue  # layout relationship isn't copied; OLE ones are deliberately dropped
        if rel.is_external:
            new_rel_id = dest_slide.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        else:
            _dedupe_part_name(dest_prs, rel.target_part, renamed_part_ids)
            new_rel_id = dest_slide.part.relate_to(rel.target_part, rel.reltype)
        rel_id_map[rel_id] = new_rel_id

    for shape_elm in list(source_slide.shapes._spTree):
        if shape_elm.tag in (qn("p:nvGrpSpPr"), qn("p:grpSpPr")):
            continue  # spTree's two fixed non-shape children, not content
        if shape_elm in ole_shape_elements:
            continue
        new_elm = copy.deepcopy(shape_elm)
        for el in new_elm.iter():
            for attr_name in ("embed", "link", "id"):
                old_rid = el.get(f"{r_ns}{attr_name}")
                if old_rid and old_rid in rel_id_map:
                    el.set(f"{r_ns}{attr_name}", rel_id_map[old_rid])
        dest_slide.shapes._spTree.append(new_elm)


def combine_presentations(pptx_sources: List, output_path) -> "str | None":
    """Combine several already-exported .pptx decks (e.g. one per batch
    entity, all built from the same template.pptx) into a single deck --
    every slide from every source, in order, via _copy_slide_into().

    pptx_sources: file paths (str) and/or file-like objects (e.g.
    io.BytesIO of already-in-memory PPTX bytes -- python-pptx's own
    Presentation() constructor accepts either, so no temp files are needed
    when combining straight from a batch run's cached pptx_download_data).
    output_path: a path (str) to save to, OR a file-like object (e.g.
    io.BytesIO) to write into instead of touching disk -- returns the path
    string in the former case, None in the latter (caller already holds
    the buffer it passed in).

    Deliberately NOT a general-purpose "merge any two PPTX files" utility:
    it assumes every input shares the same template (true for every batch
    entity, since they all come from export_pptx_from_structured_data_combined
    with the same template_path), which is what makes layout-name matching
    a safe way to pick the destination layout for each copied slide.
    """
    if not pptx_sources:
        raise ValueError("combine_presentations requires at least one input source")

    combined_prs = Presentation(pptx_sources[0])
    renamed_part_ids: set = set()
    for source in pptx_sources[1:]:
        source_prs = Presentation(source)
        for source_slide in source_prs.slides:
            _copy_slide_into(combined_prs, source_slide, renamed_part_ids)

    combined_prs.save(output_path)
    logger.info("Combined %s presentation(s)", len(pptx_sources))
    if isinstance(output_path, str):
        return output_path
    return None
# --- end pptx/exporters.py ---
